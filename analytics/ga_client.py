"""GA4 Data API クライアント"""
from __future__ import annotations
import os, json, calendar
from pathlib import Path
from datetime import date, timedelta
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import (
    DateRange, Dimension, Metric, RunReportRequest, OrderBy,
)
from google.oauth2 import service_account

ENV_PATH = Path(__file__).parent / ".env"

SCOPES = [
    "https://www.googleapis.com/auth/analytics.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]


def load_env() -> dict[str, str]:
    env = {}
    if ENV_PATH.exists():
        for line in ENV_PATH.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            env[k.strip()] = v.strip()
    for key in ["GOOGLE_CREDENTIALS_PATH", "GOOGLE_CREDENTIALS_JSON",
                "SLACK_WEBHOOK_TEST", "SLACK_WEBHOOK_PROD", "SLACK_BOT_TOKEN",
                "SPREADSHEET_ID"]:
        val = os.environ.get(key)
        if val:
            env[key] = val
    return env


def _get_credentials():
    env = load_env()
    creds_json_str = env.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json_str:
        info = json.loads(creds_json_str)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    creds_path = env.get("GOOGLE_CREDENTIALS_PATH",
                          str(Path(__file__).parent / "credentials.json"))
    return service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)


def get_ga_client() -> BetaAnalyticsDataClient:
    return BetaAnalyticsDataClient(credentials=_get_credentials())


def get_gspread_client():
    import gspread
    return gspread.authorize(_get_credentials())


def run_report(property_id: str, start_date: str, end_date: str,
               metrics: list[str], dimensions: list[str] | None = None,
               order_metric: str | None = None, limit: int = 100) -> list[dict]:
    client = get_ga_client()
    order_bys = []
    if order_metric:
        order_bys = [OrderBy(metric=OrderBy.MetricOrderBy(metric_name=order_metric), desc=True)]

    request = RunReportRequest(
        property=f"properties/{property_id}",
        date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
        metrics=[Metric(name=m) for m in metrics],
        dimensions=[Dimension(name=d) for d in (dimensions or [])],
        order_bys=order_bys,
        limit=limit,
    )
    response = get_ga_client().run_report(request)
    rows = []
    for row in response.rows:
        record = {}
        for i, dim in enumerate(dimensions or []):
            record[dim] = row.dimension_values[i].value
        for i, met in enumerate(metrics):
            record[met] = row.metric_values[i].value
        rows.append(record)
    return rows


def fetch_site_data(property_id: str, year: int, month: int) -> dict:
    """1サイト分の月次データを全部取得して返す"""
    last_day = calendar.monthrange(year, month)[1]
    start = date(year, month, 1).strftime("%Y-%m-%d")
    end = date(year, month, last_day).strftime("%Y-%m-%d")

    if month == 1:
        prev_year, prev_month = year - 1, 12
    else:
        prev_year, prev_month = year, month - 1
    prev_last_day = calendar.monthrange(prev_year, prev_month)[1]
    prev_start = date(prev_year, prev_month, 1).strftime("%Y-%m-%d")
    prev_end = date(prev_year, prev_month, prev_last_day).strftime("%Y-%m-%d")

    METRICS = ["sessions", "totalUsers", "newUsers", "screenPageViews",
               "averageSessionDuration", "engagementRate"]

    def to_float(v):
        try:
            return float(v)
        except Exception:
            return 0.0

    rows = run_report(property_id, start, end, METRICS)
    totals = {m: to_float(rows[0].get(m, 0)) for m in METRICS} if rows else {m: 0.0 for m in METRICS}

    prev_rows = run_report(property_id, prev_start, prev_end, METRICS)
    prev_totals = {m: to_float(prev_rows[0].get(m, 0)) for m in METRICS} if prev_rows else {m: 0.0 for m in METRICS}

    channel_rows = run_report(property_id, start, end,
                               ["sessions", "totalUsers"],
                               dimensions=["sessionDefaultChannelGroup"],
                               order_metric="sessions", limit=10)

    device_rows = run_report(property_id, start, end,
                              ["sessions"],
                              dimensions=["deviceCategory"],
                              order_metric="sessions")

    country_rows = run_report(property_id, start, end,
                               ["sessions"],
                               dimensions=["country"],
                               order_metric="sessions", limit=5)

    daily_rows = run_report(property_id, start, end,
                             ["sessions"],
                             dimensions=["date"],
                             limit=35)
    daily_rows.sort(key=lambda r: r.get("date", ""))

    page_rows = run_report(property_id, start, end,
                            ["screenPageViews", "sessions"],
                            dimensions=["pagePath"],
                            order_metric="sessions", limit=10)

    return {
        "totals": totals,
        "prev_totals": prev_totals,
        "channels": channel_rows,
        "devices": device_rows,
        "countries": country_rows,
        "daily": daily_rows,
        "pages": page_rows,
        "period": {"start": start, "end": end, "year": year, "month": month},
    }


def fetch_funnel_data(property_id: str, start_date: str, end_date: str) -> dict:
    """購買ファネルのイベント数を取得する。"""
    FUNNEL_EVENTS = ["view_item", "add_to_cart", "begin_checkout", "purchase"]
    rows = run_report(property_id, start_date, end_date,
                      ["eventCount"], ["eventName"], limit=500)
    counts = {r.get("eventName", ""): int(r.get("eventCount", 0)) for r in rows}
    return {e: counts.get(e, 0) for e in FUNNEL_EVENTS}
