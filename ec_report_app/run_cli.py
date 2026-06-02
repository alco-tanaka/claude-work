"""
EC 昨対比較分析 CLI（Claude から実行用）
data/raw/昨対比較分析/ 内の xlsx を自動検出して PPTX を生成する。
"""
import os, sys
from datetime import datetime
sys.path.insert(0, os.path.dirname(__file__))
from core import load, compute_stats, make_pos, make_int

RAW = os.path.join(os.path.dirname(__file__), '..', '..', 'data', 'raw', '昨対比較分析')
OUT = os.path.join(os.path.dirname(__file__), '..', '..', 'data', 'outputs')

def detect_files(folder):
    """xlsx を見つけ、データ内の年度順でソートして (前年, 今年) を返す。"""
    files = sorted(
        [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.xlsx')],
        key=lambda p: os.path.basename(p)
    )
    if len(files) < 2:
        raise ValueError(f'昨対比較分析フォルダに xlsx が {len(files)} 件しかありません（2件必要）')
    if len(files) > 2:
        print(f'注意: {len(files)} 件見つかりました。ファイル名の昇順で最初の2件を使用します。')
    return files[0], files[1]

def main():
    print('=== EC 昨対比較分析レポート生成 ===')
    raw_dir = os.path.abspath(RAW)
    print(f'フォルダ: {raw_dir}')

    f25_path, f26_path = detect_files(raw_dir)
    print(f'前年: {os.path.basename(f25_path)}')
    print(f'今年: {os.path.basename(f26_path)}')

    print('データ読み込み中...')
    d25 = load(f25_path)
    d26 = load(f26_path)

    print('集計・分析中...')
    s = compute_stats(d25, d26)
    print(f'ブランド: {s.brand}  /  期間: {s.year25}年 vs {s.year26}年 {s.period}')
    print(f'全データ 前年比 数量{s.TOT_UNIT_A:+.1f}%  売上高{s.TOT_REV_A:+.1f}%')

    today = datetime.today().strftime('%Y%m%d')
    fname = f'{s.brand}_EC昨対分析_{s.year25}vs{s.year26}_{today}'

    print('展示会用生成中...')
    pos_buf = make_pos(s)
    out1 = os.path.join(OUT, f'{fname}_展示会用.pptx')
    with open(out1, 'wb') as f:
        f.write(pos_buf.getvalue())
    print(f'OK: {out1}')

    print('社内版生成中...')
    int_buf = make_int(s)
    out2 = os.path.join(OUT, f'{fname}_社内版.pptx')
    with open(out2, 'wb') as f:
        f.write(int_buf.getvalue())
    print(f'OK: {out2}')

    print('Done')

if __name__ == '__main__':
    main()
