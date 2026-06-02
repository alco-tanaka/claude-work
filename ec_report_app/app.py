"""
EC 昨対比較分析レポート生成アプリ
Streamlit Cloud にデプロイして使用する
"""
import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from core import load, compute_stats, make_pos, make_int

st.set_page_config(page_title='EC 昨対比較分析', page_icon='📊', layout='wide')

# セッション状態の初期化
if 'pos_bytes' not in st.session_state:
    st.session_state.pos_bytes = None
if 'int_bytes' not in st.session_state:
    st.session_state.int_bytes = None
if 'fname' not in st.session_state:
    st.session_state.fname = ''

st.title('📊 EC 昨対比較分析レポート生成')
st.caption('前年・今年の販売データ（Excel）をアップロードしてPowerPointレポートを生成します')

col1, col2 = st.columns(2)
with col1:
    st.subheader('前年データ')
    f25 = st.file_uploader('前年 .xlsx をアップロード', type=['xlsx'], key='f25')
with col2:
    st.subheader('今年データ')
    f26 = st.file_uploader('今年 .xlsx をアップロード', type=['xlsx'], key='f26')

if f25 and f26:
    if st.button('レポートを生成する', type='primary'):
        try:
            with st.spinner('データ読み込み中...'):
                d25 = load(f25)
                d26 = load(f26)

            with st.spinner('集計・分析中...'):
                stats = compute_stats(d25, d26)

            with st.spinner('PowerPoint生成中（1〜2分かかります）...'):
                pos_buf = make_pos(stats)
                int_buf = make_int(stats)

            st.session_state.pos_bytes = pos_buf.getvalue()
            st.session_state.int_bytes = int_buf.getvalue()
            st.session_state.fname = (
                f'{stats.brand}_EC昨対分析'
                f'_{stats.year25}vs{stats.year26}'
                f'_{stats.MONTHS[0].replace("月","")}-{stats.MONTHS[-1]}'
            )

        except Exception as e:
            st.error(f'エラーが発生しました: {e}')
            st.exception(e)

else:
    st.info('前年・今年のデータファイルをアップロードしてください。')

if st.session_state.pos_bytes:
    st.success('生成完了！')
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label='📥 展示会用 PPTX をダウンロード',
            data=st.session_state.pos_bytes,
            file_name=f'{st.session_state.fname}_展示会用.pptx',
            mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        )
    with dl2:
        st.download_button(
            label='📥 社内版 PPTX をダウンロード',
            data=st.session_state.int_bytes,
            file_name=f'{st.session_state.fname}_社内版.pptx',
            mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        )
    with st.expander('データ形式について'):
        st.markdown("""
**必要な列（5行目ヘッダー）:**

| 列 | 内容 |
|---|---|
| 実績年月 / JAN / ブランドCD / ブランド名 / 商品CD / 商品名 | 基本情報 |
| カラーCD / カラー / サイズCD / サイズ | バリエーション |
| 請求先 / 得意先 / 担当者 | 取引情報 |
| 売上数 / 売上金額 | 実績数値 |
| 旧在庫区分 / 最新在庫区分 / 詳細分類 | 区分情報 |

**在庫区分D** = 区分Dアイテム（社内版のみに分析スライドを掲載）
        """)
