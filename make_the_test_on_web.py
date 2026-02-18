import streamlit as st
import pandas as pd

st.title("Excelデータ解析ツール")

# 1. ファイルアップローダーの設置
uploaded_file = st.file_uploader("xlsxファイルを選択してください", type=["xlsx"])

if uploaded_file is not None:
    # 2. pandasでExcelファイルを読み込む
    try:
        df = pd.read_excel(uploaded_file)
        
        st.success("ファイルの読み込みに成功しました！")
        
        # 3. データの表示（最初の5行）
        st.write("### プレビュー")
        st.dataframe(df.head())
        
        # 例: 特定の列の計算など
        # st.write(f"データ件数: {len(df)} 件")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
