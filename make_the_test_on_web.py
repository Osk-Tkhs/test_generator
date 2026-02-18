import streamlit as st
import pandas as pd
import random
import io

st.title("📝 自動テスト作成ツール")

# 1. ファイルアップローダー
uploaded_file = st.file_uploader("問題データ(xlsx)を読み込んでください", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # データの中身を確認
    st.write("### 元データプレビュー (最初の5件)")
    st.dataframe(df.head())

    st.divider()
    st.sidebar.header("テスト生成設定")

    # 2. ユーザーに入力してもらう設定
    # 通し番号の最小値と最大値を自動取得
    min_no = int(df["通し番号"].min())
    max_no = int(df["通し番号"].max())

    start_num = st.sidebar.number_input("開始番号", min_no, max_no, min_no)
    end_num = st.sidebar.number_input("終了番号", start_num, max_no, max_no)
    
    # 選択可能な最大問題数を計算
    available_range_df = df[(df["通し番号"] >= start_num) & (df["通し番号"] <= end_num)]
    max_questions = len(available_range_df)
    
    count = st.sidebar.number_input(f"問題数 (最大 {max_questions}件)", 1, max_questions, min(10, max_questions))

    # 3. テスト生成ボタン

    if st.button("テストを生成する"):
        # ランダム抽出 (重複なし)
        test_df = available_range_df.sample(n=count).sort_values("通し番号")
    
        st.write("### 生成されたテストプレビュー")
        st.dataframe(test_df)

        # --- Excelファイルを作成する処理 ---
        # メモリ上にバイナリデータを保存するためのバッファを作成
        output = io.BytesIO()
    
        # PandasのExcelWriterを使用して、xlsxwriterエンジンで書き込み
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            test_df.to_excel(writer, index=False, sheet_name='TestSheet')
    
        # ポインタを先頭に戻す
        processed_data = output.getvalue()

        # ダウンロードボタンの設置
        st.download_button(
            label="📥 テストをExcel(.xlsx)で保存",
            data=processed_data,
            file_name=f"test_{start_num}_to_{end_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
