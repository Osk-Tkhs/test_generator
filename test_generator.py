import streamlit as st
import pandas as pd
import io

st.title("📝 汎用テスト作成ツール")

uploaded_file = st.file_uploader("問題データ(xlsx)を読み込んでください", type=["xlsx"])

if uploaded_file is not None:
    # header=0 で1行目を見出しとして読み込みますが、
    # その後の処理では列名ではなく「位置」を使います
    df = pd.read_excel(uploaded_file)
    
    # 列数チェック（最低3列あるか）
    if len(df.columns) < 3:
        st.error("Excelファイルには最低3つの列（通し番号、問題、解答）が必要です。")
    else:
        # 列名に関わらず、位置でリネームして扱いやすくする
        # 0番目: 通し番号, 1番目: 問題, 2番目: 解答
        col_names = df.columns
        df_working = df.copy()
        
        st.write(f"### 元データプレビュー")
        st.dataframe(df.head())

        st.sidebar.header("テスト生成設定")

        # 列の位置（0番目の列）を「通し番号」として数値を抽出
        # 数値以外のデータが混ざっている場合に備えてエラーハンドリング
        try:
            ids = pd.to_numeric(df.iloc[:, 0])
            min_no = int(ids.min())
            max_no = int(ids.max())
        except:
            st.error("1列目（通し番号）に数字以外のデータが含まれています。確認してください。")
            st.stop()

        start_num = st.sidebar.number_input("開始番号", min_no, max_no, min_no)
        end_num = st.sidebar.number_input("終了番号", start_num, max_no, max_no)
        
        # 指定範囲の行を抽出（1列目の値でフィルタリング）
        mask = (ids >= start_num) & (ids <= end_num)
        available_range_df = df[mask]
        max_questions = len(available_range_df)
        
        count = st.sidebar.number_input(f"問題数 (最大 {max_questions}件)", 1, max(1, max_questions), min(10, max_questions))

        if st.button("テストを生成する"):
            if max_questions == 0:
                st.error("指定された範囲に問題が見つかりません。")
            else:
                # 重複なしランダム抽出
                test_df = available_range_df.sample(n=count).sort_values(by=df.columns[0])
                
                st.success(f"範囲内から {count}問 を抽出しました。")
                st.dataframe(test_df)

                # Excel出力処理
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # 元の列名のまま出力します
                    test_df.to_excel(writer, index=False, sheet_name='TestSheet')
                
                processed_data = output.getvalue()

                st.download_button(
                    label="📥 テストをExcel(.xlsx)で保存",
                    data=processed_data,
                    file_name="generated_test.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
