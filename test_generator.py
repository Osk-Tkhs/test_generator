import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Test Generator", layout="centered") 
st.title("📝 Test Generator for Excel")

# --- ①：出題リスト(xlsx)の準備 ---
st.write("### ①：出題リスト(xlsx)の準備")

tab1, tab2 = st.tabs(["A: 新しく作成する", "B: 既存のファイルを使う"])

with tab1:
    st.info("これから作成する場合は、以下のひな型をダウンロードして入力してください。")
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        if os.path.exists("template.xlsx"):
            with open("template.xlsx", "rb") as f:
                st.download_button("📁 ひな型(空)をダウンロード", f, "template.xlsx", use_container_width=True)
    with col_dl2:
        if os.path.exists("sample_data.xlsx"):
            with open("sample_data.xlsx", "rb") as f:
                st.download_button("💡 見本(データ入)をダウンロード", f, "sample_data.xlsx", use_container_width=True)
    
    st.success("""
    **作成した出題リスト(xlsx)について、以下の2点をご確認ください：**
    - 1行目は「問題No」「問題」「解答」などの**見出し行**である
    - 2行目以降は 左端（A列）が **「半角数字」** で **「1～問題数」** の **「連番」** になっている（1, 2, 3...問題数）
    """)

with tab2:
    st.success("""
    **お手持ちの出題リスト(xlsx)について、以下の2点をご確認ください：**
    - 1行目は「問題No」「問題」「解答」などの**見出し行**である
    - 2行目以降は 左端（A列）が **「半角数字」** で **「数値（通し番号）」** が入っている
    """)

st.divider()

# --- ②：出題リスト(xlsx)のアップロード ---
st.write("### ②：出題リスト(xlsx)のアップロード")

uploaded_file = st.file_uploader("出題リスト(xlsx)をアップロードしてください", type=["xlsx"], accept_multiple_files=False)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        with st.expander("🔍 元データを確認する (先頭10件)"):
            st.dataframe(df.head(10), use_container_width=True)

        # 1列目の数値チェック
        first_col_raw = df.iloc[:, 0]
        first_col_numeric = pd.to_numeric(first_col_raw, errors='coerce')
        
        if first_col_numeric.isna().any():
            error_mask = first_col_numeric.isna()
            error_rows = df[error_mask].index + 2 
            st.error(f"⚠️ 1列目(問題No.)に数値以外のデータが含まれています。")
            st.warning(f"該当するExcel行番号: {list(error_rows[:10])} ...")
            st.info("【解決策】1列目の見出し以外をすべて「半角数字」に修正して、再度アップロードしてください。")
            st.stop()
        # --- 2. エラーハンドリング：連番チェック ---
        # 1から始まるべき理想の連番（1, 2, 3, ..., 行数）を作成
        expected_series = pd.Series(range(1, len(df) + 1))
        
        # 実際の1列目と理想の連番が一致するかチェック
        # (valuesを比較することで、インデックスの差を無視して中身をチェック)
        is_sequential = (first_col_numeric.values == expected_series.values).all()

        if not is_sequential:
            st.error("⚠️ 1列目(問題No.)が正しくありません。")
            
            # 詳細な原因を分析して表示
            if first_col_numeric.max() != len(df):
                st.warning(f"原因：最大番号({int(first_col_numeric.max())})と、実際のデータ行数({len(df)}行)が一致していません。")
            
            st.info("""
            **【修正ガイド】**
            - 1列目は必ず **「1」から始まる連番** にしてください。
            - 途中に欠番（例: 1, 2, 4...）や重複（例: 1, 2, 2...）がないか確認してください。
            - データの入っていない空行がExcelの下部に残っていないか確認してください。
            """)
            st.stop() # 処理を中断

        # --- 3. 空欄（NaN）チェックと場所の特定 ---
        # B列(1)とC列(2)を対象に空欄をチェック
        target_cols = df.iloc[:, 1:3]
        if target_cols.isna().any().any():
            st.error("⚠️ 問題、または解答の列に空欄がある行が見つかりました。")
            
            # 空欄がある行と列を特定してリスト化
            error_details = []
            for col_idx in [1, 2]: # B列とC列
                # その列の中でNaNがあるインデックスを取得
                nan_indices = df[df.iloc[:, col_idx].isna()].index
                if not nan_indices.empty:
                    col_name = df.columns[col_idx]
                    # Excelの行番号に変換（インデックス+2）
                    rows = [str(i + 2) for i in nan_indices]
                    error_details.append(f"・**{col_name}** 列の {', '.join(rows[:10])} 行目")

            # 具体的な場所を表示
            for detail in error_details:
                st.warning(detail)
            
            if len(rows) > 10:
                st.info("※該当箇所が多いため、先頭10件のみ表示しています。")
                
            st.info("すべての空欄を埋めてから、再度アップロードしてください。")
            st.stop()


        # --- ③：設定入力 ---
        st.divider()
        st.subheader("③：出題範囲、出題数の設定")
        
        col1, col2, col3 = st.columns(3)
        
        min_val = int(first_col_numeric.min())
        max_val = int(first_col_numeric.max())

        with col1:
            start_num = st.number_input("始点問題No.", min_val, max_val, min_val)
        with col2:
            end_num = st.number_input("終点問題No.", start_num, max_val, max_val)
            
        mask = (first_col_numeric >= start_num) & (first_col_numeric <= end_num)
        filtered_df = df[mask]
        available_count = len(filtered_df)

        with col3:
            count = st.number_input(f"問題数 (最大:{available_count})", 1, max(1, available_count), min(10, available_count))

        sort_option = st.radio(
            "問題の並び順を選んでください",
            ["昇順固定 (番号の小さい順)", "降順固定 (番号の大きい順)", "順番ランダム"],
            horizontal=True
        )

        # --- 生成実行 ---
        st.divider()
        _, btn_col, _ = st.columns([1, 2, 1])
        
        if btn_col.button("🚀 この条件でテストを生成する", use_container_width=True):
            if available_count == 0:
                st.warning("指定された範囲にデータがありません。番号設定を確認してください。")
            else:
                # 1. まずはランダムに必要数を抽出
                sampled_df = filtered_df.sample(n=count)

                # 2. 並び順設定に応じてソート処理
                if sort_option == "昇順固定 (番号の小さい順)":
                    test_df = sampled_df.sort_values(by=df.columns[0], ascending=True)
                elif sort_option == "降順固定 (番号の大きい順)":
                    test_df = sampled_df.sort_values(by=df.columns[0], ascending=False)
                else:
                    test_df = sampled_df
                
                st.success(f"抽出完了！ ({count}問)")
                st.dataframe(test_df, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    test_df.to_excel(writer, index=False, sheet_name='Test')
                
                st.download_button(
                    label="📥 生成したExcelファイルを保存する",
                    data=output.getvalue(),
                    file_name=f"test_{start_num}-{end_num}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True 
                )

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
else:
    st.info("上の枠にExcelファイルをドラッグ＆ドロップしてください。")

