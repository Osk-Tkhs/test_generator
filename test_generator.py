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
        # Excel読み込み
        df_raw = pd.read_excel(uploaded_file)

        # --- ① 全範囲に対する処理 (スペースのみのセル & 書式のみセルの対策) ---
        def clear_pure_spaces(x):
            if isinstance(x, str):
                # 前後の空白を消して、中身が何もなくなれば None にする
                cleaned = x.strip().replace('　', '')
                return None if cleaned == '' else x
            return x

        # 全セルに対し「スペースのみ」なら空にする処理を適用
        df_raw = df_raw.applymap(clear_pure_spaces)

        # --- ② A列(通し番号)特化の処理 (数値内のスペースも削除) ---
        def remove_all_spaces(x):
            if isinstance(x, str):
                # A列は文字の中にあるスペースもすべて排除
                return x.strip().replace(' ', '').replace('　', '')
            return x

        df_raw.iloc[:, 0] = df_raw.iloc[:, 0].apply(remove_all_spaces)

        # --- ③ 有効範囲の特定 ---
        # A列に有効な値がある「一番下の行」を特定 (書式のみの空セルはここで排除される)
        last_idx = df_raw.iloc[:, 0].dropna().index.max()
        
        if pd.isna(last_idx):
            st.error("1列目に数値を入力してください。")
            st.stop()
            
        # A列の末尾までを「有効データ」として切り出す
        df = df_raw.loc[:last_idx].copy()

        # --- 以降、数値チェック・連番チェック・空欄チェック ---
        # (ここからは、A列が数値か、連番か、B・C列に空欄がないかをチェックするロジック)

        # --- ① 1列目の数値・形式チェック ---
        # A列を数値変換（数値化できないものはNaNにする）
        first_col_numeric = pd.to_numeric(df.iloc[:, 0], errors='coerce')
        
        if first_col_numeric.isna().any():
            # A列に文字が混じっている行を特定
            error_rows = df[first_col_numeric.isna()].index + 2
            st.error(f"⚠️ 1列目（問題No.）に数値以外のデータが含まれています。")
            st.warning(f"該当するExcel行番号: {list(error_rows)}")
            st.stop()

        # --- ② 1列目の連番チェック (1〜Nになっているか) ---
        # 期待される連番 [1, 2, 3, ..., 行数]
        expected_series = pd.Series(range(1, len(df) + 1))
        
        if not (first_col_numeric.values == expected_series.values).all():
            st.error("⚠️ 1列目が「1からの連番」になっていません。")
            st.info(f"期待される最終番号: {len(df)} (現在の最大: {int(first_col_numeric.max())})")
            st.warning("途中に欠番、重複、または1から始まっていない箇所があります。")
            st.stop()

        # --- ③ 2列目(問題)・3列目(解答)の空欄チェック ---
        # 1列目に番号がある行の中で、B列(1)かC列(2)が空の場所を特定
        error_details = []
        for col_idx in [1, 2]:
            nan_mask = df.iloc[:, col_idx].isna()
            if nan_mask.any():
                col_name = df.columns[col_idx]
                # Excelの行番号（index + 2）を取得
                nan_rows = df[nan_mask].index + 2
                rows_str = ", ".join([str(r) for r in nan_rows])
                error_details.append(f"・**{col_name}** 列の {rows_str} 行目")

        if error_details:
            st.error("⚠️ 問題、または解答に記入漏れ（空欄）があります。")
            for detail in error_details:
                st.warning(detail)
            st.info("1列目に番号がある行は、問題と解答をすべて埋める必要があります。")
            st.stop()

        # --- ここまで来ればデータは完璧 ---
        st.success(f"データチェック完了：{len(df)}件の問題を正しく読み込みました。")


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




