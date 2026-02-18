import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Test Generator", layout="centered") 
st.title("📝 Test Generator for Excel")

# --- STEP 1: ファイル読み込み ---
uploaded_file = st.file_uploader("1. Excelファイルをアップロード", type=["xlsx"])

if uploaded_file is not None:
    try:
        # header=0（1行目を見出し）として読み込み
        df = pd.read_excel(uploaded_file)

        # 【追加】プレビュー機能（エキスパンダーでスッキリ表示）
        with st.expander("🔍 元データを確認する (先頭50件)"):
            st.dataframe(df.head(50), use_container_width=True)

        # 【追加】エラーハンドリング：1列目の数値チェック
        # errors='coerce' で数値化できないものを NaN に変換
        first_col = pd.to_numeric(df.iloc[:, 0], errors='coerce')
        
        if first_col.isna().any():
            # エラーがある行を特定
            error_rows = df[first_col.isna()].index + 2 # Excelの行番号に合わせる(+2)
            st.error(f"⚠️ 1列目（通し番号）に数値以外のデータが含まれています。 (該当行: {list(error_rows[:5])} ...)")
            st.info("Excelの1列目は必ず半角数字のみにしてください。見出し行は自動で除外されます。")
            st.stop() # ここで処理を中断

        # --- STEP 2: 設定入力 ---
        st.divider()
        st.subheader("2. 抽出条件の設定")
        
        col1, col2, col3 = st.columns(3)
        
        min_val = int(first_col.min())
        max_val = int(first_col.max())

        with col1:
            start_num = st.number_input("開始番号", min_val, max_val, min_val)
        with col2:
            end_num = st.number_input("終了番号", start_num, max_val, max_val)
            
        # 範囲内のデータ数を計算
        mask = (first_col >= start_num) & (first_col <= end_num)
        filtered_df = df[mask]
        available_count = len(filtered_df)

        with col3:
            # max(1, ...) で0エラーを回避
            count = st.number_input(f"問題数 (最大:{available_count})", 1, max(1, available_count), min(10, available_count))

        # --- STEP 3: 生成実行 ---
        st.divider()
        
        _, btn_col, _ = st.columns([1, 2, 1])
        
        if btn_col.button("🚀 この条件でテストを生成する", use_container_width=True):
            if available_count == 0:
                st.warning("指定された範囲にデータがありません。番号設定を確認してください。")
            else:
                # ランダム抽出（1列目の名前を取得してソート）
                test_df = filtered_df.sample(n=count).sort_values(by=df.columns[0])
                
                st.success(f"抽出完了！ ({count}問)")
                st.dataframe(test_df, use_container_width=True)

                # Excel出力用バッファ
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
        st.error(f"予期せぬエラーが発生しました: {e}")

