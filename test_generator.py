import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Test Generator", layout="centered") 
st.title("📝 Test Generator for Excel")


# --- 冒頭の使い方のセクション ---
st.write("### 使い方")

col1, col2 = st.columns(2)


with col1:
    if os.path.exists("sample_data.xlsx"):
        with open("sample_data.xlsx", "rb") as f:
            st.download_button("💡 見本(データ入)をDL", f, "sample_data.xlsx")

with col2:
    if os.path.exists("template.xlsx"):
        with open("template.xlsx", "rb") as f:
            st.download_button("📁 雛形(空)をDL", f, "template.xlsx")


st.divider()



# --- STEP 1: ファイル読み込み ---
# accept_multiple_files=False（デフォルト）により、1つしか選択できません
uploaded_file = st.file_uploader("1. Excelファイルを1つアップロードしてください", type=["xlsx"], accept_multiple_files=False)

if uploaded_file is not None:
    try:
        # Excel読み込み
        df = pd.read_excel(uploaded_file)

        # 1. プレビュー機能
        with st.expander("🔍 元データを確認する (先頭10件)"):
            st.dataframe(df.head(10), use_container_width=True)

        # 2. エラーハンドリング：1列目の数値チェック
        # errors='coerce' で数値化できないものを NaN に変換
        first_col_raw = df.iloc[:, 0]
        first_col_numeric = pd.to_numeric(first_col_raw, errors='coerce')
        
        if first_col_numeric.isna().any():
            # エラーがある行（数値化に失敗した行）を特定
            error_mask = first_col_numeric.isna()
            error_rows = df[error_mask].index + 2 # Excelの行番号(見出し+1, indexは0からなので+1)
            
            st.error(f"⚠️ 1列目(問題No.)に数値以外のデータが含まれています。")
            st.warning(f"該当するExcel行番号: {list(error_rows[:10])} ...")
            st.info("【解決策】1列目の見出し以外をすべて「半角数字」に修正して、再度アップロードしてください。")
            st.stop() # 2項以降を表示させない

        # --- STEP 2: 設定入力 ---
        st.divider()
        st.subheader("2. 抽出条件の設定")
        
        col1, col2, col3 = st.columns(3)
        
        min_val = int(first_col_numeric.min())
        max_val = int(first_col_numeric.max())

        with col1:
            start_num = st.number_input("始点問題No.", min_val, max_val, min_val)
        with col2:
            end_num = st.number_input("終点問題No.", start_num, max_val, max_val)
            
        # 範囲内のデータ数を計算
        mask = (first_col_numeric >= start_num) & (first_col_numeric <= end_num)
        filtered_df = df[mask]
        available_count = len(filtered_df)

        with col3:
            count = st.number_input(f"問題数 (最大:{available_count})", 1, max(1, available_count), min(10, available_count))

        # 【追加】並び順の設定（ラジオボタンを横並びで配置）
        sort_option = st.radio(
            "問題の並び順を選んでください",
            ["昇順固定 (番号の小さい順)", "降順固定 (番号の大きい順)", "順番ランダム"],
            horizontal=True
        )

        # --- STEP 3: 生成実行 ---
        st.divider()
        
        # ボタンを中央に寄せるためのカラム設定
        _, btn_col, _ = st.columns([1, 2, 1])
        
        if btn_col.button("🚀 この条件でテストを生成する", use_container_width=True):
            if available_count == 0:
                st.warning("指定された範囲にデータがありません。番号設定を確認してください。")
            else:
                # ランダム抽出（並び順は1列目の昇順に固定）
                sampled_df = filtered_df.sample(n=count).sort_values(by=df.columns[0])

                # 2. 並び順設定に応じてソート処理
                if sort_option == "昇順固定 (番号の小さい順)":
                    test_df = sampled_df.sort_values(by=df.columns[0], ascending=True)
                elif sort_option == "降順固定 (番号の大きい順)":
                    test_df = sampled_df.sort_values(by=df.columns[0], ascending=False)
                else:
                    # 「順番ランダム」の場合はソートせず、抽出したまま（ランダムな順）にする
                    test_df = sampled_df
                
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
        st.error(f"ファイルの読み込み中にエラーが発生しました: {e}")
        st.info("ファイルが壊れているか、パスワード保護されている可能性があります。")

else:
    # ファイルがアップロードされていない時のガイド
    st.info("上の枠にExcelファイルをドラッグ＆ドロップしてください。")
