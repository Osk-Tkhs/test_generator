import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Test Generator", layout="centered") # 中央寄せで読みやすく
st.title("📝 Test Generator for Excel")

# --- STEP 1: ファイル読み込み ---
uploaded_file = st.file_uploader("1. Excelファイルをアップロード", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # --- STEP 2: 設定入力 (横並びの入力欄) ---
    st.divider()
    st.subheader("2. 抽出条件の設定")
    
    # 3つの入力欄を横に並べる
    col1, col2, col3 = st.columns(3)
    
    min_val = int(df.iloc[:, 0].min())
    max_val = int(df.iloc[:, 0].max())

    with col1:
        start_num = st.number_input("開始番号", min_val, max_val, min_val)
    with col2:
        end_num = st.number_input("終了番号", start_num, max_val, max_val)
        
    # 範囲内のデータ数を計算
    mask = (df.iloc[:, 0] >= start_num) & (df.iloc[:, 0] <= end_num)
    filtered_df = df[mask]
    available_count = len(filtered_df)

    with col3:
        count = st.number_input(f"問題数 (最大:{available_count})", 1, max(1, available_count), min(10, available_count))

    # --- STEP 3: 生成実行 ---
    st.divider()
    
    # ボタンを中央寄せっぽく配置するために空の列を挟む
    _, btn_col, _ = st.columns([1, 2, 1])
    
    if btn_col.button("🚀 この条件でテストを生成する", use_container_width=True):
        if available_count == 0:
            st.warning("指定された範囲にデータがありません。")
        else:
            # ランダム抽出
            test_df = filtered_df.sample(n=count).sort_values(by=df.columns[0])
            
            st.success(f"抽出完了！ ({count}問)")
            st.dataframe(test_df, use_container_width=True)

            # Excel出力用バッファ
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                test_df.to_excel(writer, index=False, sheet_name='Test')
            
            # ダウンロードボタンを目立たせる
            st.download_button(
                label="📥 生成したExcelファイルを保存する",
                data=output.getvalue(),
                file_name=f"test_{start_num}-{end_num}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True # ボタンを横いっぱいに広げる
            )
