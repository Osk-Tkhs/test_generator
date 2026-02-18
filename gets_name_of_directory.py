from pathlib import Path

def create_subdirectory_list(target_directory_path):
    # パスオブジェクトの作成と正規化
    target_dir = Path(target_directory_path).resolve()

    # ディレクトリが存在するかチェック
    if not target_dir.exists() or not target_dir.is_dir():
        print(f"❌ エラー: 指定されたディレクトリが見つかりません: {target_dir}")
        return

    print(f"📂 探索中: {target_dir}")

    # 子ディレクトリの名前をリストに取得
    # .iterdir() で中身を走査し、.is_dir() でディレクトリのみを抽出
    sub_dirs = [item.name for item in target_dir.iterdir() if item.is_dir()]

    # 名前順にソート（任意）
    sub_dirs.sort()

    # 保存するテキストファイルのパスを設定（指定ディレクトリ内）
    output_file_path = target_dir / "folder_list.txt"

    try:
        # テキストファイルとして書き出し
        # utf-8エンコーディングを指定して文字化けを防止
        with output_file_path.open("w", encoding="utf-8") as f:
            if sub_dirs:
                f.write("\n".join(sub_dirs))
                print(f"✅ 子ディレクトリの一覧を保存しました: {output_file_path}")
                print(f"📁 取得件数: {len(sub_dirs)} 件")
            else:
                f.write("(子ディレクトリは見つかりませんでした)")
                print("⚠️ 子ディレクトリが見つかりませんでした。空のリストを作成しました。")
                
    except Exception as e:
        print(f"❌ ファイル書き込み中にエラーが発生しました: {e}")

# ==========================================
# 実行設定
# ==========================================
# リストを取得したいディレクトリのパスを指定してください
# 例: "C:\\Users\\fufuf\\Documents" など
TARGET_DIR = "C:\\Users\\fufuf\\OneDrive\\デスクトップ\\FILES_SLN" 

if __name__ == "__main__":
    create_subdirectory_list(TARGET_DIR)
