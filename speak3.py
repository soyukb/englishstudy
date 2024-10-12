import pandas as pd
from gtts import gTTS
import os
import tkinter as tk
from tkinter import filedialog

# tkinterを用いてエクセルファイルを選択する関数
def select_excel_file():
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを表示しない
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],  # Excelファイルのみ表示
        title="エクセルファイルを選択してください"
    )
    return file_path

# エクセルファイルをエクスプローラーから選択
excel_file = select_excel_file()

# エクセルファイルが選択されなかった場合は終了
if not excel_file:
    print("ファイルが選択されませんでした。プログラムを終了します。")
    exit()

# エクセルファイルがあるディレクトリを取得
excel_dir = os.path.dirname(excel_file)

# 英語と日本語の保存先フォルダをエクセルファイルのディレクトリ内に作成
output_folder_en = os.path.join(excel_dir, 'english_audio_files')
output_folder_jp = os.path.join(excel_dir, 'japanese_audio_files')

# 英語の保存先フォルダが存在しない場合は作成
if not os.path.exists(output_folder_en):
    os.makedirs(output_folder_en)

# 日本語の保存先フォルダが存在しない場合は作成
if not os.path.exists(output_folder_jp):
    os.makedirs(output_folder_jp)

# エクセルファイルを読み込み（最初のシートを自動選択）
df = pd.read_excel(excel_file, sheet_name=0)

# 各行の1列目（英語）と2列目（日本語）を読み込み、MP3ファイルを生成
for index, row in df.iterrows():
    # 1列目の英語の例文
    english_sentence = str(row[0])  # 1列目の値を取得し文字列に変換
    if english_sentence:  # もし値が存在する場合
        tts_en = gTTS(text=english_sentence, lang='en')
        file_name_en = f'{index + 1}.EN.mp3'
        file_path_en = os.path.join(output_folder_en, file_name_en)
        tts_en.save(file_path_en)
        print(f'{file_name_en} を英語フォルダに作成しました。')

    # 2列目の日本語の例文
    japanese_sentence = str(row[1])  # 2列目の値を取得し文字列に変換
    if japanese_sentence:  # もし値が存在する場合
        tts_jp = gTTS(text=japanese_sentence, lang='ja')
        file_name_jp = f'{index + 1}.JP.mp3'
        file_path_jp = os.path.join(output_folder_jp, file_name_jp)
        tts_jp.save(file_path_jp)
        print(f'{file_name_jp} を日本語フォルダに作成しました。')

print(f"全ての例文の音声ファイルが作成されました。")
print(f"英語音声ファイルは '{output_folder_en}' フォルダに保存されました。")
print(f"日本語音声ファイルは '{output_folder_jp}' フォルダに保存されました。")
