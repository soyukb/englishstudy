from PIL import Image
import tempfile

# PIL.Image.ANTIALIAS のエラーを回避するための修正
if not hasattr(Image, 'ANTIALIAS'):
    Image.ANTIALIAS = Image.Resampling.LANCZOS  # ANTIALIAS の代わりに LANCZOS を設定

from moviepy.editor import AudioFileClip, VideoFileClip, concatenate_videoclips, ImageClip, TextClip, CompositeVideoClip, vfx
import os
import re
import tkinter as tk
from tkinter import filedialog
import pandas as pd

# tkinterを使用して、エクセルファイルを選択する関数を定義
def select_excel_file():
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを非表示にする
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],  # Excelファイルのみ選択可能
        title="エクセルファイルを選択してください"
    )
    return file_path

# エクセルファイルをエクスプローラーから選択
excel_file = select_excel_file()

# エクセルファイルが選択されなかった場合はプログラムを終了
if not excel_file:
    print("ファイルが選択されませんでした。プログラムを終了します。")
    exit()

# エクセルファイルが存在するディレクトリを取得
excel_dir = os.path.dirname(excel_file)

# 音声ファイルの保存先フォルダをエクセルファイルのディレクトリ内に作成
output_folder = os.path.join(excel_dir, 'english_audio_files')

# 保存先フォルダが存在しない場合は作成
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# エクセルファイルの最初のシートを読み込み
df = pd.read_excel(excel_file, sheet_name=0)

# ffmpeg のパスを設定し、環境変数に追加
ffmpeg_path = r"C:\ffmpeg\ffmpeg-master-latest-win64-gpl-shared\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"
print(f"Using ffmpeg path: {ffmpeg_path}")
os.environ["IMAGEIO_FFMPEG_EXE"] = ffmpeg_path

# 無音ファイル (Silence.mp3) のパスを指定
silence_path = r"C:\Users\soyuk\OneDrive\デスクトップ\YouTubeEnglish\Silence.mp3"
if not os.path.exists(silence_path):
    print(f"Error: 'Silence.mp3' not found in {output_folder}. Please add 'Silence.mp3' to the folder.")
    exit()

# 背景画像のパスをファイルダイアログで指定
background_image_path = filedialog.askopenfilename(
    title="背景画像を選択してください",
    filetypes=[("Image files", "*.jpg *.png *.jpeg")]  # 画像ファイルのみ選択可能
)

# 背景画像が選択されなかった場合はプログラムを終了
if not background_image_path:
    print("背景画像が選択されませんでした。プログラムを終了します。")
    exit()


# 背景画像のサイズを取得し、リサイズし、JPEG形式に変換
def resize_background_image(image_path, max_width=854, max_height=480):
    img = Image.open(image_path)
    
    # 画像がRGBAモードならRGBモードに変換
    if img.mode == 'RGBA':
        img = img.convert('RGB')

    img.thumbnail((max_width, max_height), Image.ANTIALIAS)  # 画像を指定サイズにリサイズ
    temp_img_path = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg").name
    img.save(temp_img_path, "JPEG")  # リサイズ後の画像を一時ファイルとして保存
    return temp_img_path

# 背景画像をリサイズして新しいファイルを使用
resized_background_image_path = resize_background_image(background_image_path)

# 音声ファイルが保存先フォルダにあるか確認し、処理
file_names = os.listdir(output_folder)
pattern = re.compile(r'(\d+)\.JP\.mp3')  # 日本語音声ファイル名のパターンを定義
index_numbers = [int(pattern.search(f).group(1)) for f in file_names if pattern.search(f)]  # 該当ファイルから番号を抽出
max_index = max(index_numbers) if index_numbers else 0  # 最大の番号を取得

# 有効な音声ファイルが見つからない場合は終了
if max_index == 0:
    print("Error: No valid audio files found in the specified folder.")
    exit()

# 一時ファイルのパスを格納するリスト
temp_files = []
for i in range(1, max_index + 1):
    jp_path = os.path.join(output_folder, f"{i}.JP.mp3")  # 日本語音声ファイルのパス
    en_path = os.path.join(output_folder, f"{i}.EN.mp3")  # 英語音声ファイルのパス

    # 日本語と英語の音声ファイルが存在するか確認
    if os.path.exists(jp_path) and os.path.exists(en_path):
        try:
            # 日本語、英語、無音のオーディオクリップを作成
            jp_clip = AudioFileClip(jp_path)
            en_clip = AudioFileClip(en_path)
            silence_clip = AudioFileClip(silence_path)

            print(f"Processing clip {i}:")
            print(f"  JP Clip Duration: {jp_clip.duration}")
            print(f"  EN Clip Duration: {en_clip.duration}")
            print(f"  Silence Clip Duration: {silence_clip.duration}")

            # 無音クリップの再生速度を2倍にする
            silence_2x_clip = silence_clip.fx(vfx.speedx, 2)
            
            # 15文字ごとに改行を挿入する関数
            def split_text_by_length(text, length=10):
                return '\n'.join([text[i:i+length] for i in range(0, len(text), length)])
            
            # 10単語ごとに改行を挿入する関数
            def split_text_by_words(text, num_words=10):
                words = text.split()  # テキストを単語ごとに分割
                return '\n'.join([' '.join(words[i:i+num_words]) for i in range(0, len(words), num_words)])

            # エクセルファイルから対応するテキストを取得
            en_text = str(df.iloc[i - 1, 0])
            jp_text = str(df.iloc[i - 1, 1])
            
            # 日本語テキストを15文字ごとに改行
            jp_text = split_text_by_length(jp_text, 15)
            
            # 英語テキストを10単語ごとに改行
            en_text = split_text_by_words(en_text, 10)
            
            # 使用するフォントパスを指定
            font_path = r"C:/Windows/Fonts/msgothic.ttc"

            # 日本語音声用のビデオクリップを作成（リサイズした背景画像とテキスト）
            jp_background = ImageClip(resized_background_image_path).set_duration(jp_clip.duration).resize((854, 480))
            jp_text_clip = TextClip(jp_text, fontsize=40, color='black', font=font_path).set_position('center').set_duration(jp_clip.duration)
            jp_video = CompositeVideoClip([jp_background, jp_text_clip.set_start(0)]).set_audio(jp_clip)

            # 無音音声用背景クリップ（背景画像を設定）
            jp_si_background = ImageClip(resized_background_image_path).set_duration(silence_clip.duration).resize((854, 480))
            jp_si_text_clip = TextClip(jp_text, fontsize=40, color='black', font=font_path).set_position('center').set_duration(silence_clip.duration)
            jp_si_video = CompositeVideoClip([jp_si_background, jp_si_text_clip.set_start(0)]).set_audio(silence_clip)
            
            # 無音音声用背景クリップ（背景画像を設定）
            en_si_background = ImageClip(resized_background_image_path).set_duration(silence_2x_clip.duration).resize((854, 480))
            en_si_text_clip = TextClip(en_text, fontsize=40, color='black', font=font_path).set_position('center').set_duration(silence_2x_clip.duration)
            en_si_video = CompositeVideoClip([en_si_background, en_si_text_clip.set_start(0)]).set_audio(silence_2x_clip)

            # 英語音声用のビデオクリップを作成（リサイズした背景画像とテキスト）
            en_background = ImageClip(resized_background_image_path).set_duration(en_clip.duration).resize((854, 480))
            en_text_clip = TextClip(en_text, fontsize=40, color='black', font=font_path).set_position('center').set_duration(en_clip.duration)
            en_video = CompositeVideoClip([en_background, en_text_clip.set_start(0)]).set_audio(en_clip)

            # 日本語と英語のビデオクリップを連結
            combined_clip = concatenate_videoclips([jp_video, jp_si_video, en_video, en_si_video, en_video, en_si_video], method="compose")

            # 一時ファイルとしてクリップを保存
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".mp4")
            combined_clip.write_videofile(temp_file.name, fps=24, codec="libx264", audio_codec="aac")
            temp_files.append(temp_file.name)
            temp_file.close()

        except Exception as e:
            print(f"Error processing audio and text for clip {i}: {e}")
            continue

# 一時ファイルを最終的に連結して MP4 ファイルとして書き出す処理
# （この部分は元のコードと同じ）
# 一時ファイルを最終的に連結して MP4 ファイルとして書き出す
if temp_files:
    try:
                # temp_files を 10 分の 1 に分割する関数
        def split_into_chunks(files, chunk_size):
            """
            temp_files リストを指定された chunk_size ごとに分割するジェネレータ関数。
            """
            for i in range(0, len(files), chunk_size):
                yield files[i:i + chunk_size]

        # ファイルを10個のグループに分けて処理
        chunk_size = max(1, len(temp_files) // 10)  # 10分の1のサイズ
        chunked_files = list(split_into_chunks(temp_files, chunk_size))

        partial_clips = []  # 後で連結するための部分クリップのリスト

        for i, file_chunk in enumerate(chunked_files):
            """
            各チャンクに対してビデオクリップを連結し、部分的なMP4を作成する
            """
            video_clips = [VideoFileClip(f) for f in file_chunk]
            final_clip = concatenate_videoclips(video_clips, method="compose")

            # 出力する部分MP4ファイルのパスを生成
            output_mp4_path = os.path.join(output_folder, f"instant_english_practice_part_{i + 1}.mp4")
            
            # 部分クリップを保存
            final_clip.write_videofile(output_mp4_path, fps=24, codec="libx264", audio_codec="aac")
            print(f"MP4 file saved successfully to: {output_mp4_path}")
            
            # 後で連結するために部分クリップをリストに追加
            partial_clips.append(final_clip)

        # 全ての部分クリップを連結して1つの最終MP4ファイルを作成
        final_video = concatenate_videoclips(partial_clips, method="compose")
        final_output_path = os.path.join(output_folder, "final_instant_english_practice.mp4")

        # 1つの最終ファイルを書き出し
        final_video.write_videofile(final_output_path, fps=24, codec="libx264", audio_codec="aac")
        print(f"Final MP4 file saved successfully to: {final_output_path}")

        # 作成された一時ファイルを削除
        for f in temp_files:
            os.remove(f)
                    
    except Exception as e:
        # エラーが発生した場合の処理
        print(f"Error writing MP4 file: {e}")
else:
    print("No valid video clips to concatenate.")
