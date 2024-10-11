from moviepy.editor import AudioFileClip, concatenate_audioclips, ColorClip
import os
import re

# ffmpeg のパスを設定
ffmpeg_path = r"C:\ffmpeg\ffmpeg-master-latest-win64-gpl-shared\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"
print(f"Using ffmpeg path: {ffmpeg_path}")
os.environ["IMAGEIO_FFMPEG_EXE"] = ffmpeg_path

# 結合したいファイルが入っているフォルダを指定
folder_path = r"C:\Users\soyuk\Downloads\test\est"
print(f"Folder path set to: {folder_path}")

# 音声ファイルの確認と処理
file_names = os.listdir(folder_path)
pattern = re.compile(r'(\d+)\.JP\.mp3')
index_numbers = [int(pattern.search(f).group(1)) for f in file_names if pattern.search(f)]
max_index = max(index_numbers) if index_numbers else 0

# Silence.mp3 のパスを設定
silence_path = os.path.join(folder_path, "Silence.mp3")
if not os.path.exists(silence_path):
    print(f"Error: 'Silence.mp3' not found in {folder_path}. Please add 'Silence.mp3' to the folder.")
else:
    if max_index == 0:
        print("Error: No valid audio files found in the specified folder.")
    else:
        audio_clips = []
        for i in range(1, max_index + 1):
            jp_path = os.path.join(folder_path, f"{i}.JP.mp3")
            en_path = os.path.join(folder_path, f"{i}.EN.mp3")

            # Silence.mp3 を JP と EN の間に挿入
            if os.path.exists(jp_path) and os.path.exists(en_path):
                try:
                    jp_clip = AudioFileClip(jp_path)
                    print(f"Loaded JP audio: {jp_path}")  # デバッグ用出力
                    en_clip = AudioFileClip(en_path)
                    print(f"Loaded EN audio: {en_path}")  # デバッグ用出力
                    silence_clip = AudioFileClip(silence_path)
                    print(f"Loaded Silence audio: {silence_path}")  # デバッグ用出力

                    # JP, Silence, EN を順番に結合
                    combined_clip = concatenate_audioclips([jp_clip, silence_clip, en_clip])
                    audio_clips.append(combined_clip)

                except Exception as e:
                    print(f"Error loading audio clips for {i}: {e}")
                    continue  # エラーが発生した場合はスキップ

        if audio_clips:
            # 結合されたオーディオクリップのリストを最終的なオーディオクリップに結合
            try:
                final_audio = concatenate_audioclips(audio_clips)
                print(f"final_audio type: {type(final_audio)}")  # 型確認

                # 音声をベースにした mp4 を作成するための背景を生成
                # 白い背景（ColorClip）を設定し、動画のサイズを指定
                # (例: サイズ=(1920, 1080)の白背景を作成)
                background_clip = ColorClip(size=(1920, 1080), color=(255, 255, 255), duration=final_audio.duration)

                # 音声と背景を組み合わせて mp4 を作成
                video_clip = background_clip.set_audio(final_audio)
                output_mp4_path = os.path.join(folder_path, "instant_english_practice.mp4")
                video_clip.write_videofile(output_mp4_path, fps=1, codec="libx264", audio_codec="aac")

                print(f"MP4 file saved successfully to: {output_mp4_path}")
            except Exception as e:
                print(f"Error writing MP4 file: {e}")
        else:
            print("No valid audio clips to concatenate.")
