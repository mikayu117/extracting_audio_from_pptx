from spire.presentation import *
import mimetypes
import ffmpeg
import os

# Presentationオブジェクトを作成
presentation = Presentation()
print("input file")
file_name = input()

# PowerPointファイルを読み込む
presentation.LoadFromFile(file_name)

k=0
j=""
while file_name[k]!='.':
    j += file_name[k]
    k+=1
i = 1
audio_files = []  # オーディオファイルのリスト

# すべてのスライドを反復処理
for slide in presentation.Slides:
    # スライド内のすべての図形を反復処理
    for shape in slide.Shapes:
        # 図形がオーディオかどうかを確認
        if isinstance(shape, IAudio):
            # オーディオデータを取得
            audioData = shape.Data
            output_file = f"Audio/Audio{j}-{i}{mimetypes.guess_extension(audioData.ContentType)}"
            audioData.SaveToFile(output_file)
            audio_files.append(output_file)  # リストにファイルパスを追加
            i += 1

presentation.Dispose()

# オーディオファイルを結合
if len(audio_files) > 1:
    inputs = [ffmpeg.input(file) for file in audio_files]
    concat = ffmpeg.concat(*inputs, v=0, a=1)
    output_combined = f"Audio/CombinedAudio{j}.m4a"

    # 結合処理
    concat.output(output_combined).run()
    print(f"Combined audio saved to: {output_combined}")

# 一時ファイルの削除
for file in audio_files:
    if os.path.exists(file):
        os.remove(file)

print("Processing complete.")
