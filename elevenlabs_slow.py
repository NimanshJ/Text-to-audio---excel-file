from elevenlabs.client import ElevenLabs
from openpyxl import load_workbook
from dotenv import load_dotenv
import subprocess
import os
import sys

load_dotenv()

def get_ffmpeg_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, "ffmpeg.exe")
    return "ffmpeg"

def slow_audio_ffmpeg(input_mp3, output_wav, speed):
    ffmpeg = get_ffmpeg_path()

    # atempo supports 0.5–2.0, so we chain if needed
    atempo = speed
    cmd = [
        ffmpeg,
        "-y",
        "-i", input_mp3,
        "-filter:a", f"atempo={atempo}",
        output_wav
    ]

    subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.remove(input_mp3)

name = input("File name: ")
leng = int(input("Length: "))

client = ElevenLabs(api_key=os.getenv("KEY"))
workbook = load_workbook(f"{name}.xlsx")
sheet = workbook.active

for i in range(leng):
    text = sheet[f"{os.getenv('COLUMN')}{i + int(os.getenv('STARTING_ROW'))}"].value

    audio = b"".join(
        client.text_to_speech.convert(
            text=text,
            voice_id=os.getenv("VOICE_ID"),
            model_id="eleven_multilingual_v2",
            voice_settings={
                "stability": float(os.getenv("STABILITY")),
                "similarity_boost": float(os.getenv("SIMILARITY")),
                "style": float(os.getenv("STYLE")),
                "use_speaker_boost": True
            }
        )
    )

    mp3 = f"{i}.mp3"
    wav = f"{i}.wav"

    with open(mp3, "wb") as f:
        f.write(audio)

    slow_audio_ffmpeg(mp3, wav, float(os.getenv("SPEED")))

input("Press Enter to exit...")
