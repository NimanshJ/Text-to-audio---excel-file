from elevenlabs.client import ElevenLabs
from openpyxl import load_workbook
from dotenv import load_dotenv
import librosa
import soundfile as sf
import os

load_dotenv()

def slow_audio_pitch_preserved(input_mp3, output_wav, speed=float(os.getenv("SPEED"))):
    y, sr = librosa.load(input_mp3, sr=None)
    y_slow = librosa.effects.time_stretch(y, rate=speed)
    sf.write(output_wav, y_slow, sr)
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
            voice_settings = {
        "stability": float(os.getenv("STABILITY")),
        "similarity_boost": float(os.getenv("SIMILARITY")),
        "style": float(os.getenv("STYLE")),
        "use_speaker_boost": True}
        )
    )

    mp3_file = f"{i}.mp3"
    wav_file = f"{i}.wav"

    with open(mp3_file, "wb") as f:
        f.write(audio)

    slow_audio_pitch_preserved(mp3_file, wav_file, speed=float(os.getenv("SPEED")))

print("Pitch-preserved slowed WAV files saved")
input("Press Enter to exit...")
