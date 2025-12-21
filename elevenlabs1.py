from elevenlabs.client import ElevenLabs
from openpyxl import load_workbook
from dotenv import load_dotenv
import os

load_dotenv()

name = input("File name: ")
leng = int(input("Length: "))

client = ElevenLabs(api_key=os.getenv("KEY"))
workbook = load_workbook(f"{name}.xlsx")
sheet = workbook.active
for i in range(0,leng):

    audio = b"".join(
        client.text_to_speech.convert(
            text=sheet[f"{os.getenv("COLUMN")}{i+int(os.getenv("STARTING_ROW"))}"].value,
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

    with open(f"{i}.mp3", "wb") as f:
        f.write(audio)

print("Audio file saved")
input("Press Enter to exit...")


