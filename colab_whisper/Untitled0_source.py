# %% Cell 1
!pip install git+https://github.com/openai/whisper.git
!sudo apt update && sudo apt install ffmpeg

# %% Cell 2
!whisper "20241128_160952-Michael-Serva.mp3" --model medium --language Spanish
