
import pyttsx3
import os
from pydub import AudioSegment

#AudioSegment.converter = "C:\\ffmpeg-6.0-full_build\\bin\\ffmpeg.exe"
#AudioSegment.ffprobe= "C:\\ffmpeg-6.0-full_build\\bin\\ffprobe.exe"

def text_to_audio(text, save_as):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    #engine.setProperty('voice', voices[0].id)  #changing index, changes voices. 0 for male
    engine.setProperty('voice', voices[1].id)# 1 for female
    engine.setProperty('rate', 150) #speed of voice
    engine.save_to_file(text, save_as)
    engine.runAndWait()

def texts_to_audio_files(presentation_name, texts):
    presentation_name="audios\\"+presentation_name + "_audios"

    # Create a directory with the presentation name, if it doesn't exist
    directory = os.path.join(os.getcwd(), presentation_name.replace(" ","_"))
    if not os.path.exists(directory):
        os.makedirs(directory)
        
    audio_files = []
    #count=0
    for i, text in enumerate(texts):
        #count=count+1
        audio_file = os.path.join(directory, f"audio_{i}.mp3")
        text_to_audio(text, audio_file)
        audio_files.append(audio_file)
    
    
    return audio_files

def combine_audio_files(audio_files, output_file):
    combined = AudioSegment.empty()
    print("Combining audios")
    for audio_file in audio_files:
        sound = AudioSegment.from_mp3(audio_file)
        combined += sound

    combined.export(output_file, format="mp3")

# presentation_name = "example_1"
# texts = ["Slide 1 text", "Slide 2 text", "Slide 3 text"]
# audio_files = texts_to_audio_files(presentation_name, texts)
# print(audio_files)
# combine_audio_files(audio_files, f"{presentation_name.replace(' ', '_')}/combined.mp3")

