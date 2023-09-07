import os
import re
from moviepy.editor import AudioFileClip, ImageClip, concatenate_videoclips

# Helper function to extract index from file name using your specified format
def extract_index(file_name):
    matches = re.findall(r'image(\d+)|audio_(\d+)', file_name)
    return int(matches[0][0] if matches[0][0] else matches[0][1]) if matches else None

def create_audio_video_slides(presentation_name, slide_paths, audio_paths):
    video_clips = []

    # Create new directory 
    output_directory = 'C:\\PPTVoiceOver\\pptData\\{}'.format(presentation_name)
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Sort image paths based on index in file name
    slide_paths.sort(key=extract_index)

    # Create a map of indices to audio files
    audio_map = {extract_index(f): f for f in audio_paths}

    for slide_path in slide_paths:
        img = ImageClip(slide_path).set_duration(5)  # Set default duration for slides without audio
        img = img.resize(height=480)
        
        idx = extract_index(slide_path)
        if idx in audio_map:  # if an audio file exists for this slide index
            audio = AudioFileClip(audio_map[idx])
            img = img.set_duration(audio.duration)
            img = img.set_audio(audio)

        video_clips.append(img)

    final_video = concatenate_videoclips(video_clips)

    output_file = os.path.join(output_directory, "final_output.mp4")
    final_video.write_videofile(output_file, fps=24)  # Specify fps
#,codec='libx265'

# images_paths=['C:\\PPTVoiceOver\\Images\\example_1\\image0.png', 'C:\\PPTVoiceOver\\Images\\example_1\\image1.png', 'C:\\PPTVoiceOver\\Images\\example_1\\image2.png', 'C:\\PPTVoiceOver\\Images\\example_1\\image3.png', 'C:\\PPTVoiceOver\\Images\\example_1\\image4.png', 'C:\\PPTVoiceOver\\Images\\example_1\\image5.png', 'C:\\PPTVoiceOver\\Images\\example_1\\image6.png', 'C:\\PPTVoiceOver\\Images\\example_1\\image7.png', 'C:\\PPTVoiceOver\\Images\\example_1\\image8.png']
# audio_paths=['C:\\PPTVoiceOver\\audios\\example_1_audios\\audio_0.mp3','C:\\PPTVoiceOver\\audios\\example_1_audios\\audio_1.mp3','C:\\PPTVoiceOver\\audios\\example_1_audios\\audio_2.mp3','C:\\PPTVoiceOver\\audios\\example_1_audios\\audio_3.mp3','C:\\PPTVoiceOver\\audios\\example_1_audios\\audio_4.mp3','C:\\PPTVoiceOver\\audios\\example_1_audios\\audio_5.mp3','C:\\PPTVoiceOver\\audios\\example_1_audios\\audio_6.mp3','C:\\PPTVoiceOver\\audios\\example_1_audios\\audio_7.mp3']
# create_audio_video_slides("example_1", images_paths, audio_paths)