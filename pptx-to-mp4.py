import os
import argparse
import time
import glob
import re
import logging
import shutil
import natsort

from pptx_tools import utils as pptx_utils

from ffmpeg import FFmpeg
import soundfile

parser = argparse.ArgumentParser("pptx-to-mp4")
parser.add_argument("filename", help="Filename of PowerPoint presentation to be converted (e.g. test.pptx)", type=str)
args = parser.parse_args()

logging.basicConfig(level=logging.INFO)

# File processing
logging.info("Filename: " + args.filename)
script_dir = os.path.dirname(os.path.realpath(__file__))
filename_pptx = os.path.join(script_dir, args.filename)
logging.info("Presentation location: " + filename_pptx)
working_subdir = os.path.join(script_dir, "".join(["working_dir_pptx","-",re.sub(".pptx", "", args.filename)]))
logging.info("Temporary Working Directory: " + working_subdir)
pptx_utils.save_pptx_as_png(working_subdir, filename_pptx, overwrite_folder=True)
logging.info("Conversion of presentation to images complete.")

# Extracting pptx to get at audio
shutil.copy(args.filename, working_subdir)
shutil.move(os.path.join(working_subdir, args.filename), os.path.join(working_subdir, "working.zip"))
shutil.unpack_archive(os.path.join(working_subdir, "working.zip"),os.path.join(working_subdir,"working"))

# Use slide audio - CAUTION Assumes all media audio is for slides, and that these are sequential
media_dir = os.path.join(working_subdir, "working", "ppt", "media")
audio_list = natsort.natsorted(glob.glob("*.wav", root_dir=media_dir))
slide_list = natsort.natsorted(glob.glob(("*.PNG"), root_dir=working_subdir))
audio_file_durations = []
for audio_file in audio_list:
    audio_file_var = soundfile.SoundFile(os.path.join(media_dir,audio_file))
    audio_file_duration = audio_file_var.frames/audio_file_var.samplerate
    audio_file_durations.append(audio_file_duration)
    shutil.copy(os.path.join(media_dir,audio_file), working_subdir)
os.chdir(working_subdir)
for slide in range(len(slide_list)):
    image_name = slide_list[slide]
    flag_no_audio = False
    try:
        audio_name = audio_list[slide]
    except:
        logging.warning("No audio found for slide " + str(slide + 1) + " of " + str(len(slide_list)))
        flag_no_audio = True
    output_name = "".join(["output", str(slide), ".ts"])
    tick = time.perf_counter()
    if flag_no_audio == False:
        video_out = FFmpeg().option("y").option(key="loop", value=1).input(image_name).input(audio_name).output(output_name, shortest=None, vcodec="libx264", acodec="aac")
    else: 
        video_out = FFmpeg().option("y").option(key="loop", value=1).input(image_name).output(output_name, vcodec="libx264", duration=5)
    video_out.execute()
    tock = time.perf_counter()
    render_time = tock - tick
    if flag_no_audio == False: 
        audio_length = float(audio_file_durations[slide])
        speed_up = audio_length/render_time
    else: 
        audio_length = 0
        speed_up = 0
    logging.info("".join(["Slide ", str(slide + 1), " of ", str(len(slide_list)), ": ", str(round(audio_length, 2)), " seconds. Render: ", str(round(render_time, 2)), " seconds, speed-up ", str(round(speed_up, 2)), " x."]))

# Concatenate all intermediate files to a full video
input_argument = "concat:" + "|".join(f"output{s}.ts" for s in range(len(slide_list)))
final_video = FFmpeg().input(input_argument).output("output.mp4", codec="copy")
tick = time.perf_counter()
final_video.execute()
tock = time.perf_counter()
render_time = tock - tick
audio_length = float(sum(audio_file_durations))
speed_up = audio_length/render_time
logging.info("".join(["Final render: ", str(round(audio_length, 2)), " seconds. Render: ", str(round(render_time, 2)), " seconds, speed-up ", str(round(speed_up, 2)), " x."]))

# Move video and tidy up
os.chdir('..')
final_video_name = re.sub("pptx", "mp4", args.filename)
shutil.move(os.path.join(working_subdir, "output.mp4"), os.path.join(script_dir, final_video_name))
logging.info("Output file moved to original directory.")
confirmation = input("The working directory is " + working_subdir + ". As a part of the clean-up process, please confirm you would like to delete it by entering DELETE.")
if confirmation == "DELETE": 
    shutil.rmtree(working_subdir)
else: 
    logging.warning("Confirmation not received. Working directory retained.")

