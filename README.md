# pptx-to-mp4
Converts narrated PowerPoint presentations to small mp4 videos, using recorded narration and timings. 
Be warned: this is very much a script written without the intent of being fully robust. 

Requires: 
- A Windows environment (Python file handling in the script does not like WSL. I have not debugged and fixed this).
- Microsoft PowerPoint (Required by a library)
- Python
- Pip3

# Preparation
- Install pip and python, if not present already.
- Copy the pptx-to-mp4.py script and the requirements.txt to a new directory (or git clone etc.). 
- Run `pip install -r requirements.txt`.

# Execution
- Copy your presentations into the folder.
- Rename your presentations to contain only alphanumeric characters, underscores, and dashes. 
- Run `python ./pptx-to-mp4.py your-presentation-name-here.pptx`.
- Keep an eye on the log.
- Confirm deletion of the temporary working directory set up by the script.
- Note: this will not work and there will be some files left. I have not debugged and fixed this either.

