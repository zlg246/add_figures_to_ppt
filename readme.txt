this is a tool for pasting figures into ppt. It extracts figures in each subfolder, rescales to the same width and height, and pastes them into one slide, i.e. creating one slide for each subfolder.
It curently assumes all subfolders are directly under the input directory, and it supports a maximum of 8 figures in each subfolder.
Supported image types: png, jpg, tif, tiff and more types (needs code revision). Note: tif and tiff images are internally converted to png.

1) how to use:

Step 1: prepare the Python environment (python 3.7): 
pip install -r requirements.txt

Step 2: prepare the data. 
In a certain directory like "D:\experiments\figures", create subfolders of any valid names (e.g., "page 1", "page 2", ... "page n") and then add figures into every subfolder. For empty subfolder (no image), a blank slide will be created.

Step 3: open Windows command terminal by typing "cmd" in the search bar. Then change the working directory to the figure directory by typing "cd "D:\experiments\figures"".

Step 4: type a command like: python main.pyc -i "D:\experiments\figures" -o "C:\Users\lgzha\Desktop" -n "ppt_name"
see more command samples below.

python main.pyc
python main.pyc -i "D:\experiments\figures"
python main.pyc -o "C:\Users\lgzha\Desktop"
python main.pyc -i "D:\experiments\figures" -o "C:\Users\lgzha\Desktop"
python main.pyc -i "D:\experiments\figures" -o "C:\Users\lgzha\Desktop" -n "ppt_name"


2) get help: run python main.pyc -h

usage: main.py [-h] [-i I] [-o O] [-n N]

A tool for pasting figures to PPT

optional arguments:
  -h, --help  show this help message and exit
  -i I        input directory, default: D:\experiments\figures_to_ppt\input
  -o O        output directory, default: D:\experiments\figures_to_ppt\output
  -n N        ppt name, default: ppt_name
  -r {y,n}    if keep image ratio: default: y

notes: "-i" is for input figure directory; 
     	"-o" is for output ppt directory; 
	"-n" is the output ppt name;
	"-r" keeping original image ratio, or using the same image width and height.

