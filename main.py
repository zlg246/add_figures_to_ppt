# python -m compileall main.py

import os
import glob
import argparse
import tifffile
import numpy as np
import PIL
import sys

from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

FIGURE_EXT = "png"

def scale_image(im):
    """
    Scale an image to the range [0, 1]
    """
    dynamic_range = im.max() - im.min()

    if dynamic_range > 0:
        return (im - im.min()) / float(dynamic_range)
    else:
        return np.zeros_like(im)

def open_image(fn, rgb2gray=True, cast_long=True):
    im = None
    if os.path.splitext(fn)[1].lower() in ['.tif', '.tiff']:
        im = tifffile.imread(fn)
        if cast_long:
            im = im.astype(np.uint16)
    else:
        im = PIL.Image.open(fn)
        im = np.array(im)
    if im.ndim == 3 and rgb2gray:
        print('WARNING: 3-channel image. Using first channel.')
        im = np.ascontiguousarray(im[0, :, :])

    return im

def save_image(fnOut, im, scale=True):
    if im.dtype.type is np.uint8:
        if im.ndim == 2 and scale:
            im = (scale_image(im.astype(np.float64)) * 255).astype(np.uint8)
        pil_im = PIL.Image.fromarray(im)
        pil_im.save(fnOut)
    elif im.dtype.type is np.uint16:
        tifffile.imsave(fnOut, im)
    elif im.dtype.type is np.float32 and im.ndim == 2:
        im = (scale_image(im) * 255).astype(np.uint8)
        pil_im = PIL.Image.fromarray(im)
        pil_im.save(fnOut)
    else:
        print("ERROR: Unsupported file type. Type: %s  Dim: %d  Filename: %s" % (str(im.dtype.type), im.ndim, fnOut))
        sys.exit()

def main():
    # test
    # add arguments
    parser = argparse.ArgumentParser(description='paste figures to PPT')
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parser.add_argument('--figure-path', type=str, default=os.path.join(current_dir, "figures"), help = "default: %(default)s")
    parser.add_argument('--output-path', type=str, default=os.path.join(current_dir, "output"), help = "default: %(default)s")
    parser.add_argument('--ppt-name', type=str, default="output", help = "default: %(default)s")
    args = parser.parse_args()
    figure_path = args.figure_path
    output_path = args.output_path
    ppt_name = args.ppt_name

    if not os.path.exists(output_path):
        print (f"ERROR: {figure_path} does not exist!")

    if not os.path.exists(output_path):
        os.makedirs(output_path)
    else:
        print ("Warning: output_path already exists!")

    # convert tif to png if tifs exist
    for root, dirs, files in os.walk(figure_path):
        for file in files:
            if file.lower().endswith(('.tiff', '.tif')):
                fn = os.path.join(root, file)
                im = open_image(fn).astype(np.float32)
                im_name = os.path.splitext(os.path.split(fn)[1])[0]
                im_out = os.path.join(root, im_name + f".{FIGURE_EXT}")
                save_image(im_out, im)

    prs = Presentation()

    # default slide width
    #prs.slide_width = 9144000
    # slide height @ 4:3
    #prs.slide_height = 6858000
    # slide height @ 16:9
    prs.slide_height = 5143500

    # title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])

    # set title
    title = slide.shapes.title
    title.text = "add title here"

    # setting for the 2nd and following slides
    figure_spacing   = int(prs.slide_height * 0.05)
    title_top   = int(prs.slide_height * 0.03)
    title_height = int(prs.slide_height * 0.15)
    title_width = int(prs.slide_width * 0.85)

     # for images in all subdirectories
    for folder_name in os.listdir(figure_path):
        folder_path = os.path.join(figure_path, folder_name)
        print ("folder: %s" % folder_path)
        if os.path.isdir(folder_path):
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            
            # set title
            title = slide.shapes.title
            title.text = folder_name
            title.top = title_top
            title.height = title_height
            title.width = title_width
            # force title box at center
            title.left = (prs.slide_width - title.width)//2

            # control text size and alignment
            title_para = title.text_frame.paragraphs[0]
            title_para.font.size = Pt(40)
            title_para.alignment = PP_ALIGN.CENTER

            # get path of all images
            fns = glob.glob(os.path.join(folder_path, f"*.{FIGURE_EXT}"))
            num_figures = len(fns)

            for e, fn in enumerate(fns):
                im = open_image(fn).astype(np.float32)
                im_name = os.path.split(fn)[1].strip()

                if num_figures <= 2:
                    figure_top   = int(prs.slide_height * 0.3)
                    figure_bottom   = int(prs.slide_height * 0.1)
                    figure_height = prs.slide_height - figure_top - figure_bottom
                    figure_width = int(figure_height * im.shape[1] / im.shape[0])
                    figure_left = (prs.slide_width - figure_width * num_figures - figure_spacing)//2
                    figure_left = e*(figure_width + figure_spacing) + figure_left
                elif num_figures <= 4:
                    figure_top   = int(prs.slide_height * 0.3)
                    figure_left  = int(prs.slide_width * 0.05)
                    figure_width = int((prs.slide_width - 2*figure_left - (num_figures - 1)* figure_spacing)/num_figures)
                    figure_height = int(figure_width * im.shape[0] / im.shape[1])
                    figure_left = e*(figure_width + figure_spacing) + figure_left
                elif num_figures <= 8:
                    # divide into two rows
                    figure_top   = int(prs.slide_height * 0.20)
                    figure_left  = int(prs.slide_width * 0.05)
                    num_figures_row = 4
                    figure_width = int((prs.slide_width - 2*figure_left - (num_figures_row - 1)* figure_spacing)/num_figures_row)
                    figure_height = int(figure_width * im.shape[0] / im.shape[1])
                    
                    if e <= 3:
                        figure_left = e*(figure_width + figure_spacing) + figure_left
                    else:
                        figure_top = figure_top + figure_height + figure_spacing
                        # recalculate left position for the second row
                        figure_left = (e - num_figures_row)*(figure_width + figure_spacing) + figure_left

                else:
                    print("ERROR: support a max of 8 images per slide!")
                    sys.exit()

                slide.shapes.add_picture(fn, figure_left, figure_top, figure_width, figure_height)

                # set textbox position
                text_left = figure_left
                text_top = figure_top + figure_height
                text_width = figure_height
                text_height = figure_spacing

                # add image name
                txBox = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
                tf = txBox.text_frame
                p = tf.paragraphs[0]
                p.text = im_name
                p.font.size = Pt(15)
    
    ppt_path = os.path.join(output_path, f"{ppt_name}.pptx")
    prs.save(ppt_path)
    print(f"ppt saved to {ppt_path}!")

if __name__ == '__main__':
    main()