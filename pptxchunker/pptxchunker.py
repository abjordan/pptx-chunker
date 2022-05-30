# Resources used:
#   https://medium.com/@chasekidder/controlling-powerpoint-w-python-52f6f6bf3f2d
#   https://pbpython.com/windows-com.html
#   https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentation

import logging
import os
from pathlib import Path

from win32com.client import constants
from loguru import logger

from . import utils

def by_title(infile, outdir):
    return None

def by_section(infile, outdir):
    # Calling gencache.EnsureDispatch makes sure that pywin32 regenerates
    # the Python wrappers for the COM objects we're going to use. We need
    # this to access the constants for SaveAsPNG (for example)
    app = utils.dispatch("PowerPoint.Application")

    filepath = Path(infile).resolve()

    # before Python 3.10, there's a bug: https://bugs.python.org/issue38671 where
    # if you have a relative path to a non-existant directory or file, and you 
    # call resolve() on it, you get back the same relative path you provided, rather
    # than an absolute path to a missing entry.
    outdir = Path(outdir).resolve()
    outdir.mkdir(exist_ok=True)    
    outdir = Path(outdir).resolve()

    pres = app.Presentations.Open(filepath)

    section_count = pres.SectionProperties.Count

    sections = []

    for idx in range(1, section_count+1):
        name = pres.SectionProperties.Name(idx)
        first_slide = pres.SectionProperties.FirstSlide(idx)
        number = pres.SectionProperties.SlidesCount(idx)
        logger.debug("%s -- %d through %d" % (name, first_slide, first_slide+number-1))
        sections.append({ 
            'index': idx,
            'name': name, 
            'first_slide': first_slide,
            'count': number})

    pres.SaveAs(outdir, constants.ppSaveAsPNG)
    app.Quit()

    # At this point, we have a directory at `outdir` full of PNGs called slideNN.png
    # where NN is the slide count. We'll make a subdirectory for each of the sections
    # and then move the files into the right directory, based on the section it's in.
    # I don't know why this isn't an option in actual Powerpoint...

    orig_dir = os.getcwd()
    os.chdir(outdir)

    for section in sections:
        subdirname = "%02d-%s" % (section['index'], section['name'])
        section_dir = outdir / subdirname
        section_dir.mkdir(exist_ok=True)
        logger.debug(f"Moving {section['name']}: slides {section['first_slide']} -> {section['first_slide'] + section['count'] - 1}")
        for idx in range(section['first_slide'], section['first_slide'] + section['count']):
            os.rename("Slide%d.PNG" % idx, section_dir / f"Slide{idx}.PNG")
    logger.info(f'Moved slides into {len(sections)} sections at {outdir}')