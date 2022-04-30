# pptx-chunker

Splits a Powerpoint presentation into PNG files organized into directories 
based on presentation sections.

This uses Win32 COM to open Powerpoint and save slides as PNG files, then figures
out which slides go where.

> Note: This requires you to be on Windows and have Powerpoint installed. There
> are probably ways to do this without using Powerpoint itself, but this matches
> my workflow.

## Installation

To get started, create a virtual environment and install `pywin32`:

```
> python -m venv venv
> .\venv\Scripts\Activate.ps1
> pip install --upgrade pip
> pip install -r .\requirements.txt
```

## Usage

To run this script, pass it the path to the presentation and a directory 
to use as the base for the output files:

```
> .\presentation_chunker.py SLIDES.PPTX BASE_OUTPUT_DIR
```

For example, if you use this on the `example.pptx` presentation included
in this repo, you might do it like this:

```
> .\presentation_chunker.py .\example.pptx .\output_dir
```

and you'd end up with this:

```
> tree /F .\output_dir\
C:\ALEX\PPTX-CHUNKER\OUTPUT_DIR
├───01-First Section
│       Slide1.PNG
│
├───02-Second Section
│       Slide2.PNG
│       Slide3.PNG
│
└───03-Another Section
        Slide4.PNG
        Slide5.PNG
```