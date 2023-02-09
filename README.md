# Project Austerlitz

# MD&A Shell Maker
This is a script that prepares the sections in which results of operations are read, where cash flows are discussed and where critical accounting policies are declared. 

# How To Use
To use it the script, you must have a valid TeX distribution. If you do not have one, you may download one [here](https://tug.org/mactex/morepackages.html). A word version can be made available upon request. You must save the script to the same folder as the folder in which you have saved the excel file that holds the data you are reading. You'll need to modify the script to indicate the sheet numbers as appropriate. For convenience, I've marked the places in the script where you'll need to make these edits. 

## Step 1. Download financials in Excel. 

If you don't know how to do this, send me an email at javaad.ali@advocatesclose.nyc and I'll walk you through it.

## Step 2. Download the scrpt.
Be sure to save the script to the same file as the file to which you have downloaded the financials.

## Step 3. Open Terminal. 

## Step 4. Navigate to the project file in terminal. 
For example, if you have saved the project files to your desktop, navigate to the desktop by typing `cd Desktop`.

## Step 5. Install the necessary Python packages.
To do this, type the following commands in terminal:

> - `pip install pylatex`
> - `pip install openpyxl`

## Step 6. Run the script. 
In terminal, navigate to where you have saved the project files, per step 4 above. Then type:

> - `python3.10 MDandAShellGenerator.py`
