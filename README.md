# Excel-automator thing
<a href="#overview">Overview</a> •
<a href="#prerequisites">Prerequisites</a> •
<a href="#set-up">Set-up</a> 

## Overview 
Just used for extracting the sheets from a spreadsheet and placing them in their own appropriately titled PDF
## Prerequisites
1. Python 
    * To check if downloaded you can run `python3 --version`
    * If not downloaded then go to the official website [here](https://www.python.org/ftp/python/3.9.10/python-3.9.10-macos11.pkg) and download version **3.9.10**
2. Git
    * To check if downloaded you can run `git --version`
3. Excel (ofc)

## Set-up
1. git clone this repo using the command `git clone https://github.com/rdavis-bc/invoice_automation.git`
2. Then move into the directory using `cd invoice_automation`
3. Run `git checkout feature_functions`
3. create a python environment with `python3 -m venv <Any name you want>` so for example `python3 -m venv excel_venv`
4. Activate the environment with the command `source <environment name>/bin/activate`
5. Install the needed libraries with `pip install -r requirements.txt`
6. Once those have finished installing , make a directory to place your excel workbooks and another directory for your pdfs
> For example you can create a structure like the following ![ok](images/Screen%20Shot%202022-03-23%20at%2011.52.23%20AM.png) where there is a subdirectory called `excel_sheets` and in this folder we place our `.xlsx` files . Moreoever, we have another subdirectory called `excel_pdfs` where the pdfs will go.
7. So inside `invoice_automation` we can just run a command with the following format `python src.main.py <directory of xlsx> <directory for pdfs>`\
In my case this would look like `python src/main.py excel_sheets excel_sheets/excel_pdfs` with no trailing slashes
> Note: Some errors will pop up in the command line but these aren't important for the time being\
Also you will get some pop-ups referring to access permissions on the folders where you are storing the files and these should be granted. One of them will be recurrent for the new folder (corresponding to the day) in `excel_pdfs`
