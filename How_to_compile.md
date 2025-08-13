#How to compile excel_processor on win10
The newest python which was verified is 3.9.13, any version newer than that could cause compatibility trouble.
##step-by-step compile tutorial
###Install Python 3.9.13
Download the Python 3.9.13 installer from: https://www.python.org/downloads/release/python-3913/.
Run the installer, checking "Add Python 3.9 to PATH".
Verify installation:
```
textpython3.9 --version
```
Output should be Python 3.9.13.
###Create a Virtual Environment
Copy excel_processor.py to a new folder.
open cmd in that folder.
Create a virtual environment in that new folder:
```
python3.9 -m venv venv39
```
Activate it:
```
venv39\Scripts\activate
```
Confirm Python version:
```
python --version
```
###Install Dependencies
Install the required packages:
```
pip install numpy==1.23.5 pandas==1.5.3 xlrd==1.2.0 openpyxl xlwt xlutils xlsxwriter pyinstaller==5.13.2
```
Verify versions:
```
pip show numpy pandas xlrd openpyxl xlwt xlutils xlsxwriter pyinstaller
```
Ensure numpy==1.23.5 pandas==1.5.3, xlrd==1.2.0, pyinstaller==5.13.2.
###Clean Build Artifacts
Remove previous PyInstaller build files:
```cmd
rmdir /S /Q build
rmdir /S /Q dist
```
###Test the Script
Place rules.xls, input files, and output files in the folder where excel_processor.py located.
Run:
```
python excel_processor.py
```
Check if there is any error.
###Compile with PyInstaller
Update the --add-data paths to match the virtual environment’s site-packages **(use pip show openpyxl to find exact paths)**:
pyinstaller --clean --onefile --hidden-import=pandas --hidden-import=openpyxl --hidden-import=xlrd --hidden-import=xlwt --hidden-import=xlutils --hidden-import=xlutils.copy --add-data "$$THE_NEW_FOLDER$$\venv39\Lib\site-packages\openpyxl;openpyxl" --add-data "$$THE_NEW_FOLDER$$\venv39\Lib\site-packages\xlrd;xlrd" --add-data "$$THE_NEW_FOLDER$$\venv39\Lib\site-packages\xlwt;xlwt" --add-data "$$THE_NEW_FOLDER$$\venv39\Lib\site-packages\xlutils;xlutils" excel_processor.py
###Check if the compiled program could run
run:
```
.\excel_processor.py
```
Check the output, if there is no error and the output file is modified just as rules.xls asked, the compile work is complicated.