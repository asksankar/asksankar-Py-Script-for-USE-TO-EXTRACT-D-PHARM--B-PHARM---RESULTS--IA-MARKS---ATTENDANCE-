RESULT ANALYZER
PYTHON INSTALLATION & SETUP HELP
For Windows users running the supplied General D.Pharm / B.Pharm Result + IA Analyzer
This guide explains how to install Python and the packages required to run the supplied RESULT ANALYZER .py file on a Windows computer. The analyzer itself states that it is compatible with Python 3.10+ / Python 3.14. Its main external packages are PyMuPDF, pandas and openpyxl; xlrd is needed when legacy .xls IA files are to be read. Tkinter is used for the GUI and is normally included with the standard Windows Python installation.
1. What You Need
Item	Required?	Purpose
Windows PC	Yes	Run the GUI program
Python 3.10 or newer	Yes	Runs the .py program
PyMuPDF	Yes	Reads PDF files
pandas	Yes	Reads/handles spreadsheet data
openpyxl	Yes	Creates/reads Excel .xlsx files
xlrd	Yes for .xls	Reads older Excel .xls files
Tkinter	Normally included	Provides the GUI
2. Recommended Python Version
For a new Windows installation, use a current 64-bit Python 3 release. As of this guide, the official Python Windows downloads page lists Python 3.14.7 as the latest Python 3 release. The supplied analyzer is designed for Python 3.10+ / Python 3.14. 
Official Python Windows downloads: https://www.python.org/downloads/windows/
Official Python Windows documentation: https://docs.python.org/3/using/windows.html
3. Install Python - Easiest Windows Method
1.	Open the official Python Windows downloads page.
2.	Choose the current Windows 64-bit installation option suitable for your computer. The traditional Windows installer remains available for Python 3.14 releases.
3.	Run the downloaded installer.
4.	On the first installer screen, enable the option to add Python to PATH when that option is shown.
5.	Continue with the normal installation.
6.	After installation, close any Command Prompt / PowerShell window that was already open and open a new one.
4. Check That Python Is Installed
Open Command Prompt:
Win + R  →  type: cmd  →  Enter
Then try these commands:
py --version
python --version
A valid installation should show a Python 3.x version. On newer Windows Python installations, the py command can be available even when the python command is not.
•	If py --version works, you can continue using the py command for package installation and for running the analyzer.
•	If both commands fail, Python is not correctly installed or the command path is not configured.
5. Important: Python Command Not Recognized
If you see an error such as:
'python' is not recognized as an internal or external command
first try:
py --version
•	If py works, use py -3.14 or py -3 instead of python for the remaining commands.
•	If neither py nor python works, reinstall Python and enable the PATH option when available, or use the Python installation manager / Windows Python configuration described in the official Python documentation.
•	Do not install a second copy blindly. First check which Python executable Windows is using.
6. Install the Required Packages
Recommended commands for Python 3.14:
py -3.14 -m pip install --upgrade pip
py -3.14 -m pip install --upgrade pymupdf pandas openpyxl xlrd
If your system uses py -3 rather than py -3.14:
py -3 -m pip install --upgrade pip
py -3 -m pip install --upgrade pymupdf pandas openpyxl xlrd
Using python -m pip is also valid when the python command is correctly configured:
python -m pip install --upgrade pymupdf pandas openpyxl xlrd
7. Why These Packages Are Required
Package	Used by analyzer	Function
PyMuPDF	Yes	Opens and extracts text from result and IA PDF files.
pandas	Yes	Reads spreadsheet/CSV IA data and processes tables.
openpyxl	Yes	Creates the consolidated Excel workbook.
xlrd	For .xls inputs	Allows reading older Excel .xls files.
tkinter	Yes	Provides the desktop GUI; normally bundled with standard Python on Windows.
8. Test Every Dependency
py -3.14 -c "import pymupdf, pandas, openpyxl, tkinter; print('ALL REQUIRED PACKAGES OK')"
If you plan to use old .xls IA files, also test:
py -3.14 -c "import xlrd; print('xlrd OK')"
Expected result:
ALL REQUIRED PACKAGES OK
•	If a ModuleNotFoundError appears, install the missing package using the pip command in Section 6.
9. Put the Result Analyzer in a Convenient Folder
7.	Create a folder such as C:\Result_Analyzer
8.	Copy the supplied .py file into that folder.
9.	Keep your result PDFs and IA files in separate folders if that makes organization easier.
10.	Do not rename the .py file to .txt.
11.	Windows Explorer should show the file extension as .py.
10. Run the Result Analyzer
From Command Prompt:
cd /d C:\Result_Analyzer
py -3.14 General_DPharm_BPharm_Result_IA_Extractor_TWICE_CHECKED(1).py
If your filename is different, use the actual filename. You can also drag the .py file into the Command Prompt window after typing py -3.14 and a space.
•	The GUI window should open.
•	Add Result PDF file(s).
•	Add IA PDF / Excel / CSV file(s).
•	Click PROCESS + EXPORT EXCEL.
•	Choose the destination for the .xlsx workbook.
11. If the GUI Does Not Open
•	Run the program from Command Prompt instead of double-clicking it. The Command Prompt will display the actual error message.
•	If the error mentions pymupdf, pandas, openpyxl or xlrd, reinstall the missing package.
•	If the error mentions tkinter, verify that standard Python was installed rather than a minimal/embedded distribution.
•	If the error is related to the Python path, test py --version and use py -3.14 for package installation and execution.
12. Check Which Python Is Being Used
py -0p
where python
where py
These commands help identify multiple Python installations. When there is more than one Python installation, use the same launcher/version for both package installation and program execution.
13. Recommended One-Time Setup Sequence
py --version
py -3.14 --version
py -3.14 -m pip install --upgrade pip
py -3.14 -m pip install --upgrade pymupdf pandas openpyxl xlrd
py -3.14 -c "import pymupdf, pandas, openpyxl, tkinter, xlrd; print('SETUP COMPLETE')"
cd /d C:\Result_Analyzer
py -3.14 General_DPharm_BPharm_Result_IA_Extractor_TWICE_CHECKED(1).py
14. Common Errors and Solutions
Error / Symptom	Action
python is not recognized	Try py --version. If that also fails, reinstall/configure Python.
No module named 'pymupdf'	Run py -3.14 -m pip install --upgrade pymupdf.
No module named 'pandas'	Run py -3.14 -m pip install --upgrade pandas.
No module named 'openpyxl'	Run py -3.14 -m pip install --upgrade openpyxl.
No module named 'xlrd'	Run py -3.14 -m pip install --upgrade xlrd, especially if using .xls.
GUI does not appear	Run the .py from Command Prompt and read the error shown there.
Works in one Command Prompt but not another	Close old terminals and open a new one after changing PATH/installing Python.
Wrong Python version is used	Use py -3.14 explicitly and install packages into that same version.
15. Important Note About Packages
The analyzer source itself checks for missing pandas, PyMuPDF and openpyxl and displays installation guidance. Its Excel reader also supports .xls through xlrd. The analyzer's standard GUI uses tkinter.
16. Security / Good Practice
•	Download Python from the official python.org site.
•	Do not download random copies of Python or packages from unknown websites.
•	Keep the analyzer, input files and exported workbooks in known folders.
•	Back up the original university result and IA source files.
•	When troubleshooting, keep the Command Prompt error text because it identifies the actual problem.
17. Quick Reference
Task	Command
Check Python	py --version
Check Python 3.14	py -3.14 --version
Upgrade pip	py -3.14 -m pip install --upgrade pip
Install analyzer dependencies	py -3.14 -m pip install --upgrade pymupdf pandas openpyxl xlrd
Test dependencies	py -3.14 -c "import pymupdf, pandas, openpyxl, tkinter, xlrd; print('SETUP COMPLETE')"
Run analyzer	py -3.14 General_DPharm_BPharm_Result_IA_Extractor_TWICE_CHECKED(1).py
List installed Python versions	py -0p
18. Based on the Supplied Result Analyzer
The supplied program requires PyMuPDF for PDF processing and uses pandas and openpyxl when available. It also includes a spreadsheet reader for .xlsx, .xls and .csv IA inputs. The program title states Python 3.10+ / Python 3.14 compatibility.
Official References
Python Windows downloads: https://www.python.org/downloads/windows/
Python on Windows: https://docs.python.org/3/using/windows.html
Installing Python modules: https://docs.python.org/3.14/installing/
