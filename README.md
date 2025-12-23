Tech stack:
==================

For this timesheet‑automation project, the tech stack is quite small and focused. Here is what you need.

Core language and runtime:
---------------------------------
Python 3.8+ – main language for data processing, Excel handling, and GUI.
​

Python libraries (code side):
-------------------------------
1. Excel and data handling

pandas – python library (open source)for reading the source timesheet Excel, filtering, grouping, aggregating hours per employee/task, and reshaping data.
​

openpyxl – for creating the final Concentrix‑style Excel files with borders, colors, column widths, and writing calculated totals as values.
​

2. Desktop GUI:
tkinter (standard library) – to build the simple desktop UI with:

File‑open dialog for input Excel

Folder‑select for output location

Buttons and status labels
Tkinter ships with Python, so no extra install is needed.
​

3. Packaging to .exe (for Windows users)

PyInstaller – to convert app.py (Tkinter GUI) into a single Windows executable:

Command pattern: 
pyinstaller --onefile --windowed --name CNXTimesheetGenerator app.py

Produces ConcentrixTimesheet.exe that users can run without installing Python.
​

Python dependencies summary:

Install once in your virtualenv or system:
open windows powershell and run the following command.
$ pip install pandas openpyxl pyinstaller

After installing the dependencies, clone the repo and run the following command to launch the app.
$ python app.py
