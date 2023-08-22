PYTHON PROGRAM TO GENERATE A VISUAL BOQ OF TOUCH SWITCHES USING A BUILDTRACK PROPOSAL 

OVERVIEW:
The program called MAIN.PY will take the BuildTrack proposal and look at 2 tabs (INFINITY, DESIGNER). Then it use the BUILD YOUR SWITCH app from the SmartTouchSwitch website to produce images of each line on the 2 tabs mentioned. These images are stored in a TMP folder and then re-assembled by the MAIN.py program into a .DOC file OUTPUT which will have the same name as the proposal. After successful generation of the OUTPUT file the images in the TMP folder 
----------------------------------------------------------
CAUTION: 
* Do NOT modify the notepad file called 'REQUIREMENTS.TXT' under any circumstances
* Ensure only one Excel File (Proposal) is in the directory at all times
* Ensure that you have deleted/moved any OUTPUT (.doc) file and TMP folder which are generated after running the program
* 2 Windows will be open while the program is running - one is a python window, the other is the App for Build-Your-Switch which will be manipulated AUTOMATICALLY BY the python program to create the desired switches to capture their images
* All images captured from Build-Your-Switch app will be stored in the tmp folder. All the images of switches will be named from Switch01 onwards. This is just mentioned in case you need the images.. but they will all be in the OUTPUT doc also. 




Installing Pre-Requisites
	* Download Python 3.7 from https://www.python.org/ftp/python/3.7.0/python-3.7.0-amd64.exe
	* Install and add Python to your PATH (This should be in the Installer)
		- If this didn't work, you can manually add Python to your PATH via these steps - https://datatofish.com/add-python-to-windows-path/
	* Open a Command Prompt in the directory with the python file
	1. Download Python 3.7 from https://www.python.org/ftp/python/3.7.0/python-3.7.0-amd64.exe
	2. Install and add Python to your PATH (This should be an option that arises at the end of your installation process)
		- If this did not work, you can manually add Python to your PATH via the steps listed in the URL (https://datatofish.com/add-python-to-windows-path/ )
	3. To install the pre-requisites, open a Command Prompt in the directory where you downloaded the python file
		- Do this by going to the directory in the File Explorer
		- Click on the address bar (or press Alt + D)
		- Type in cmd into the address bar
		- In the pop-up window, run this command "pip install -r requirements.txt"

How to Run:
	* Place the Excel File (Proposal Workbook) into the directory
	* Run the Python File
		- Open the command prompt in the same way as earlier and run python main.py
		- Right Click on the File and Open With Python 3.7
	* The output will be in a Word Document with the same name as the Proposal File
How to Run the MAIN.py program
	1 Place the Excel File (BuildTrack Proposal Workbook) into the directory where the PYTHON program is stored
	2 Run the Python File using ANY one of the 2 options below
		- Open the command prompt in the same way as earlier and type "python main.py"
		- Right Click on the Python File and Open With Python 3.7
	3. The output will be in a Word Document with the same name as the Proposal File

Notes:
	* If you want to store the image folder (the tmp folder) rename it and move it outside the VisualBoQ
	* Ensure only one Excel File is in the directory at all times
	* Try not to interfere with the opened window while the program is running
