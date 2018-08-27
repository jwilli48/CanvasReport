## DEPENDANCIES
They will be automatically installed when first running the program.
	ImportExcel module
	BurntToast module

## How to Run
The ps2exe.ps1 file is included as it is used to create and update the .exe files.
Now that they are executable files you no longer need to change your PowerShell policy.

To run this program from PowerShell if your PowerShell execution policy is restricted, do the following commands:
1. Navigate to folder containing these scripts (Change path to where you have it on your computer)
	cd 'C:\Users\Username\Documents\CanvasReport'
2. Type or copy the following into your PowerShell window
	Set-ExecutionPolicy Bypass -Scope Process
3. Run the script
	.\Generate_Canvas_Accessibility_Report.ps1
	.\Generate_Canvas_Media_Report.ps1

If you wish to be able to run the program without doing the above every time, then do the following:
1. Run the following command in any PowerShell window
	Set-ExecutionPolicy Bypass -Scope CurrentUser
2. You can then just right click the program and hit 'Run with PowerShell'

## Reports
The report will be generated and saved to the Report folder within this directory. If you try to create a 2nd report for a course while there is a previous one still there it will just add to the bottom of the previous one instead of creating a new one.
## First time running
The first time you run this it will ask you to input your Canvas API and the Canvas Default URL, as well as Brightcove credentials and a Google API
1.You will need to generate your own API from your Account Settings in Canvas
2.The default/base URL for BYU's canvas is https://byu.instructure.com

### Google/YouTube API Key
In order for this program to scan YouTube videos for closed captioning, you will need to create a YouTube Data API key.

1. Go to the [Google Developer Console](https://console.developers.google.com).
2. Create a project.
3. Enable YouTube Data API
4. Create an API key

## BUGS
When generating the Media report the 2nd and 3rd Pivot Charts do not correctly display the times. In order to fix this you need to right click on column B, hit Format Cells, Choose Custom and then go down and select the [h]:mm:ss format. You will need to do this for both pivot chart sheets to see the time displayed in hours.

## RECOGNICTION
Inspired by the VAST program originally created by the University of Central Florida at https://github.com/ucfopen/VAST
