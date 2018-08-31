# Canvas Report Generator
This program is now able to run on a directory of HTML files and allows you to enter either a Canvas course ID or a directory path

## DEPENDANCIES
They will be automatically installed when first running the program.
1. ImportExcel Module
2. BurntToast Module

## How to Run
Just run the .exe file for the report you want to generate. If it is the first time running it will ask you for certain credentials needed to fully run the program and then it will save them into the Passwords directory that will also be created on the first time being run. If you need to reset any of the data entered just delete the text fils or the whole Password directory to reset them.

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

## Google/YouTube API Key
In order for this program to scan YouTube videos for closed captioning, you will need to create a YouTube Data API key.

1. Go to the [Google Developer Console](https://console.developers.google.com).
2. Create a project.
3. Enable YouTube Data API
4. Create an API key

## BUGS

***This program also only checks pages that are inside of Modules and Discussions, it does not yet check Assignments/quizzes.***

## RECOGNICTION
Inspired by the VAST program originally created by the University of Central Florida at https://github.com/ucfopen/VAST

Able to work due to the Canvas APIs for PowerShell project at https://github.com/squid808/CanvasApis

	- The code I use from that project is contained in the PoshCanvasNew.ps1 file (I just cut out all of the functions I do not use to make the file smaller)

# Accessibility Report Generator
It does not catch every accessibility issue. For example:
1. It can't check anything that appears after JavaScript is run on the page

***This check can not tell if things are inaccessible if they rely on context*** (ex. it will only check if a table has any headers, not if the headers are correct or not)

# File directory
1. Files needed for accessibility report:
	1. Generate_Canvas_Accessibility_Report.ps1 (.exe, .exe.config as well)
	2. ProcessA11yReport.ps1
2. Files needed for media report:
	1. Generate_Canvas_Media_Report.ps1 (.exe, .exe.config as well)
	2. ProcessMediaReport.ps1
	9. BrightCoveSetUp.ps1
3. Files needed for both:
	1. SearchCourse.ps1
	2. Notifications.ps1
	3. FormatExcel.ps1
	4. Util.ps1
	5. PoshCanvasNew.ps1
	6. CheckModules.ps1
