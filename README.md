# SilentSlideUpdater

## Overview
SilentSlideUpdater is a C# utility that automatically updates a company-wide PowerPoint Show (.ppsx) file used for internal security campaigns. It runs silently in the background and replaces the current slide with the next one in sequence — only on the first Tuesday of each month.

## How It Works
The shared slide is stored at:

\\INTERNAL\Campaign\SecurityPoster.ppsx
Available campaign slides are organized in numbered folders like:


\\INTERNAL_IP\information system\Campaign\FY25\Security Poster\1\
\\INTERNAL_IP\information system\Campaign\FY25\Security Poster\2\
...

Each folder contains:

A .ppsx file (PowerPoint Show – for display)

A .pptx file (PowerPoint Edit – optional)


On the first Tuesday of each month:

The script checks the current .ppsx file.

It identifies which numbered folder it matches.

It copies the next folder’s .ppsx file and overwrites the current shared slide.

## Behavior Summary

Runs silently — no terminal window or pop-up

If already on the last slide, or no match found, the script exits without changes

Logs unexpected errors to: C:\Temp\SlideUpdateLog.txt

## Deployment Instructions

1. Build the Application
   - Open SilentSlideUpdater in Visual Studio
   - Set project type to Windows Application (not Console)
   - Build the project → get the .exe from bin\Release\

2. Copy the .exe to an admin-level machine (with access to the shared drive)
3. Set Up Windows Task Scheduler
   Create a new task:
   - Trigger: Monthly → First Tuesday → At startup or a fixed time
   - Action: Run the .exe
   - Run with highest privileges
   - Set to run whether user is logged on or not
   - Optionally, set window to hidden (default for Windows Application)
   
## Project Structure

```
SecurityCampaign/
│
├── Program.cs           ← Main script logic
├── SecurityCampaign.csproj ← Visual Studio project file
├── README.md            ← This file
└── SecurityCampaign     ← C# Solution file

```

## Permissions Notes
The script must be run with access to the shared network folder

Ensure write permission to \\INTERNAL_IP\Campaign\