# Kindle Scribe Sync

![Kindle Scribe Sync Icon](https://github.com/Koloss5421/KindleScribeSync/blob/main/KindleScribeSyncIcon.png?raw=true)

### Built with
 - Python 3.13
 - Requests
 - Selenium
 - img2pdf
 - pystray
 - schedule
 - tarfile

## Getting Started
Syncs Kindle Scribe notebooks using an android user agent to access the notebook files. 
Keeps a local record of the last update time. Runs a check every 5 minutes.

Using `pystray` to have a "Force Sync" and "Last Update" button in the system tray.

![Kindle Scribe Sync Screenshot](https://github.com/Koloss5421/KindleScribeSync/blob/main/docs/screenshot.png?raw=true)

This is currently only made for running on windows. 
You must authenticate through the selenium browser, 2 min timeout when it opens.
The cookies are then saved for use later by requests.

There may be some bugs. Not sure how much I will maintain it.

## Installation

Clone the repository
```
git clone https://github.com/Koloss5421/KindleScribeSync
```

Setup virtual environment
```
python3 -m venv ./venv
```

Active virtual environment
```
./venv/Scripts/activate
```

Install requirements
```
pip install -r requirements.txt
```

Run it!
```
python .\KindleScribeSync.py
```

