# FreeTime

FreeTime is a lightweight system tray application that helps you quickly find and share your available meeting times. It checks your calendars and generates a list of free time slots that can be instantly pasted anywhere when you type a trigger phrase (eg ":ttt").

<img src="screenshots/freetime_paste.gif" width="400">

## Features

- 🔍 Monitors multiple calendars simultaneously
- ⚡ Instant paste of available times with customizable trigger phrase
- 🕒 Configurable working hours and date range
- 🔄 Automatic calendar updates
- 🌐 Works with any calendar that provides iCal/ICS feeds
- 💻 Runs in system tray
- 🚀 Single executable file, no installation needed

## Download

Download the latest version for your platform:
- [Windows](link-to-latest-windows-release)
- [macOS](link-to-latest-mac-release)

## Quick Start

1. Download and run FreeTime
2. Add your calendar URLs (see below for how to get these)
3. Set your preferred working hours and options
6. Use the trigger phrase anywhere (default ":ttt") to paste your available times

## Setting Up Calendar URLs

### Google Calendar

1. Open [Google Calendar](https://calendar.google.com/)
2. Click the three dots next to your calendar name
3. Select 'Settings and sharing'
4. Scroll to 'Integrate calendar'
5. Copy the 'Secret address in iCal format'

<img src="screenshots/google_calendar_setup.png" width="600">

**Note:** This URL is private - anyone with this URL can view your calendar. There is no security risk using this in FreeTime. The calendar information never leaves your device.
If you do not want to use the secret address, you can alternately use the 'Public address in iCal format' address. For this to work, you need to enable the option 'Make calendar available to public' near the top of the settings screen. You can make a calendar public but only share free/busy time.


### Microsoft Office 365

1. Open [Outlook Calendar](https://outlook.office.com/calendar)
2. Click the gear icon (Settings)
3. Select 'Shared Calendar'
4. Select your calendar
5. Choose "Can view when I am busy" under "Permissions"
6. Click "Publish"
7. Copy the ICS link

<img src="screenshots/office365_setup.png" width="600">

## FreeTime Configuration Options

### Time Settings
- **Meeting Hours:** Set your standard working hours
- **Lookahead Days:** Number of days to check for availability
- **Include Current Day:** Include today in available times
- **Exclude Weekends:** Remove weekends from available times
- **Timezone:** Set your local timezone

### App Settings
- **Update Interval:** How often to refresh calendar data
- **Trigger:** Customize the trigger phrase
- **Run at startup:** Launch automatically with system
- **Custom Text:** Modify the introductory text when pasting

## Building from Source

1. Clone the repository
```bash
git clone https://github.com/vince-p/freetime.git
cd freetime