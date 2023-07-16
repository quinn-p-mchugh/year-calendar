---
title: "YearSpread: Visualize your year in a spreadsheet"
sidebar_label: "YearSpread" 
slug: year-spread 
description: "A handy calendar spreadsheet to visualize and sync annual events, trips, and more to a specified Google Calendar."
keywords: ["google sheets to google calendar", "sync spreadsheet to calendar", "year view google calendar", "yearly planning spreadsheet", "annual planning spreadsheet", "google sheets calendar sync", "spreadsheet integration with calendar", "google calendar synchronization", "yearly view in google calendar", "planning in spreadsheets"]
tags: ["Productivity", "Life", "Google Calendar", "Google Sheets"]
#image: ""
sidebar_position: 2
draft: false
publish: true
publish-path: docs-projects
publish-filename: 
created: 2023-07-01T11:53:11-04:00
modified: 2023-07-06T21:54:53-04:00
---

**YearSpread** is a yearly calendar template to help you plan out your year(s) in a digital format. I've found this calendar to be extremely useful for blocking out time for all-day events, weekend trips, vacations, etc. and keeping track of annual events like festivals or conferences. I hope you also find it useful!

## 👉 [Get the spreadsheet](https://docs.google.com/spreadsheets/d/1UezLLrUs2CTYpgBbuxlh0Mds6k-ykcy1hNU9sLdnbeQ/edit?usp=sharing)

## Sync your spreadsheet with a Google Calendar
YearSpread features one-way syncing (spreadsheet → calendar) with a Google Calendar of your choice. To enable this, follow the steps below:

### Part 1 of 2: Add Your Calendar ID
1. If you have not already done so, [make a copy of the spreadsheet](https://support.google.com/docs/answer/49114?hl=en&co=GENIE.Platform%3DDesktop#zippy=%2Cmake-a-copy-of-a-file).
2. Navigate to the `⚙️CONFIG` sheet.
3. In cell B3 (white box), enter the ID of the Google Calendar you wish to sync spreadsheet entries with. [Find your Google Calendar ID using these instructions](https://it.umn.edu/services-technologies/how-tos/google-calendar-find-your-google#:~:text=Finding%20Your%20Google%20Calendar%20ID&text=Click%20on%20the%20three%20vertical,will%20find%20your%20Calendar%20ID.). It should look something like `<list of letters and numbers>@group.calendar.google.com`.

:::danger
**All events in the calendar you specify will be DELETED prior to syncing.** It is strongly recommended you create a new calendar to use exclusively with this spreadsheet.
:::

### Part 2 of 2: Setup Periodic Syncing
Now that the calendar ID is added, we will setup a [time-driven trigger](https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers) to sync spreadsheet entries with our Google Calendar each hour:

1. On the top menu bar, click **Extensions > Apps Script**. You should be directed to the Google Apps Script interface.
2. Hover over the left sidebar and click **Triggers** (clock icon).
3. On the bottom right, press the **Add Trigger** button.
4. Specify the following settings:
    - Set **Choose which function to run** to `syncEventsWithGoogleCalendar`.
    - Set **Select event source** to `Time-driven`.
    - Set **Select type of time based trigger** to `Hour timer`.
    - Set **Select hour interval** to `Every hour` (or whichever time interval you prefer).
6. Click **Save**. An authorization window should appear.
7. In the authorization window, select your Google account.

:::info
After selecting your Google Account, you should receive a warning that "Google hasn't verified this app". This is a standard warning indicating that the script you're authorizing requires sensitive info from your Google Account (in this case, authorization to write to your calendar). While this warning should give you pause, there is no need to be alarmed here. The script is meant to simply create a calendar event for each entry in the calendar spreadsheet. There is no malicious code being run.

If you are hesitant to proceed, you are welcome to press **Back to Safety**, go to the **Editor** tab and inspect the script yourself. This code is local to your account and can only be edited by you, so feel free to modify as you please.
:::

8. In the authorization window, click **Advanced**.
9. Click **Go to sync-spreadsheet-events-with-google-calendar (unsafe)**.
10. Review the information. If you are ready to proceed, click **Allow**.

🎉 **Woo hoo! Any events you add to your year calendar will now be automatically synced to the calendar specified.**

### Testing
- To verify that sync is working, navigate to the `2024` sheet and enter an event title into any of the cells.
- After an hour passes, search for the event title in your Google Calendar - it should come up. 

### Adding Multiple Events To One Date
- To add multiple events for a single date, enter the title of each event separated by a "|". There is no limit to how many events you can add per cell.
  - For example: Let's say you wanted to add two events, "My First Event" and "My Second Event" on January 1st, 2024.
    - Navigate to the `2024` sheet and edit the cell next to January 1st.
    - Enter "**My First Event | My Second Event**".
    - During the next sync, the script create two separate calendar events for **My First Event** and **My Second Event**.
   
## Issues
If you encounter any issues with this spreadsheet or would like to submit a feature request for review, please **[submit an issue](https://github.com/quinn-p-mchugh/year-calendar/issues/new/choose)** in the associated GitHub repository. A GitHub account is required to submit issues. [If you do not have a GitHub account, you can create one here.](https://github.com/join)

