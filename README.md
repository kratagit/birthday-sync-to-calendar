# 🎂 Birthday Sync To Calendar

A desktop application that helps you never forget important birthdays by syncing them with your Google Calendar. Simple, fast, and secure!

## 📌 What does this app do?

**Birthday Sync To Calendar** is a user-friendly desktop application that allows you to:
- Keep track of all your friends' and family members' birthdays in one place
- Automatically calculate and display everyone's current age
- Export all birthdays to Google Calendar with automatic yearly reminders
- Get notifications 12 hours before each birthday (via Google Calendar)
- View birthdays sorted by upcoming dates or chronologically

Once you add birthdays to the app, you can sync them with Google Calendar in one click. The app creates a dedicated "Birthdays" calendar in your Google account and adds recurring yearly events for each person. This means you'll receive calendar notifications and never miss an important birthday again!

## 💾 Where are my birthdays stored?

All your birthday data is saved locally on your computer in:
```
C:\Users\[YourUsername]\Documents\BirthdayApp\data.json
```

This file is created automatically when you first use the app. Your data stays on your computer and is never sent anywhere except to Google Calendar when you choose to sync.

## ✨ Features

### 📝 Birthday Management
- **Add birthdays** - Enter a person's name and birth date (day, month, year)
- **Delete birthdays** - Remove people from your list anytime
- **View all birthdays** - See everyone's name, birth date, and current age in a clean table
- **Automatic age calculation** - The app automatically updates everyone's age when you open it

### 🔄 Google Calendar Sync
- **One-click export** - Sync all birthdays to Google Calendar instantly
- **Dedicated calendar** - Creates a separate "Birthdays" calendar to keep things organized
- **Recurring events** - Birthdays repeat every year automatically
- **Smart reminders** - Default notification 12 hours before each birthday
- **No duplicates** - The app checks for existing events to avoid duplicates

### 📅 Sorting Options
- **Sort by nearest birthdays** - See whose birthday is coming up next
- **Sort chronologically** - View birthdays in calendar order (January to December)

### 🔐 Privacy & Security
- **Local storage** - All data stays on your computer
- **No data collection** - We don't collect any personal information
- **Google authorization only** - Only connects to Google Calendar when you choose to sync
- **Secure authentication** - Uses official Google OAuth 2.0 for safe login

## 🚀 Getting Started

### Installation

1. **Download and install the application** on your computer
2. **Run the application** - Double-click the application icon

### First-Time Setup for Google Calendar Sync

To use the Google Calendar sync feature:

1. Click the **"Eksportuj do Google Calendar"** (Export to Google Calendar) button
2. A browser window will open asking you to sign in to your Google account
3. Grant the application permission to access Google Calendar
4. Close the browser window when prompted
5. Your birthdays will be synced to a new "Birthdays" calendar in Google Calendar

**Note:** You only need to authorize once. The app will remember your authorization for future syncs.

## 📖 How to Use

### Adding a Birthday

1. Type the person's name in the **"Imię"** (Name) field
2. Select their birth date using the dropdown menus:
   - **Dzień** (Day): 1-31
   - **Miesiąc** (Month): 1-12
   - **Rok** (Year): 1900-current year
3. Click **"Dodaj osobę"** (Add Person)
4. The person will appear in the table below with their age

### Removing a Birthday

1. Click **"Usuń osobę"** (Delete Person)
2. Select the person you want to remove from the dropdown list
3. Confirm the deletion
4. The person will be removed from your list

### Syncing with Google Calendar

1. Click **"Eksportuj do Google Calendar"** (Export to Google Calendar)
2. Confirm that you want to sync
3. Wait for the sync to complete
4. Check your Google Calendar - you'll see a new "Birthdays" calendar with all events

### Sorting Your Birthday List

- Click **"Sortuj od najbliższych urodzin"** (Sort by Nearest Birthdays) to see upcoming birthdays first
- Click **"Sortuj chronologicznie"** (Sort Chronologically) to see birthdays in calendar order

## 💡 Tips

- The app automatically updates everyone's age each time you open it
- You can add birthdays anytime and sync later when convenient
- The Google Calendar sync creates yearly recurring events, so birthdays appear every year
- If you reinstall the app, your birthday data is safe in your Documents folder
- You can manually edit your birthday list by opening the `data.json` file in your Documents/BirthdayApp folder

## ⚙️ System Requirements

- **Operating System:** Windows
- **Internet Connection:** Required only for Google Calendar sync
- **Google Account:** Required only if you want to sync with Google Calendar

## 🆘 Troubleshooting

**Problem:** Can't find my birthday data
- **Solution:** Check `C:\Users\[YourUsername]\Documents\BirthdayApp\data.json`

**Problem:** Google Calendar sync isn't working
- **Solution:** Make sure you're connected to the internet and have authorized the app to access your Google Calendar

**Problem:** Birthdays aren't showing up in Google Calendar
- **Solution:** Look for a calendar named "Birthdays" in your Google Calendar sidebar and make sure it's checked/visible

## 👤 Author

**Krata**
- GitHub: [@gitzakratami](https://github.com/gitzakratami)

## 📄 License

This project is licensed under the MIT License.

---

**Enjoy never forgetting birthdays again! 🎉**
