# Project Summary – Member Info Scraper and Auto Emailer

## 📌 Project Overview

This script was developed to **automatically collect member information** from a specific website, organize the data into an Excel file, and **send it via email once the process is complete**. The entire process is automated to save time and reduce manual work.

---

## ✅ What This Script Does

1. **Scrapes Data from a Website**  
   - It opens a webpage that lists member profiles.
   - From each profile, it collects:
     - Full name
     - Email address
     - Phone number
     - City and country
     - Organization and job title
     - Event-related content (link and date)

2. **Stores Data in Excel File**  
   - All collected information is saved into an Excel file.
   - The file is automatically updated during the process.
   - File name: `member_details.xlsx`

3. **Sends Email with Excel File**  
   - Once the data collection is complete, the script automatically sends the Excel file to a predefined email address.
   - This step uses Gmail (or another email service) to deliver the file without manual involvement.

4. **Runs with a Single Command or Automatically**  
   - You can run the script manually.
   - Or, it can be integrated into a scheduler or system that runs automatically on a regular basis.

---

## 🎥 Demo Video

You can watch a short video showing how the script works here:  
👉 [Click to watch the demo](https://www.loom.com/share/dab49efb1bb54e1a96ca20e4f9a3b71b?sid=f1d37322-3dd8-4c38-9542-df9c020f3f2f)

---

## 🧠 Why This Was Built

The goal was to reduce the time spent manually visiting the website, copying information, and sending reports.  
With this script:
- Data is always up to date.
- There are no manual mistakes.
- Reports are sent automatically by email when ready.

---

## 📤 Output Example

The final Excel file includes:

| Full Name | Email | Phone | City | Country | Organization | Job Title | Content Link | Event Date |
|-----------|-------|-------|------|---------|--------------|-----------|---------------|-------------|

---

## 🔐 Security Notes

- The email sending is done securely.
- You can set up your Gmail or use another mail provider.
- The script is configured to send to only trusted email addresses.

---

## 🛠 Technologies Used

- **Python** (Core scripting)
- **Selenium** (For browser automation)
- **OpenPyXL** (To write Excel files)
- **Flask + Flask-Mail** (To send email)

---

If you need this project to run on a schedule (daily, weekly, etc.), it can be extended using a scheduler like `cron` or Windows Task Scheduler.




