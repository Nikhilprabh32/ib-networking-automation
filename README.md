# Networking Email Scheduler (Google Sheets + Apps Script)

## 📌 Overview
This project automates the process of sending personalized networking emails directly from Google Sheets using **Google Apps Script**.  

I built this to help my Investment Banking friends manage outreach more efficiently. Cold emailing is critical for breaking into IB, but it’s time-consuming, repetitive, and easy to lose track of. This tool streamlines the process while ensuring outreach remains **personalized, professional, and respectful of recipients’ time**.  

---

## 🎯 Problem
- Networking outreach is essential for IB recruiting, but sending dozens of emails daily can quickly become overwhelming.  
- Manual scheduling often leads to mistakes (double sending, awkward timing, or missed follow-ups).  
- Students need a structured, disciplined way to send emails consistently **without appearing spammy**.  

---

## 💡 Solution
This script automates the entire workflow:  

- **Smart Scheduling** – Runs **Tue–Fri only**, before 11 AM.  
- **Rate-Limited Sending** – ≤3 emails per sheet per day, with times spaced **≥45 min apart**.  
- **Personalization** – Dynamically inserts `Name`, `School`, and `Company` from the spreadsheet.  
- **Status Tracking** – Marks each contact as *Sent* once delivered.  
- **Daily Report** – Sends a summary email of total outreach sent that day.  

This balances automation with authenticity — keeping outreach manageable while avoiding bulk “spammy” behavior.  

---

## 🚀 Why I Built This
I originally created this for my friends who were preparing for IB recruiting. The goals were:  
1. **Respect professionals’ time** – No blast emails, no awkward timings.  
2. **Increase accountability** – Track progress and avoid missing contacts.  
3. **Save time** – Focus more on conversations and interview prep, less on admin work.  


---

## 📝 Setup Instructions
1. Create a Google Sheet with columns:  
   - `Name`, `Email`, `School`, `Company`, `Status`  
2. Open **Extensions > Apps Script** in Google Sheets.  
3. Paste in the code from `Code.gs`.  
4. Grant permissions for Gmail + Sheets.  
5. Set a daily trigger for `scheduleEmailsAllSheets`.  
6. Add your contacts and track outreach progress!  

---

## 📧 Example Email
Hi John,

I hope you're having a great week so far! My name is Nik, and I am a sophomore finance major at Texas A&M University interested in Investment Banking opportunities at Goldman Sachs.

If possible, would you be open to a brief 10–15 minute call in the coming weeks to speak about your experience at the firm as well as your journey from Duke? I am more than happy to work around your schedule.

In case it's helpful to provide more context on my background, I have attached my resume below for your reference. I look forward to hearing back from you soon!

Best,
Nik

yaml
Copy code

---

## 🛠️ Tech Stack
- **Google Apps Script**
- **Google Sheets**
- **GmailApp (built-in)**  

---

## 👤 Author
**Nik** – Finance Student at Texas A&M | Aspiring Product Manager  
Originally built for friends preparing for IB recruiting — now shared as an example of us
