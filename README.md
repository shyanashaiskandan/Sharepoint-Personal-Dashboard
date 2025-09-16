# SharePoint Personal Dashboard

A SharePoint Framework (SPFx) solution with **two web parts** and **one extension**:

- **Calendar (Web Part)** – Shows upcoming meetings for the next 24 hours <br/>

Outlook Calendar: <br/>
<img width="600" height="376" alt="Screenshot 2025-09-12 at 10 46 57 AM" src="https://github.com/user-attachments/assets/e355f03d-ad57-49e9-87dd-9d79d27b7775" /> <br />

Web Part: <br/>
<img width="314" height="376" alt="Screenshot 2025-09-12 at 10 47 27 AM" src="https://github.com/user-attachments/assets/2cb497dc-1da3-4d88-82e6-7204e690fe49" /> <br/>


- **ToDoApp (Web Part)** – Lists all current upcoming tasks from a SharePoint list on your site.

Sharepoint List: <br/>
<img width="500" height="219" alt="Screenshot 2025-09-12 at 10 51 14 AM" src="https://github.com/user-attachments/assets/8d8b2b43-f954-4556-99db-cba1c38d555c" /> <br/>

Web Part: <br/>
<img width="600" height="267" alt="Screenshot 2025-09-12 at 10 51 38 AM" src="https://github.com/user-attachments/assets/1f50530b-e73b-4288-b422-57a0950f31ea" /> <br/>

Additional Feature: Option to Add Task

https://github.com/user-attachments/assets/9221d5e2-e7e6-4e26-be16-4ba1073fbf32

<br/>

- **MeetingNotification (Extension)** – Displays a top banner 5 minutes before a meeting starts to remind you to join.

Extension: <br/>
<img width="1440" height="163" alt="Screenshot 2025-09-12 at 10 56 04 AM" src="https://github.com/user-attachments/assets/90dba58a-1f3e-4984-917f-929ae7f3fab4" /> <br/>

---

## Prerequisites
- Microsoft 365 tenant with a SharePoint App Catalog
- SPFx development environment (Node.js, Gulp, Yeoman)
- Microsoft Graph permission **Calendars.Read** approved in the tenant

---

## Local Development
```bash
git clone git@github.com:shyanashaiskandan/Sharepoint-Personal-Dashboard.git
cd sharepoint-personal-dashboard
npm install
gulp trust-dev-cert   # first time only
gulp serve
