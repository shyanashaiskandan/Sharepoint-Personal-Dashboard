# SharePoint Personal Dashboard

A SharePoint Framework (SPFx) solution with **two web parts** and **one extension**:

- **Calendar (Web Part)** – Shows upcoming meetings for the next 24 hours 
- **ToDoApp (Web Part)** – Lists all current upcoming tasks from a SharePoint list on your site.
- **MeetingNotification (Extension)** – Displays a top banner 5 minutes before a meeting starts to remind you to join.

---

## Prerequisites
- Microsoft 365 tenant with a SharePoint App Catalog
- SPFx development environment (Node.js, Gulp, Yeoman)
- Microsoft Graph permission **Calendars.Read** approved in the tenant

---

## Local Development
```bash
git clone <repo-url>
cd sharepoint-personal-dashboard
npm install
gulp trust-dev-cert   # first time only
gulp serve
