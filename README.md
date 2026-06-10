# SAVED Group System

A full-stack savings & credit management system for the SAVED group.

## Tech Stack
- **Backend:** Node.js + Express
- **Auth:** JWT tokens + bcrypt password hashing
- **Frontend:** Vanilla JS SPA (no framework needed)
- **Charts:** Chart.js

## Setup & Run

```bash
npm install
npm start
```
Server runs on http://localhost:3000

## Login Credentials

| Name | Email | Default Password |
|------|-------|-----------------|
| **NIRINGIYUMUKIZA Eric (Admin)** | ericniring@gmail.com | **7599** (last 4 of 0794707599) |
| Mukarukundo Philomene | philomenemukarukundo0@gmail.com | 1892 |
| NIYOBUHUNGIRO AMOS | niyobuhungiroamos034@gmail.com | 9381 |
| Twizerimana Francoise | ftwizerimana41@gmail.com | 8422 |
| Byiringiro Bonfils Kevin | byiringirokevin242@gmail.com | 7755 |
| Uwayisaba Agnes | uwayisabaagnes3@gmail.com | 3370 |
| Ishimwe Umulisa Fifi | ishimweumulisafifi@gmail.com | 9121 |
| SHEMA Shalifu | shemashalifu1@gmail.com | 2245 |
| IRADUKUNDA Theopiste | theopisteiradukunda7@gmail.com | 3752 |
| Umwizerwa Christine | umwizerwatina@gmail.com | 6693 |
| Kwizera Divine | kwizeraleadivine@gmail.com | 5125 |
| NGIRUMUKIZA Eric | ericngirumukiza1@gmail.com | 1436 |
| NIYONSHUTI THEONESTE | theoniyonshuti2022@gmail.com | 7782 |
| Mbayeho Happy Selligue | mbayehohappy@gmail.com | 0701 |
| TUYIRAMYE Ancile | anciletuyiramye@gmail.com | 2695 |
| Ingabire Solange | solangeingabire370@gmail.com | 4562 |
| NIWEMAHORO Providence | providenceniwemahoro10@gmail.com | 9573 |
| Muhozi Julia | juliamuhoozi057@gmail.com | 5344 |

## Access Levels

**Admin (Eric):**
- Full dashboard with group statistics
- Member registry with search
- Weekly savings tracker for all members
- Loan management for all 5 active loans
- Analytics charts

**Members:**
- Personal dashboard with their own stats
- Weekly savings history
- Their own loan status
- Change password

## Deployment

This app can run on any Node.js host. Use these production settings:

- Build command: `npm install`
- Start command: `npm start`
- Environment variables:
  - `NODE_ENV=production`
  - `JWT_SECRET=<strong random secret>`
  - `DATA_DIR=<persistent writable folder>`

For Render, this repo includes `render.yaml`. Create a Blueprint from the repo. The free demo config uses:

- `DATA_DIR=/tmp/saved-data`
- No persistent disk
- Auto-generated `JWT_SECRET`

This deploys for free, but saved data can reset after redeploys/restarts. For real use, upgrade to persistent storage and set `DATA_DIR=/var/data`.

For a VPS, copy the files, run `npm install`, then run the app with a process manager such as PM2:

```bash
JWT_SECRET="replace-with-a-strong-secret" DATA_DIR="./data" pm2 start server.js --name saved-system
```
