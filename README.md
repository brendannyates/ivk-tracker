# IVK Tracker

Excel on the web add-in for the IV ketamine front desk, piloted on the Castro IVK Patient Tracker. It answers one question for the
person at the desk: who's here, which chair are they in, and what do we need from the providers before they leave?

The RN takes vitals and doesn't use the tracker, so the tracker tells the RN in Teams instead. The desk's usual ask to the
providers, "Can we please confirm every 2 week protocol for SMIJAN and send to the desk at discharge for scheduling?", is
written for them and posted to the Castro IVK group chat in one click.

The hosted files hold no patient data. Patients only ever appear as LASFIR codes (Jane Smith is SMIJAN), on screen and in Teams.

## Screens

- **Today**: the chair map (a file from the [Bay Map Builder](https://brendannyates.github.io/bay-map-builder/), or chair tiles
  until one is loaded), who's checked in and waiting, who's still to come, and one visit card with the next step. Red only means
  a patient has waited 10+ minutes for a chair.
- **Schedule**: the week (Sunday to Saturday) by day and time, from the Appt List tab. Replaces the Index Planner tab, which reads
  another workbook's Appt List.
- **Patients**: everyone on the Master Patient Tracker, sorted by what they need: auth ended or ending, visits left on the auth,
  nothing booked, protocol not confirmed. Each need has the action that clears it.
- **Consults**: the Consults tab, approved-but-not-scheduled first.
- **Settings**: your first name, the NextGen upload, Teams workflows and message wording, the chair map, timing, dark mode,
  workbook setup.

## The visit, at the desk

1. **Check in.** Pick a chair or leave them in the lobby. The RN's chat gets "SMIJAN is here for their 13:30 IVK visit with
   Dr. Rivera. Chair 2." If the patient needs booking out or a protocol confirmation, the group chat gets the ask at the same time.
2. **Seat.** Picking a chair later tells the RN which one.
3. **Ask.** While they're in the chair, the card's next step is to ask the group chat (if it wasn't sent at check-in). The message
   can be edited before it goes.
4. **Check out.** Say whether the next visit is booked, book later, or not needed. Ticking "The provider confirmed the protocol"
   writes Protocol Updated (today) and, if it changed, Frequency on the Master Patient Tracker.

A patient needs booking when they're active (Ongoing or Induction), not PRN, and booked less than 3 weeks ahead. The protocol needs
confirming when Protocol Updated is blank or older than 90 days. Both are in Settings.

## Where things are kept

The clinic's tabs are found by sheet name: Master Patient Tracker, Appt List, Consults (and Access R5:R14 for the frequency list).
The add-in adds one hidden tab, **IVK Tracker Data**, when someone presses Set up this workbook:

| Table | What |
|---|---|
| tblIvkVisits | One row per visit the desk touched: arrived, chair, seated, out, booked, asked, protocol, who, when |
| tblIvkConfig | Clinic, chairs, visit length, thresholds, Teams on/off |
| tblIvkRoutes | arrival → the RN's chat, protocol → the Castro IVK group chat, each with its Teams workflow URL |
| tblIvkTemplates | The message wording |
| tblIvkLog | Every Teams post and its result |

The chair map is a document setting (`ivkBayPlan`), the same file format as ESK Tracker's bay map.

## Teams

Each chat gets a workflow made from Teams' **Send webhook alerts to a chat** template (who can trigger it: Anyone). The add-in
posts an Adaptive Card to it, the same way ESK Tracker does (`notify.js`). A workflow URL works like a password: paste it only into
Settings › Teams, never into chat, email or a document.

## Install

Download `manifest.xml`, then in Excel on the web: Home › Add-ins › More add-ins › My add-ins › Upload my add-in.

## Code

`core.js` (rules, no Office.js, tested), `notify.js` (Teams), `excel.js` (workbook reads and writes, task pane only), `app.js` (screens;
full screen is a floating Office dialog that asks the pane for every write), `bayplan.js` (byte-identical copy of the Bay Map Builder's
`render.js` shared block), `map.css` (the ESK Tracker's map rules).

```
node --test tests/*.test.js
BM_BUILDER=../bay-map-builder/render.js BM_ESK=../esk-tracker/src/taskpane.html node --test tests/*.test.js
```
