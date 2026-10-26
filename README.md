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
- **Schedule**: the week (Sunday to Saturday) by day and time, from the Appt List tab, with open template slots and holds.
- **Index planner**: each provider's template, and the index visits for each new start, proposed from it, held, then booked in
  NextGen (below).
- **Patients**: everyone on the Master Patient Tracker, sorted by what they need: auth ended or ending, visits left on the auth,
  nothing booked, protocol not confirmed. Each need has the action that clears it.
- **Consults**: the Consults tab, approved-but-not-scheduled first.
- **Settings**: your first name, the NextGen upload, Teams workflows and message wording, the chair map, timing, dark mode,
  workbook setup.

## The visit, at the desk

1. **Check in.** Pick a chair or leave them in the lobby. The RN's chat gets "SMIJAN is here for their 1:30 PM IVK visit with
   Dr. Rivera. Chair 2." If the patient needs booking out or a protocol confirmation, the group chat gets the ask at the same time.
2. **Seat.** Picking a chair later tells the RN which one.
3. **Ask.** While they're in the chair, the card's next step is to ask the group chat (if it wasn't sent at check-in). The message
   can be edited before it goes.
4. **Check out.** Say whether the next visit is booked, book later, or not needed. Ticking "The provider confirmed the protocol"
   writes Protocol Updated (today) and, if it changed, Frequency on the Master Patient Tracker.

A patient needs booking when they're active (Ongoing or Induction), not PRN, and booked less than 6 weeks ahead. The protocol needs
confirming when Protocol Updated is blank or older than 90 days. Both are in Settings.

## Index planning

- **Template**: each provider's week on the **IVK Template** tab, laid out like the clinic's provider templates (Location,
  Provider, Last Modified, Effective From, half hours from 8:00 AM, an Event Types tally of live COUNTIF formulas beside it), so
  the Clinic Template Tool can read it too. Cells are coloured by slot type with conditional formatting and have a slot-type
  dropdown, so the tab can be edited in Excel as well. In the add-in: Index planner › Template, drag across half hours, pick a type.
  Add a provider from the NextGen report's providers, for the days they're in clinic, empty or copied from another grid; a later
  Effective From records a revision.
- **Slot types** (the **Slot Types** tab, the Clinic Template Tool's columns plus IVK use and colours): what each type is for in
  IVK, and the NextGen events it takes (chosen from the events on the report).
  - *IV Ketamine Start/Index*: an infusion start. It's an **index start** only when an IVK Post Check or Flex Post Check sits
    90 minutes later: the provider spends the first 30 and the last 30 minutes with the patient, for every index visit,
    including starts. Follow-ups can book these times second; they don't need the post check.
  - *IVK Post Check*: kept for the last 30 minutes of an index visit. *Flex Post Check*: the same, or a TMS follow-up, tele or
    other 30-minute visit outside IVK when no index patient needs it.
  - *IV Ketamine Follow Up*: an infusion start without a post check. Every other type shows plainly on the planner.
- **Requires post-check**: when a provider asks for a post check on a patient's visits (acuity), tick it on the patient's card.
  Their visits then use the post check at their end, and the card lists the next open starts that end on a free post check.
- **What's open**: a start is open when no IVK visit or hold has it and a chair is free for the 2 hours; it's an open index
  start when its post check is free too. Index visits (treatment starts, Induction patients), holds and patients who require a
  post check use the post check. The wider NextGen report's other visits (telehealth, TMS, consults) take the half hour they
  start in, so a flex post check already booked for tele isn't offered for index. NextGen post check holds don't count.
- **Protocol**: 2 visits in each 7 days for 3 weeks unless the patient's own protocol is set. Week 1 starts on treatment 1's day
  (a Wednesday start makes week 1 Wednesday to Tuesday). Visits are never back to back, no more than 5 days apart when the open
  slots allow it, and as evenly spaced as they can be; treatment 1 moves a day or two later if starting on the first open day
  would force a longer gap. Starts ending on an IVK post check are used before ones ending on a flex post check. After index:
  1x per week. All of it is in Settings › Index protocol.
- **Plan**: the Access week shows every template slot as a faint tile: open index starts, follow-ups, post checks, what's booked
  and held. Pressing an open index start makes it treatment 1 (or the treatment being changed); any visit can also be moved
  from its Change list. The page says how many more patients could start index in the next 4 weeks on what's left.
- **Holds**: Hold saves the times straight to the workbook, so everyone sees them on the Schedule, the Access week and the Index
  Planner tab. Book them in NextGen and press Booked in NextGen within 10 minutes; after that the hold turns red. Releasing frees
  the times.
- **Upload**: run the NextGen report for the IVK providers' whole schedule. IV ketamine visits go to the Appt List tab; the rest
  are kept on the hidden tab with no patient on them (date, time, provider, event, status). The Index Planner tab's grid is then
  rewritten as plain values; the lookups into the other workbook are gone.

## Where things are kept

The clinic's tabs are found by sheet name: Master Patient Tracker, Appt List, Consults (and Access R5:R14 for the frequency list).
The add-in adds one hidden tab, **IVK Tracker Data**, when someone presses Set up this workbook (a workbook set up before the
Index planner gets the newer tables from Update the workbook):

| Table | What |
|---|---|
| tblIvkVisits | One row per visit the desk touched: arrived, chair, seated, out, booked, asked, protocol, who, when |
| tblIvkConfig | Clinic, chairs, visit length, thresholds, Teams on/off |
| tblIvkRoutes | arrival → the RN's chat, protocol → the Castro IVK group chat, each with its Teams workflow URL |
| tblIvkTemplates | The message wording |
| tblIvkLog | Every Teams post and its result |
| tblIvkPlans | One row per patient being planned: preferred start, protocol, availability |
| tblIvkHolds | Every hold: patient, date, time, provider, treatment, held / booked / released, who and when |
| tblIvkOther | The providers' other NextGen visits from the wider report: date, time, provider, event, status (no patient) |
| tblIvkFlags | Patients who require a post-check |

The provider templates are on two visible tabs: **IVK Template** (the grids) and **Slot Types** (tblIvkSlotTypes).

The chair map is a document setting (`ivkBayPlan`), the same file format as ESK Tracker's bay map.

## Teams

Each chat gets a workflow made from Teams' **Send webhook alerts to a chat** template (who can trigger it: Anyone). The add-in
posts an Adaptive Card to it, the same way ESK Tracker does (`notify.js`). A workflow URL works like a password: paste it only into
Settings › Teams, never into chat, email or a document.

## Install

Download `manifest.xml`, then in Excel on the web: Home › Add-ins › More add-ins › My add-ins › Upload my add-in.

## Code

`template.js` (provider templates in the Clinic Template Tool's format: parse, lay out, tally; meant to be shared with that tool), `core.js` (rules, no Office.js, tested), `notify.js` (Teams), `excel.js` (workbook reads and writes, task pane only), `app.js` (screens;
full screen is a floating Office dialog that asks the pane for every write), `bayplan.js` (byte-identical copy of the Bay Map Builder's
`render.js` shared block), `map.css` (the ESK Tracker's map rules).

```
node --test tests/*.test.js
BM_BUILDER=../bay-map-builder/render.js BM_ESK=../esk-tracker/src/taskpane.html node --test tests/*.test.js
```
