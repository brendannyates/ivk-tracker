# IVK Tracker

Excel on the web add-in for the IV ketamine front desk, piloted on the Castro IVK Patient Tracker. It answers one question for the
person at the desk: who's here, which chair are they in, and what do we need from the providers before they leave?

The RN takes vitals and doesn't use the tracker, so the tracker tells the RN in Teams instead. The desk's usual ask to the
providers, "Can we please confirm every 2 week protocol for SMIJAN and send to the desk at discharge for scheduling?", is
written for them and posted to the Castro IVK group chat in one click.

The hosted files hold no patient data. Patients only ever appear as LASFIR codes (Jane Smith is SMIJAN), on screen and in Teams.

## Screens

The rail: Today, Schedule, Access, Patients, Availability, then Settings at the bottom (above Exit full screen in full screen). In
the task pane, full screen sits at the top right of every page with refresh and dark mode.

- **Today**: In treatment, Checked in, Still to come and Checked out, and one visit card with the next step. A visit is
  **Checked in**, then **In treatment**, then **Checked out**; waiting time isn't tracked, so nothing on Today is red. The chair map
  (a file from the [Bay Map Builder](https://brendannyates.github.io/bay-map-builder/)) only shows once one is loaded in Settings;
  with no map, visits are tracked without chairs. **New patients to add** lists everyone on the
  appointment list in the next 4 weeks with no row on the Master Patient Tracker, treatment starts first.
- **Schedule**: the week (Sunday to Saturday) by day and time, from the Appt List tab, with open template slots and holds.
  Treatment starts carry a blue tab and index visits (Induction patients' visits, and holds) a lavender one; **PC** marks a patient
  who requires a post-check, and a shaded tile shows where each post check lands (hover a visit and its post check lights up). A
  visit booked off its template slot (a start in a follow-up-only slot, or a time with no IVK slot) has a small **(i)**: hover or
  press it for "This visit does not match the templated slot-type". Pressing a booked visit opens the patient view; pressing an
  open slot holds it for a patient (an open index start can plan all 6 index visits from there instead).
- **Access**: four pages, picked with the chips at the top.
  - **Demand**: how far out active patients are scheduled, and what's still to book. **Week by week** (Sunday to Saturday, every
    active patient, not paused): **Avg Need** at current protocols (Induction at 2 a week until treatment 6, then the after-index
    protocol, Settings, 1x per week; PRN and blank none), **Booked** (IVK visits and holds on the list) out of the template's IVK
    slots, **Scheduled Through** (the share of active patients with a protocol, not PRN or paused, booked that week or later), and
    **Projected Need** (the visits still to book if each patient keeps their cadence past their last booked visit; index patients
    weekly after treatment 6). Under each week, small pills count the **consults** (IVK consults at this clinic on the wider
    NextGen report, kept on tblIvkOther with their location) and the treatment **starts** on the list that week, cancelled and
    no-shows aside. The tiles lead with the share scheduled out 4 weeks; the clinic decides what share it wants before
    starting new patients. **Approved not scheduled** counts Consults tab rows whose stage says Approved (not "Not Yet Approved")
    and whose patient hasn't started, noting any that already have a start booked; **Consults scheduled** counts IVK consults
    from today on. **By protocol** below keeps the Unscheduled only / Everyone / Scheduled only filter and Include paused.
    **No-shows and late cancels** lists patients by count in the last 30, 60 and 90 days (each one, with its exemption, folded
    away below). **Projected Access Post-Index** lists each index patient with the week they step down and the visits a week
    that frees.
  - **Waitlist** (below), **Index planner** (below) and **Templates** (the provider grids; the providers are toggled at the top
    right of the template card).
- **Patients**: a table of everyone on the Master Patient Tracker that sorts by any column, ascending or descending (press the
  header again): patient, protocol, status, protocol last confirmed, booked through, needs. Chips: Needs attention, New patients
  (on the list, not on the tracker), Active, Paused, Discontinued, All. Pressing a patient opens the **patient view**, laid out as
  the TMS Tracker's: the lead need and
  its action, today's visit if they're here, then one section at a time: **Schedule** (visits with start and index tabs, holds,
  printouts, availability), **Protocol** (frequency, last confirmed and every confirmation since, index progress, post check,
  Oakland), **Benefits** (who covers the provider visit and the infusion, the latest Sohar benefits, the authorization, the BIDF)
  and **Financial** (below). Every field on the tracker row can still be edited (**Edit tracker fields**). Picking a patient
  selects their row in the workbook behind the add-in (Settings › Appearance › Follow along turns it off).
- **Availability**: every patient's availability as the IVK Patient Availability Form asks it (each weekday AM and PM: Preferred,
  Mostly, Sometimes, Not available; effective and expiry dates), with how many patients can come each half day above. Needs
  updating: no form on file, expired, only the tracker's AM/PM columns, or fewer than the 3 days the form asks for.
- **Settings**: your first name, the NextGen upload, Teams workflows and message wording, the chair map, timing (visit length,
  chairs in the room for open slots when there's no map, booking and protocol rules), dark mode, workbook setup.

On open, a loading screen ticks off each read (Connecting, Master Patient Tracker, Appt List, IVK Tracker Data, the chair map).
Refresh dims the page while it reads and then says how many visits and patients it read.

## Patient profiles

A patient on the Master Patient Tracker has their profile in its columns. New patients (on the appointment list, especially a
treatment start, with no row on the tracker) are prompted on Today, Patients and their visit card. **Create profile** asks for the
LASFIR code, status (Induction for a new start), protocol, who covers what
(insurance for both, self pay for both, or one of each), the payer and what the patient pays each visit for the **provider visit
(99214/99215)** and for the **infusion**, and the latest Sohar benefits (payer, plan, network, coverage dates, copay, coinsurance,
deductible and OOP max with what's left, status, verified date, reference). The Sohar fields can be typed or filled from the Sohar
Eligibility add-in's benefits CSV by MRN. Settings › **Sohar benefits** › Import Sohar benefits takes the add-in's whole IVN export (pick
this clinic there) and fills every matched profile at once: patients on the tracker or the appointment list, by MRN; a newer saved
verification is kept, the same date and ref is left alone, MRNs not here are counted, not added (`core.planSoharImport`, `excel.upsertMany`). Saving adds the patient to the Master Patient Tracker (Patient Name is the code; Copay,
Infusion, Primary Encounter Insurance, Infusion Insurance, Status, Frequency, Primary Location, Tx Start Date; the table's formula
columns fill themselves) and the rest to tblIvkProfile. For a patient already on the tracker, Benefits › **Edit coverage and
benefits** opens the same form.

## Index and after

A patient whose status is **Induction** is an index patient: their visits carry the lavender tab, they count at 2 a week in
Demand, and they're the ones listed to change in the Index planner. Once they've kept their 6th visit (cancelled visits and
no-shows don't count), check-out of that visit offers **Move to Ongoing** with the protocol after index (usually 1x per week; the
provider sets it by acuity), and until it's done the patient has the need "Index done: move to Ongoing and confirm the protocol".
Moving writes Status Ongoing, Frequency and Protocol Updated on the Master Patient Tracker.

**Start paperwork** (treatment start visits only) and **Start email** (only before the start’s day, so on the patient view) don’t show for follow-ups. Start paperwork prompts for the BIDF: it's completed by the benefits team and printed from NextGen, so
the add-in asks the desk to print it and saves the day they tick it. The start email still needs the BIDF uploaded as a PDF.

## The visit, at the desk

1. **Check in.** With a chair map loaded, pick the chair (or No chair yet). The **Copay check** is asked out loud: tick
   **LumaPay card on file** and **Verbal consent to collect copay at time of service**. Check-in still works with either unticked;
   the toast and the visit card then say "Not confirmed". Payment can be recorded here too (optional). The RN's chat
   gets "SMIJAN is here for their 1:30 PM IVK visit with Dr. Rivera." (plus "Chair 2." with a map). If the patient needs booking
   out or a protocol confirmation, the group chat gets the ask at the same time.
2. **Start treatment.** Straight away with no map; with a map and no chair yet, it asks which chair and can tell the RN.
3. **Ask.** While they're in treatment, the card's next step is to ask the group chat (if it wasn't sent at check-in). The message
   can be edited before it goes.
4. **Check out.** Say whether the next visit is booked, book later, or not needed. Ticking "The provider confirmed the protocol"
   writes Protocol Updated (today) and, if it changed, Frequency on the Master Patient Tracker.

A patient needs booking when they're active (Ongoing or Induction), not PRN, and booked less than 6 weeks ahead. The protocol needs
confirming when Protocol Updated is blank or older than 90 days. Both are in Settings.

## Payments

Each IVK visit owes the provider visit (the tracker's Copay) and the infusion (Infusion), collected separately. Check-out asks
**Collected** or **Not collected** for each part that owes something and has no answer from check-in; what the patient owes today can be changed there (and, ticked, saved as their Copay and Infusion on the tracker). **Record / Change today’s payment** on the visit card changes the answer any time, after check-out too (a new row; the latest answer counts); what isn't collected becomes the patient's balance. Payment
rows are kept per part (the visit's key with |prov or |inf); rows saved for the whole visit before the split still count. On the
patient's Financial page: the balance and each missed part (**Collected**, **Waive**), past visits nobody answered (since the
payment log started), **LumaPay card on file** and **Verbal consent to collect copay at time of service** (the patient's
standing answer; ticking either at check-in sets it), and **Record a payment
or reminder**: a payment on account (comes off the oldest balance first) or a balance reminder sent (date, In-person or Luma). A
balance with no reminder in 14 days is a need.

**Copay check, per visit.** Each check-in adds one tblIvkPay row with Kind `consent`, Key = the visit's key, Amount 0, and
Method `card+consent`, `card`, `consent` or `none`, with By and At. Changing it later (Copay check and payment on the visit
card) adds a new row; the latest one counts. No consent row for a visit means nobody did the check. To list a day: filter
tblIvkPay on Date and Kind = consent; visits on the Appt List with no row were checked in without it (or not checked in).
Consent rows never count toward balances or collected totals.

## Transfers

A patient whose Primary Location isn’t this clinic (or an Oakland patient) has **Transfer** on the patient view’s Protocol page: **Discontinue now** (Status Discontinued, a dated transfer note in Clinic Notes, a log line) or **Schedule a transfer** for a future date (tblIvkTransfers). A scheduled transfer is confirmed at check-out of their last visit here; visits booked here after the date are a need. Transferring to Oakland can mark OAK tasked to schedule them.

## Cancellation waitlist

A patient can be waitlisted for a sooner visit or a day and time they'd rather have, from a visit on the Schedule or from the
patient view (Schedule › Cancellation waitlist). The form takes the visit to move (or none), the earliest and latest dates (latest
blank = the day before the visit, so sooner only; later = a better time, not just sooner), weekdays and halves (from their
availability), particular dates ("Thursday the 13th", AM/PM or one time), a start-time window, and whether only their provider will
do (tblIvkSooner, one row per patient; the Days cell holds the tokens, see core.js). Days they already come in are never offered.

Openings are open template slots and times a cancellation freed. Each NextGen upload pops up the openings nobody has been shown yet;
cancellations that fit also go on Today's **Cancellation openings** card (tblIvkOffers). **Offer** opens the call: **They'll take
it** (then the NextGen steps: book the new time, cancel the old visit; **Done in NextGen** takes them off the waitlist), **Not that
one**, or **No answer** (stays on Today to try again). When more than one patient fits a time, the card lists them by who waitlisted
first. The next upload confirms the move, or says the new visit isn't on it, or that the old visit still needs cancelling.

## Index booking

A treatment start for an Induction patient still at 2x per week (or no protocol yet) needs the other 5 index visits booked behind it; until they’re on the list the visit card, Today’s row and the patient carry **Index not fully booked**.

## Oakland patients

Ticked on the patient (or by an Oakland home clinic): their index, the first 6 treatments, is here at Castro, then they go back to
Oakland for maintenance. From treatment 3 the need is **Task OAK to book maintenance** until **Mark OAK tasked**; after treatment 6,
visits still booked here are flagged. Treatments are counted from the Tx Start Date on the Appt List, so a course that started
before the list begins isn't counted (the card says so).

## Availability

The patient availability form's levels for each weekday morning and afternoon (Preferred, Mostly, Sometimes, Not available),
effective and expiry dates, a note and dates away. Preferred and Mostly go to the tracker's AM/PM columns and the Index planner;
the ten AM/PM columns are found by position (the first ten headers reading AM or PM, Mon AM to Fri PM), so a mistyped
header still reads; dates away are skipped by the planner and a visit booked on one is a need. **Print availability form** prints the form with the
patient's name and date of birth.

## Printing and the start email

`forms.html` opens over the add-in (pane or full screen). The patient's name and DOB go to it by message and are never saved.

- **Print schedule**: the TMS Tracker's patient schedule (its schedule.html layout and PDF: logo, "As of" the list's run time, the NextGen
  columns MRN to Status, landscape), for Next 14 days, Next 30 days, Rest of course or From/To. The start email uses the same table.
  It's on the Index planner card (once a patient has holds), the patient card and the visit card (in the chair or checked out). A new
  patient who isn't on the Master Patient Tracker yet prints from their MRN; their name is typed for the consent or email, never saved. **Print this week** (Schedule): the clinic's week, codes only.
- **Print start forms** (start visits, Induction patients): the schedule and the IV ketamine consent with the patient's name, DOB, the
  clinic location and the clinician filled in. The consent is signed in person on arrival, so it's printed, never emailed. Consent wording is verbatim from `assets/ivk-consent.pdf`.
- **Start email**: the desk's "Upcoming Appointment(s)" email (Settings › Messages, HTML), subject "SECURE: Upcoming Visit
  Information | MHS", with the MRN, first name, clinic, first visit, clinic address and phone and the schedule table filled in. No
  signature is built in: Outlook adds the sender's, or they add one under Settings › You ({signature} in the template, else the end).
  A checklist (patient's email address, CC set, clinic details entered, template, schedule, BIDF attached in the window; BCC optional) has to be
  complete before **Outlook draft (.eml)** unlocks; the draft then has To, CC, BCC, the email and every attachment ({code} Appt
  {code} schedule m.d.pdf, the BIDF and any extras; no consent). CC, BCC, the clinic address and phone are in Settings › Messages.
- The date of birth comes from a DOB column on the tracker if there is one; otherwise the desk types it for the forms.

## Index planning

- **Template**: each provider's week on the **IVK Template** tab, laid out like the clinic's provider templates (Location,
  Provider, Last Modified, Effective From, half hours from 8:00 AM, an Event Types tally of live COUNTIF formulas beside it), so
  the Clinic Template Tool can read it too. Cells are coloured by slot type with conditional formatting and have a slot-type
  dropdown, so the tab can be edited in Excel as well. In the add-in: Index planner › Template, drag across half hours, pick a type.
  Add a provider from the NextGen report's providers, for the days they're in clinic, empty or copied from another grid; a later
  Effective From records a revision.
- **Time off and coverage** (Index planner › Template): a provider's PTO by date is an exception on their template: their slots close on
  those days (Access week, Schedule, every plan). Pick who covers, and the add-in asks whether the covering provider works the absent
  provider's template (pre-filled, replacing their own that day, and the absent provider's visits and holds sit in it) or keeps
  their own. Visits or holds still with the absent provider are listed to move in NextGen. Saved on tblIvkExceptions.
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
- **Index-eligible**: an open index start only counts (on the Schedule, the Access week and in "room to start") when all 6 index
  visits still fit after it in their windows: 2 in each of the 3 weeks, 2 to 5 days apart. **Next index-eligible starts** stand on
  the right of the planner. Pressing one makes it treatment 1 for the patient picked, or, with nobody picked, starts a draft plan
  that can be explored and changed and is held for whoever it's for (someone on the waitlist or the planner, a treatment start
  on the list, or a new LASFIR and MRN). Pressing the picked patient again plans without one. Hovering a start on the Access week
  shades its post check.
- **Protocol**: 2 visits in each 7 days for 3 weeks unless the patient's own protocol is set. Week 1 starts on treatment 1's day
  (a Wednesday start makes week 1 Wednesday to Tuesday). Visits are never back to back, no more than 5 days apart when the open
  slots allow it, and as evenly spaced as they can be; treatment 1 moves a day or two later if starting on the first open day
  would force a longer gap. Starts ending on an IVK post check are used before ones ending on a flex post check. After index:
  1x per week. All of it is in Settings › Index protocol.
- **Plan**: the Access week shows every template slot as a faint tile: open index starts, follow-ups, post checks, what's booked
  and held. Pressing an open index start makes it treatment 1 (or the treatment being changed); any visit can also be moved
  from its Change list. The page says how many more patients could start index in the next 4 weeks on what's left.
- **Who is planned**: a new index patient isn't on the Master Patient Tracker yet, so they're added by **LASFIR and MRN** (New patient);
  approved consults waiting for a start show too. **Change the visits of a patient in index** lists only patients already in index
  (Induction, or started within the index weeks), for changing their schedule.
- **Confirmed by NextGen**: after **Booked in NextGen**, the visits stay "booked: upload to confirm" until a NextGen report uploaded
  after the booking shows them. Visits that report doesn't have turn red ("not on the NextGen list"), with Release. The Index planner
  banner and badge carry both, and the upload says how many it confirmed.
- **Holds**: Hold saves the times straight to the workbook, so everyone sees them on the Schedule, the Access week and the Index
  Planner tab. Book them in NextGen and press Booked in NextGen within 10 minutes; after that the hold turns red. Releasing frees
  the times.
- **Upload**: run the NextGen report for the IVK providers' whole schedule. IV ketamine visits go to the Appt List tab; the rest
  are kept on the hidden tab with no patient on them (date, time, provider, event, status). The Index Planner tab's grid is then
  rewritten as plain values; the lookups into the other workbook are gone. Older exports read too: .xls (Excel 97-2003, or a web
  page saved as .xls), Excel 2003 XML, CSV or text; the report on any sheet; dates written out; the time inside the date-time.
- **Other clinics**: the report can include the neighbouring clinics. The upload asks which locations are this clinic's (kept as
  the Locations setting); only those go on the Appt List tab. The other clinics' IVK visits go to tblIvkElsewhere (LASFIR, no
  names) and are used only to check a transfer or Oakland hand-off is booked there: "nothing booked there yet" is a need, a
  booked visit a quiet tag. Nothing is said about other clinics until a report with them has been uploaded.
  Primary Location (Edit) and Transfer to offer the report's locations as NextGen names them. The patient view's Schedule page shows
  each visit's location, their appointments here (kept and booked) and any booked at other clinics.

## Who's using it, the clinic, admin

- **People** register on the workbook (tblIvkUsers). The first time the tracker opens each day on a computer, it asks **Enter your name** (anyone can be using it: desk, RN or
  provider) from that workbook's list only, nobody picked; the people who've used that computer on that workbook are listed first ("On
  this computer"). "I'm not on the list" adds a name. The pick lasts the day (Switch user in the rail or Settings › You). This computer
  remembers picks per workbook (localStorage `ivkWho:<workbook key>`, from the file URL or the workbook's `ivkWbId` setting), so a
  name never carries from one clinic's workbook to another. Full screen uses the pane's pick.
- **The clinic** is the workbook's saved Clinic, else the name the workbook starts with ("Castro IVK Patient Tracker" → Castro), else
  Today asks for it once. Nothing clinic-specific is built in. Clinic, clinic address, phone, CC and BCC are saved by whoever enters
  them first; after that they're read-only until **Admin** is unlocked (Settings › Admin, 15 minutes on that screen). Admin also takes
  people off the list. The code is checked as a salted hash, but the repo is public: it keeps settings from being changed by accident,
  it isn't a security control.
- Castro only: its workbook relied on the address, phone, CC and BCC that used to be built in; they're saved into Castro's own
  tblIvkConfig once, on open (`carryCastro` in app.js, to be removed after).

## Where things are kept

The clinic's tabs are found by sheet name: Master Patient Tracker, Appt List, Consults (and Access R5:R14 for the frequency list).
The add-in adds one hidden tab, **IVK Tracker Data**, when someone presses Set up this workbook (a workbook set up before the
Index planner gets the newer tables from Update the workbook):

| Table | What |
|---|---|
| tblIvkVisits | One row per visit the desk touched: arrived, chair (only with a map), seated (when treatment started), out, booked, asked, protocol, who, when |
| tblIvkConfig | Clinic, chairs, visit length, thresholds, Teams on/off, start email address / phone / CC / BCC |
| tblIvkRoutes | arrival → the RN's chat, protocol → the clinic's IVK group chat, each with its Teams workflow URL |
| tblIvkTemplates | The message wording |
| tblIvkLog | Every Teams post and its result |
| tblIvkPlans | One row per patient being planned: preferred start, protocol, availability |
| tblIvkHolds | Every hold: patient, date, time, provider, treatment, held / booked / released, who and when |
| tblIvkOther | The providers' other NextGen visits from the wider report: date, time, provider, event, status (no patient) |
| tblIvkFlags | Patients who require a post-check |
| tblIvkPatient | Per patient: LumaPay card on file, verbal consent to collect copay, Oakland, OAK tasked, availability, dates away, waitlist order and waiting on |
| tblIvkExceptions | Provider time off: provider, dates, who covers, whether they work the absent provider's template |
| tblIvkTransfers | Transfers to another clinic: to, date, scheduled / done / cancelled, note, confirmed |
| tblIvkSooner | The cancellation waitlist: visit to move (or none), earliest date, what they want (days, dates, times, latest date), same provider, status, openings already shown |
| tblIvkOffers | Openings offered to waitlisted patients: patient, opening, the visit it replaces, outcome (to offer, no answer, declined, accepted, moved), note, by |
| tblIvkElsewhere | IVK visits at the other clinics from a wider report: MRN, LASFIR, date, time, provider, event, status, location |
| tblIvkProfile | Per patient: insurance or self pay and the payer for the provider visit and the infusion, the latest Sohar benefits, BIDF printed |
| tblIvkPay | Payments due at each visit (collected / missed), later payments and waivers, balance reminders, the copay check at check-in (consent) |
| tblIvkUsers | The people who use this workbook (the "Enter your name" list) and their own settings (an optional email signature); no patient |

The provider templates are on two visible tabs: **IVK Template** (the grids) and **Slot Types** (tblIvkSlotTypes).

The chair map is a document setting (`ivkBayPlan`), the same file format as ESK Tracker's bay map.

## Teams

Each chat gets a posting flow built from blank in Power Automate (trigger: When a Teams webhook request is received, who can
trigger it: Anyone; a flow made from the Teams template stamps "used a Workflow template" under every card). The add-in posts a
plain Adaptive Card in the ESK Tracker's layout (`notify.js`): the message in bold, then a line per fact, no buttons. When the
arrival and the ask go to the same chat, check-in sends one card that tags the RN and the provider. Each visit's card collapses
to one line 2 hours after check-in: the payload carries the collapsed card and the minutes to wait, and the flow waits and updates
the card. Build sheet: OneDrive, Add-In Manifests › IVK Tracker › "IVK Tracker - Teams cards build sheet.html". A flow URL works
like a password: paste it only into Settings › Teams, never into chat, email or a document.

## Install

Download `manifest.xml`, then in Excel on the web: Home › Add-ins › More add-ins › My add-ins › Upload my add-in.

## Code

`template.js` (provider templates in the Clinic Template Tool's format: parse, lay out, tally; meant to be shared with that tool), `core.js` (rules, no Office.js, tested), `notify.js` (Teams), `excel.js` (workbook reads and writes, task pane only), `forms.html` (schedule, consent, availability form, start email; jsPDF and pdf-lib from cdnjs), `app.js` (screens;
full screen is a floating Office dialog that asks the pane for every write), `bayplan.js` (byte-identical copy of the Bay Map Builder's
`render.js` shared block), `map.css` (the ESK Tracker's map rules).

```
node --test tests/*.test.js
BM_BUILDER=../bay-map-builder/render.js BM_ESK=../esk-tracker/src/taskpane.html node --test tests/*.test.js
```
