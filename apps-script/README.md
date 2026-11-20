ZIET Mumbai — Hostel Preference & Arrival Facilitation
=======================================================

Current Apps Script source
--------------------------
- Code.gs: https://github.com/taprimaryzietmumbai-kjrd/Workshop-Participants-Hub-/blob/main/apps-script/Code.gs
- Hostel.html: https://github.com/taprimaryzietmumbai-kjrd/Workshop-Participants-Hub-/blob/main/Hostel.html
- Arrival.html: https://github.com/taprimaryzietmumbai-kjrd/Workshop-Participants-Hub-/blob/main/Arrival.html

Required private roster file
----------------------------
Add the supplied MumbaiRoster.gs as a separate .gs file in the Apps Script project.
It contains the Mumbai participant roster and is intentionally NOT committed to this
public GitHub repository because it includes personal data. The server-side Code.gs
calls getMumbaiParticipantOptions_() from that file.

What the integration does
-------------------------
- Mumbai forms load the 52 participants from the private MumbaiRoster.gs file.
- Available employee code, email, gender, DOB and school values are prefilled.
- Source blanks remain blank; participants may complete editable fields where the form
  permits. No missing personal details are guessed.
- DOB values are normalized for the date field; blank DOBs are not fabricated.
- Mumbai hostel submissions go to "Hostel Room Preferences - Mumbai".
- Mumbai arrival submissions go to "Arrival Facilitation - Mumbai".
- Ahmedabad continues to use the existing PRTM ADI master-sheet roster and its original
  response tabs.
- Both forms use region-aware template rendering and the Refresh Roster handler.

Deployment
----------
1. Open the existing Apps Script project:
   https://script.google.com/u/0/home/projects/1IJlCxNAg59Sis_oilmUG-P2GiuzgB8i_ROd5ZZJdwU18yWrx4UlUJoCV/edit
2. Replace Code.gs with the current apps-script/Code.gs file.
3. Add MumbaiRoster.gs as a separate script file and paste the supplied private roster.
4. Replace Hostel.html and Arrival.html with the repository versions linked above.
5. Save all files, then Deploy > Manage deployments > Edit > New version > Deploy.
6. Test both Mumbai forms and one Ahmedabad form. Confirm that each submission lands in
   its correct regional response tab.

Verification status
-------------------
The repository source has been statically checked for the region-aware doGet routes,
form-to-server function names, response-tab routing, and the Mumbai roster helper call.
The live Apps Script deployment has NOT been updated or end-to-end tested from this
repository. Do not treat the web app as live-updated until a new Apps Script deployment
has been saved and tested.
