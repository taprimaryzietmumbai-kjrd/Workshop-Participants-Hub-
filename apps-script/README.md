ZIET Mumbai — Hostel Preference & Arrival Facilitation patch
===============================================================

What this patch changes
-----------------------
- Adds the 52-record PRTM Mumbai participant roster to the existing Apps Script workflow.
- Keeps Ahmedabad's participant list and submissions in their existing sheets.
- Loads Mumbai participants when the Mumbai forms are opened with region=mumbai.
- Saves Mumbai hostel preferences to "Hostel Room Preferences - Mumbai".
- Saves Mumbai arrival plans to "Arrival Facilitation - Mumbai".
- Adds a Region field to submission records.
- Mumbai participants are asked to supply missing private details (email, gender, DOB,
  disability status) in the form. These are not guessed or copied from another region.

Important deployment steps
--------------------------
1. Open the Google Apps Script project:
   https://script.google.com/u/0/home/projects/1IJlCxNAg59Sis_oilmUG-P2GiuzgB8i_ROd5ZZJdwU18yWrx4UlUJoCV/edit

2. Replace the contents of Code.gs with the Code.gs file in this package.

3. Update the project's HTML files using the revised copies in the GitHub repository:
   - https://github.com/taprimaryzietmumbai-kjrd/Workshop-Participants-Hub-/blob/main/Hostel.html
   - https://github.com/taprimaryzietmumbai-kjrd/Workshop-Participants-Hub-/blob/main/Arrival.html

4. Save the project. The first Mumbai roster request uses UrlFetchApp to retrieve the
   public JSON roster from the repository. Google Apps Script may ask you to authorize
   the required external-request permission.

5. Go to Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
   Keep the existing deployment URL. The Mumbai hub links already include:
   ?page=hostel&region=mumbai
   ?page=arrival&region=mumbai

6. Test both Mumbai forms and confirm that new records appear in the two Mumbai-specific
   tabs in the configured spreadsheet. Also test an Ahmedabad form submission to confirm
   its existing workflow remains intact.

Notes
-----
- The source JSON contains 52 Mumbai participant records and is hosted here:
  https://github.com/taprimaryzietmumbai-kjrd/Workshop-Participants-Hub-/blob/main/data/prtm-mumbai-participants.json
- Some source records have blank Employee Codes or "99999". The form allows the participant
  to edit/enter their Employee Code; please verify these values against the official roster.
- The backend uses the existing spreadsheet configured in Code.gs. It does not change
  the master participant sheet or infer missing private details.
- This patch has been syntax-checked, but it has not yet been deployed to the live Apps
  Script project. The deployment steps above are required before live use.
