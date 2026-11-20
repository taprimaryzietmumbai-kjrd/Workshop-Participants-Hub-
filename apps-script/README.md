# Mumbai Hostel & Arrival Facilitation integration

## Changes committed
- `data/prtm-mumbai-participants.json`: 52 Mumbai roster records.
- `apps-script/Code.gs`: region-aware backend. Ahmedabad continues to use the existing `PRTM ADI` sheet workflow; Mumbai roster is loaded from the JSON file and submissions are written to separate tabs.
- `Hostel.html` and `Arrival.html`: read the `region=mumbai` query parameter, load the Mumbai roster, allow Mumbai participants to enter missing personal/contact details, and pass the region with each submission.
- `prtm-mumbai.html`: Mumbai buttons now open the region-specific sub-apps.

## Apps Script deployment still required
The repository changes do not automatically update the separately hosted Apps Script project. In project [1IJlCxNAg59Sis_oilmUG-P2GiuzgB8i_ROd5ZZJdwU18yWrx4UlUJoCV](https://script.google.com/u/0/home/projects/1IJlCxNAg59Sis_oilmUG-P2GiuzgB8i_ROd5ZZJdwU18yWrx4UlUJoCV/edit):

1. Replace `Code.gs` with `apps-script/Code.gs`.
2. Replace the Apps Script HTML files `Hostel` and `Arrival` with the corresponding repository files at the root.
3. Save, authorize the external roster fetch if prompted, then Deploy → Manage deployments → Edit → New version → Deploy.
4. Test both Mumbai forms. Confirm submissions create/update `Hostel Room Preferences - Mumbai` and `Arrival Facilitation - Mumbai` in the configured spreadsheet. Test Ahmedabad as well.

Mumbai's published roster does not contain email, DOB, gender or disability details for every participant. The forms therefore ask Mumbai participants to provide those fields; no values are fabricated. The Employee Code field is editable for Mumbai so blank/placeholder source values can be verified by the participant.
