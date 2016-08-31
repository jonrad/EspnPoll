# EspnPoll

Updates the [CSG spreadsheet](https://www.reddit.com/r/fantasyfootball/comments/4zifas/csg_fantasy_football_spreadsheet_v414/) live from your ESPN draft

## Directions
1. Make sure to have the [CSG spreadsheet](https://www.reddit.com/r/fantasyfootball/comments/4zifas/csg_fantasy_football_spreadsheet_v414/) open in excel
2. Open the "Lite Draft Application" in your ESPN draft. 
3. View source (right click and choose view source). 
4. Find a line that looks like: 	*draftToken: "1:1234567:1:12345678:123456789",* (around line 12)
5. Copy the part in the quotes (*1:1234567:1:12345678:123456789*)
6. Run this application and paste when asked for the token

## Known Issues

* Only works with the "Lite Draft Application"
* Only works for publicly viewable leagues (Should be a fairly simple fix to make it work with private leagues)
* Defenses aren't marked off accordingly (Another simple fix)
* If you reorder the spreadsheet, things will get marked off incorrectly (Hiding rows is fine though). Another fairly simple fix
