This program is used to transfer dedication text from one year's spreadsheet to the next. The reason this is neccessary is because the dedications are written by users based on the hebrew date, and so we need to move them to the correct hebrew date of the next year. The hitch is that some dates don't have enough space, etc so the solution is that the program creates another spreadsheet with all the problematic ones by problem category (exceeds character limit, not enough lines for that date, etc).

1. Two spreadsheets are required: The previous year's (filled out), and the upcoming year's (blank). The blank upcoming blank one is hardcoded, and the old one is uploaded by the different congregations.
2. They must be formatted the same way (same amount of columns, same column names, etc).
3. The program will assume that the sheet we are working with (on both years' spreadsheets) is the last page. Therefore if there needs to be a page on the spreadsheet with irrelevant data (e.g user instructions) make sure to put it before the relevant page.

#Todo
1. Change the "extras" spreadsheet to only show relevant columns and to show the correct amount of characters based on current year's limit.