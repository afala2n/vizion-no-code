# Setting up a no-code workbook
- To get started, you'll need a Google account and Vizion API Key

1. Copy [this](https://docs.google.com/spreadsheets/d/16HrYIt7UP5qqtHjIJgFjiTmWBu8ACv_Zg_s2sqqDkuM/edit?gid=0#gid=0) file to your Google Drive
2. Navigate to the workbook, and in the menu select Extensions > Apps Script
3. Clear the contents of the `Code.gs` file, and replace it with the contents of `index.js` in this repo.
4. Change the value of `MY_KEY` to your Vizion API key (keeping it in single quotes)
5. Save the Apps Script file
6. Deploy the web app
   - Click the blue "Deploy" button in the top right corner of Apps Script, and select "New Deployment"
   - Select the deployment type "Web App"
   - Write a description, and change the dropdown "Who has access" to "Anyone"
   - Select "Deploy"
   - Copy the Web App URL
7. Navigate to `Code.gs` and paste the Web App URL copied from Step 6 between the single quotes of the constant `MY_URL`
8. Save the Apps Script file
9. Create time-based triggers
   - In the left pane of Apps Script, select the Alarm Clock icon (Triggers)
   - Add a trigger for the following functions, by clicking the blue "+ Add Trigger" icon in the bottom right corner
     - Create a time-based trigger to run `writePayloads_from_cache` every 5 minutes
     - Create a time-based trigger to run `statuses` every 12 hours
     - Create a time-based trigger to run `archive_by_updated` every 3 hours
10. Select the "<>" icon (Editor), and click the "Run" button
11. Review and authorize all permissions
12. Save and close the Apps Script file
13. Refresh your Sheets document, and use the "Vizion Tools" menu to begin tracking!
