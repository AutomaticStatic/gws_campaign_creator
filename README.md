# gws_campaign_creator
campaign creator Google Apps Script for campaign calendars

This script implements a custom menu in a Google Sheet for creating Klaviyo campaigns from the campaign calendar.

The script sends the contents of a worksheet to a Lambda function URL, which in turn uses the Campaigns and Tags
endpoints to create and tag the campaigns. 

The lambda is [CampaignCreator](https://us-east-2.console.aws.amazon.com/lambda/home?region=us-east-2#/functions/CampaignCreator?tab=image) 

Details needed to properly configure the script
1. Lambda region, in our case almost always "us-east-2" (line 370)
2. Access Key for the [IAM user apps-script-campaign-creator](https://us-east-1.console.aws.amazon.com/iam/home?region=us-east-2#/users/details/apps-script-campaign-creator?section=permissions) authorized to invoke the lambda (line 371)
3. Secret Key for the [IAM user apps-script-campaign-creator](https://us-east-1.console.aws.amazon.com/iam/home?region=us-east-2#/users/details/apps-script-campaign-creator?section=permissions) authorized to invoke the lambda (line 372)
4. The Function URL for the lambda (line 375)
5. The client_id, e.g., "cor" (line 379)

To implement this script:
1. Open a campaign calendar in Google Sheets
2. Choose "Apps Script" from the "Extensions" menu
3. Paste this script into a file (the file itself can have any name)
4. Give the Apps Script project a name (top-left-ish). Convention is "ClientIDCampaignCreator"
5. Save the project to Google Drive using the Save icon (IDK where it's saved TBH but probably to the GWS user's personal drive)
6. Click on the "Deploy" button and select "New Deployment"
7. Click the gear icon next to Select type and choose Library
8. Enter the project name in the description (at least that's what I've been doing)
9. I'm pretty sure the menu will auto-reload, but you may want to refresh the Google Sheet to be sure