# CopyDiscussionBoardItems

This powershell script is to copy all the list items of source Discussion Board List to the destination Discussion Board List.

The primary reason of using this script is when in emergency.
Discussion Board list could be easiliy broken. For example, when you modify content type or views, you will soon lose the stable Discussion Board. And you will never get it back again.

In that case, we generally need to create new Discussion Board List and copy all the items to the new list.

## Prerequitesite
You need to download SharePoint Online Client Components SDK to run this script.
https://www.microsoft.com/en-us/download/details.aspx?id=42038

You can also acquire the latest SharePoint Online Client SDK by Nuget as well.

1. You need to access the following site.
https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM

2. Download the nupkg.
3. Change the file extension to *.zip.
4. Unzip and extract those file.
5. place them in the specified directory from the code. 

### Make sure that you need to create the destination discussion board list in advance.

## How to Run - parameters

-siteUrl ... Target site collection (site) or site (web) URL.

-srcListTitle ... Copy Source Discussion Board List Name (Title)

-destListTitle ... Copy Destination Discussion Board List Name (Title)

-username ... Site Administrator Account to check the workflow instances.

-password ... The password of the above user.

### Example 
.\CopyDiscussionBoardItems.ps1 -siteUrl https://tenant.sharepoint.com/sites/discussionsite -srcListTitle SourceList -destListTitle DestList -username admin@tenant.onmicrosoft.com -password PASSWORD



## Reference
Original Documentation of this sample code from the following Japan SharePoint Support Team's blog.
https://blogs.technet.microsoft.com/sharepoint_support/2015/01/21/1250-2/

## Remarks
In the long run, it is better to move to Yammer or Teams conversation, instead of keep using Discussion Board List. 
Only if you are not in a hurry to fix it.
