<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.

Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>
Param(
 $siteUrl, 
 $srcListTitle, 
 $destListTitle,
 $username,
 $password
)

If ($siteUrl -eq $null)
{
   Write-Host "Example)"
   Write-Host ">.\MigrateDiscussionBoard.ps1 -siteUrl https://tenant.sharepoint.com/sites/site -srcListTitle discussion1 -destListTitle discussion2"
   return
}

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")


function ExecuteQueryWithIncrementalRetry($retryCount, $delay)
{
  $retryAttempts = 0;
  $backoffInterval = $delay;
  if ($retryCount -le 0)
  {
    throw "Provide a retry count greater than zero."
  }
  if ($delay -le 0)
  {
    throw "Provide a delay greater than zero."
  }
  while ($retryAttempts -lt $retryCount)
  {
    try
    {
      $script:context.ExecuteQuery()
      return;
    }
    catch [System.Net.WebException]
    {
      $response = $_.Exception.Response
      if ($response -ne $null -and $response.StatusCode -eq 429)
      {
        Write-Host ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($backoffInterval/1000))
        #Add delay.
        Start-Sleep -m $backoffInterval
        #Add to retry count and increase delay.
        $retryAttempts++;
        $backoffInterval = $backoffInterval * 2;
      }
      else
      {
        throw;
      }
    }
  }
  throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}

$script:context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$pwd = convertto-securestring $password -AsPlainText -Force
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $pwd)
$script:context.Credentials = $credentials

$srclist = $script:context.Web.Lists.GetByTitle($srcListTitle)
$destlist = $script:context.Web.Lists.GetByTitle($destListTitle)
$script:context.Load($srclist)
$script:context.Load($destlist)
ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000

$camlquery = New-Object Microsoft.SharePoint.Client.CamlQuery
$items = $srclist.GetItems($camlquery)
 
$script:context.Load($items)
ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000

foreach ($item in $items)
{
  $script:context.Load($item)
  ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000

  $discussionitem = [Microsoft.SharePoint.Client.Utilities.Utility]::CreateNewDiscussion($script:context, $destlist, $item.Item("Title"))
  $discussionitem.Update()
  ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000

  $messagecamlquery = New-Object Microsoft.SharePoint.Client.CamlQuery
  $messagecamlquery.ViewXml = "<View Scope='Recursive'><Query><Where><Eq><FieldRef Name='ParentFolderId'/><Value Type='Integer'>" + $item.Id + "</Value></Eq></Where></Query></View>"
  $messageitems = $srclist.GetItems($messagecamlquery)
  $script:context.Load($messageitems)

  ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000

  $lastModified = $item.Item("Modified")

  foreach ($messageitem in $messageitems)
  {
    $replyitem = [Microsoft.SharePoint.Client.Utilities.Utility]::CreateNewDiscussionReply($script:context, $discussionitem)
    $replyitem.Item("Body") = $messageitem.Item("Body")
    $replyitem.Item("Author") = $messageitem.Item("Author")
    $replyitem.Item("Editor") = $messageitem.Item("Editor")
    $replyitem.Item("Created") = $messageitem.Item("Created")
    $replyitem.Item("Modified") = $messageitem.Item("Modified")
    $replyitem.Update() 

    if ($lastModified -lt $messageitem.Item("Modified"))
    {
      $lastModified = $messageitem.Item("Modified")
    }
  }

  $discussionitem.Item("Body") = $item.Item("Body")
  $discussionitem.Item("Author") = $item.Item("Author")
  $discussionitem.Item("Editor") = $item.Item("Editor")
  $discussionitem.Item("Created") = $item.Item("Created")
  $discussionitem.Item("Modified") = $item.Item("Modified")
  $discussionitem.Item("DiscussionLastUpdated") = $lastModified
  $discussionitem.Item("LastReplyBy") = $item.Item("LastReplyBy")
  $discussionitem.Item("BestAnswerId") = $item.Item("BestAnswerId")
  $discussionitem.Item("IsAnswered") = $item.Item("IsAnswered")
  $discussionitem.Update()

  ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000

}
