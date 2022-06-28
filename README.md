# Exchange
Delete Duplicate Emails
Delete Duplicate emails

Created a Global Admin without MFA. 
Installed EWS using https://www.microsoft.com/en-us/download/details.aspx?id=35371 
	• Set-ExecutionPolicy unrestricted
Downloaded our Template from https://gallery.technet.microsoft.com/office/Removing-Duplicate-Items-f706e1cc#content 
	• $Credentials= Get-Credential
	• .\Remove-DuplicateItems.ps1 -Mailbox chuka@chuhub.ga -Type All -Impersonation -DeleteMode SoftDelete -Mode Quick -Verbose -server outlook.office365.com -Credentials $Credentials (For All Folders).
	• .\Remove-DuplicateItems.ps1 -Mailbox chuka@chuhub.ga -IncludeFolders '#Inbox#\*' -Impersonation -DeleteMode MoveToDeletedItems -Mode Quick -Verbose -server outlook.office365.com -Credentials $Credentials (For Just the Inbox folder).
	• .\Remove-DuplicateItems.ps1 -Mailbox ks@s-point.ru -Credentials $Credentials -Type Calendar -DeleteMode MovetoDeletedItems -Mode Full -NoSize -Verbose (For Duplicate Calendar Items).


https://gallery.technet.microsoft.com/office/Removing-Duplicate-Items-f706e1cc/view/Discussions

![image](https://user-images.githubusercontent.com/16473874/176195078-e2d55a96-b890-403b-9474-bf00dec6a61d.png)
