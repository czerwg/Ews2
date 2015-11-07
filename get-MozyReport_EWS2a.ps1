<#
.Synopsis
   Check Mozy backup by creating a report based on file sent by Mozy to a mailbox. 
.DESCRIPTION
   Check Mozy backup by creating a report based on file sent by Mozy to a mailbox. 
.EXAMPLE
   .\scriptname -User user@somedomain.com -PWord 12345 
   This will connect to EWS using the specified User and Password and will send report to users specified in the default variable recpients 
.EXAMPLE
   .\scriptname -User user@somedomain.com -PWord 12345 -Recipients recipient@domain.com,anotherone@domain.com 
   This will connect to EWS using the specified User and Password and and will send to listed recipients.
.EXAMPLE
   .\scriptname.ps1  
  without parameters will ask for credentials and will send report to users specified in the default variable recpients  
.EXAMPLE
   .\scriptname.ps1  -Recipients recipient@domain.com,anotherone@domain.com 
  without parameters will ask for credentials and send to listed recipients (list recipient as coma separated)
#>
param(
    [string]$User,
    [string]$PWord,
    [string[]]$Recipients #where to send report
     )
#
# Common variables
#
$downloadDirectory = "c:\myfilepath" #where to store attachement
$foldername = "Backup Results" # where to look for attachements - we are only looking one level below inbox
$MailboxToImpersonate = "alerts@ers.ie" ##Define the SMTP Address of the mailbox to impersonate
$AccountWithImpersonationRights = $User ## Define UPN of the Account that has impersonation rights
$Subject = "MozyStatus" ## partial name of subject that we are looking for
[int]$Daystolookback = 3 # Threshold to warn about failed backups.
#
#
#check for download directory
if (!(Test-Path $downloadDirectory)){
    Write-Warning "Please create folder $downloadDirectory"
    break
}
#
#
#

# Credentials were supplied as paramaters or not ?
if (($User -eq "") -or ($PWord -eq "") ) {Write-Warning "No parameters specified or one is missing enter credentials now." 
try{
    $pscred=Get-Credential -ErrorAction SilentlyContinue
    $cred = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
    $AccountWithImpersonationRights = $psCred.UserName.ToString()
	Write-Host "No recipients for report provided please supply one" -Backgroundcolor red
	$Recipients = Read-host "Enter recipient you can supply more than one comma separated"
   } 
  catch {
            Write-Warning "No credentials provided"
            Write-Warning "Press any key to exit"
            $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
         break 
        }
} 
else 
{
write-host "Parameters provided - work in progress " 
# convert plain text password to secure string
if ($PWord) { $SecurePassword = $PWord | ConvertTo-SecureString -AsPlainText -Force }
if ($User) {
            $cred = New-Object System.Net.NetworkCredential($user,$SecurePassword)
            $pscred = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $SecurePassword
            }
            
if (!($Recipients)) { Write-Host "No recipients for report provided please supply one" -Backgroundcolor red
$Recipients = Read-host "Enter recipient you can supply more than one comma separated" 
}
}
## Load Exchange web services DLL
## Download here if not present: http://go.microsoft.com/fwlink/?LinkId=255472
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
# check for EWS module
$ewsinstalled = Test-Path $dllpath
if (!($ewsinstalled)) {
                        Write-Warning "Please install EWS http://go.microsoft.com/fwlink/?LinkId=255472"
                        Write-Warning "Press any key to exit"
                        $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                        break 
                      }
Import-Module $dllpath
## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013 
## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion) 
#connect to exchange service
$service.Credentials = $cred 
## Set the URL of the CAS (Client Access Server)
try{
    $service.AutodiscoverUrl($AccountWithImpersonationRights ,{$true}) 
   } 
  catch {
            Write-Warning "Wrong username or password" 
            Write-Warning "Press any key to exit"
            $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            break
        }

##Login to Mailbox with Impersonation
Write-Host 'Using ' $AccountWithImpersonationRights ' to Impersonate ' $MailboxToImpersonate
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate ); 
#Connect to the Inbox and find the ID of Subfolder
$InboxFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$ImpersonatedMailboxName) 
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$InboxFolder)
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$foldername)
#Now we have folder - we are going to use its ID later
$findFolderResults = $Inbox.FindFolders($SfSearchFilter,$fvFolderView)  
#Cleanup the directory where files are stored localy.
if ((Get-ChildItem $downloadDirectory ).Exists) { Remove-Item ($downloadDirectory + "\*.csv")}
# Narrow down the search only to unread items.
#$newitems = $findFolderResults.UnreadCount
#if ($newitems -gt 0)  {
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(50) 
# Define what we are looking for
# Check if e-mail was already read
$Sfir = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)
# Look for part of the subject
$Sfsub = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $Subject)
# And return only emails that have attachement
$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
# Search for items that are received for last 24hours from now.
$sfold = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,([System.Datetime]::now).AddDays(-1))
# Build the search filter
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
$sfCollection.add($Sfir)
$sfCollection.add($Sfsub)
$sfCollection.add($Sfha)
#$sfCollection.add($Sfold)
# Find the items that matches the criteria above
$fiItems = $service.FindItems($findFolderResults.Id,$sfCollection,$ivItemView)  
#Save the attachements 
foreach($miMailItems in $fiItems.Items){
	$miMailItems.Load()
	foreach($attach in $miMailItems.Attachments){
		# Only extract csv attachments. If you need additional filetypes, include them as an OR in the second if below. To extract all attachments, remove these two if loops
		If($attach -is[Microsoft.Exchange.WebServices.Data.FileAttachment]){
			if($attach.Name -like "*status*.csv"){  
				$attach.Load()
				$prefix = get-date 	
				$fiFile = new-object System.IO.FileStream(($downloadDirectory + "\" +  $attach.Name.ToString()), [System.IO.FileMode]::Create)    		

				$fiFile.Write($attach.Content, 0, $attach.Content.Length)
				$fiFile.Close()
			}
		}
	}
    $miMailItems.IsRead = $true
$miMailItems.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
}
##Process the downloaded files
#Prepare style for report
$style ="
<style>
        BODY{background-color:#b0c4de;}
    TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TH{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:#778899}
    TD{border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
    tr:nth-child(odd) { background-color:#d3d3d3;} 
    tr:nth-child(even) { background-color:white;}  
</style>
"
if ((Get-ChildItem $downloadDirectory -Filter "*.csv").Exists ){
    $file = Get-ChildItem -Path $downloadDirectory -Filter "*.csv" | select -last 1 #get only latest file
    $txt= Get-Content ($file.FullName) # read the file as we need to remove some lines
    $txt |select -Skip 1 |Out-File $file #remove the first line from file and save it again
    $csv = import-csv $file # import modified file
    # build the table
    $tbl = $csv | select  username,"Machine Alias",@{Name="LastBackup";Expression={(Get-Date([datetime]::ParseExact($_.{Last Successful Backup}, "yyyy-MM-dd HH:mm:ss", $null)) )}} |  Sort-Object LastBackup
    # first list the backups older than x days
    $part1 =$tbl | Where-Object { $_.Lastbackup -lt (Get-Date).AddDays(-($Daystolookback))} |ConvertTo-Html -As Table -PreContent "Report generated on $(get-date),backup didnt run for last $Daystolookback " -Head $style |Out-String
    # and now the remainder
    $part2 =$tbl | Where-Object { $_.Lastbackup -ge (Get-Date).AddDays(-($Daystolookback))} |ConvertTo-Html -As Table -PreContent "All other backups" -Head $style |Out-String
    # ready to build html for attach to e-mail and send it.
    $mailatt = ConvertTo-Html -head $Style -PostContent $part1, $part2 -PreContent "<h1>Mozy Backups</h1> <h2> based on file $file </h2>" |Out-String
    Send-MailMessage -to $Recipients -from $AccountWithImpersonationRights -Credential $pscred -SmtpServer smtp.office365.com -Port 587 -UseSsl -subject "Mozy Report for $(Get-Date -Format d)" -BodyAsHtml $mailatt
}
else
{
    Send-MailMessage -to $Recipients -from $AccountWithImpersonationRights -Credential $pscred -SmtpServer smtp.office365.com -Port 587 -UseSsl -subject "Someone already looked at Mozy Backups?" -Body "No reports from Mozy to process, either someone already checked or report not delivered"
}

#
#
#
#

