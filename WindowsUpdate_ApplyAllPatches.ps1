#Search Windows Update for all applicable updates
# Call install and force reboot
# Jon Carracino 2020

#Start debug log:
Start-Transcript "C:\windows\installer\WindowsUpdate_ApplyPatches.log" -force -append

$Session = New-Object -ComObject Microsoft.Update.Session           
$Searcher = $Session.CreateUpdateSearcher() 
<#
#Server Selection should use Windows Update unless you only want driver updates*
$Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
$Searcher.SearchScope =  1 # MachineOnly
$Searcher.ServerSelection = 3 # Third Party (Microsoft Update)  
#>

$Searcher.ServerSelection = 2 # WindowsUpdate

$Criteria = "IsInstalled=0 and ISHidden=0"
    Write-Information 'Searching Updates...'
$SearchResult = $Searcher.Search($Criteria)          
$Updates = $SearchResult.Updates

#Show available Patches:
    #$Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | fl
   $CountApplicable = $Updates.Count
   Write-Information "$CountApplicable Patches Found."

#Check if none found:
If ($CountApplicable -eq 0)
{
Write-Information "No Patches Found"
Exit 0
}

#Download the Drivers from Microsoft
$UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
$updates | % { $UpdatesToDownload.Add($_) | out-null }
    Write-Information 'Downloading Drivers...'
$UpdateSession = New-Object -Com Microsoft.Update.Session
$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$Downloader.Download()

#Check if the Drivers are all downloaded and trigger the Installation

$UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
$updates | % { if($_.IsDownloaded) { $UpdatesToInstall.Add($_) | out-null } }

    Write-Information 'Installing Drivers...'
$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToInstall
$InstallationResult = $Installer.Install()

if($InstallationResult.RebootRequired) 
{  
Write-Information('Reboot required! Reboot in 60 min') -Fore Red 
Shutdown /r /t 3600
Exit 3010 
} 
else 
{ 
Exit 0
 }
