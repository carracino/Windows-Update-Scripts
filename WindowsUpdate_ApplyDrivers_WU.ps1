# Search Windows Update for all applicable updates
# v1.2 - will only detect driver updates. If a communication error interrupts WU check, it will restart service and retry*
# Jon Carracino 2020


#Start debug log:
Start-Transcript "C:\windows\installer\WindowsUpdate_ApplyDrivers.log" -force -Append

$Session = New-Object -ComObject Microsoft.Update.Session           
$Searcher = $Session.CreateUpdateSearcher() 
<#  Below is for reference only*
#Server Selection should use Windows Update unless you only want driver updates*
$Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
$Searcher.SearchScope =  1 # MachineOnly
$Searcher.ServerSelection = 3 # Third Party (Microsoft Update)  
#>

$Searcher.ServerSelection = 2 # WindowsUpdate is selected channel

Try
{ # First try to check for updates
$Criteria = "IsInstalled=0 and Type='Driver' and ISHidden=0"   #Type is Drivers only that are not installed
    Write-Information 'Searching Updates...'
$SearchResult = $Searcher.Search($Criteria)   
    Write-Information "Update Search complete with code: $SearchResult"       
$Updates = $SearchResult.Updates

#Show available Patches:
    $Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | fl
   $CountApplicable = $Updates.Count
   Write-Information "$CountApplicable Patches Found."

#Check if none found:
If ($CountApplicable -eq 0)
{
   Write-Information "No Patches Found 1st check."
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
Write-Information('Reboot required!') -Fore Red 
#Shutdown /r /t 3600
Stop-Transcript
Exit 3010 
} 
else 
{ 

#Exit 0
 }

 } # End of 1st try for check for updates 
 
 Catch
 {
  # Error was detected, windows update service will be restarted and 2nd check will begin:
  try 
  {
  Write-host "Restrating windows update service..."
  Stop-Service 'wuauserv' -Force
  Start-Sleep -Seconds 5
  Start-Service 'wuauserv' 
  Start-Sleep -Seconds 5
  }
  Catch
  {
  #try again after 30 sec
  Start-Sleep -Seconds 30
    Write-host "Restrating windows update service..."
  Stop-Service 'wuauserv' -Force
  Start-Sleep -Seconds 5
  Start-Service 'wuauserv' 
  Start-Sleep -Seconds 5
  }
  ##

  ## 2nd try to ensure a successful connection:
  $Criteria = "IsInstalled=0 and Type='Driver' and ISHidden=0"   #Type is Drivers only that are not installed
    Write-Information 'Searching Updates...'
$SearchResult = $Searcher.Search($Criteria)          
$Updates = $SearchResult.Updates

#Show available Patches:
    $Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | fl
   $CountApplicable = $Updates.Count
   Write-Information "$CountApplicable Patches Found."

#Check if none found:
If ($CountApplicable -eq 0)
{
Write-Information "No Patches Found"
Stop-Transcript
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
Write-Information('Reboot required!') -Fore Red 
#Shutdown /r /t 3600
Stop-Transcript
Exit 3010 
} 
else 
{
Stop-Transcript 
Exit 0
 }

 }
