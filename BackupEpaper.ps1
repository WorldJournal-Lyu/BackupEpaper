<#

BackupEpaper.ps1

    2018-07-17 Initial Creation

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Import-Module WorldJournal.Ftp -Verbose -Force
Import-Module WorldJournal.Log -Verbose -Force
Import-Module WorldJournal.Email -Verbose -Force
Import-Module WorldJournal.Server -Verbose -Force
Import-Module WorldJournal.Database -Verbose -Force

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$hasError   = $false

$newlog     = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log        = $newlog.FullName
$logPath    = $newlog.Directory

$mailFrom   = (Get-WJEmail -Name noreply).MailAddress
$mailPass   = (Get-WJEmail -Name noreply).Password
$mailTo     = (Get-WJEmail -Name lyu).MailAddress
$mailSbj    = $scriptName
$mailMsg    = ""

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 100 -Path $log

###################################################################################





# Set up variables

$epaper = (Get-WJPath -Name epaper).Path
$eppub  = (Get-WJHTTP -Name epaper_production).Path + "pub/"
$volumeName = "FantomHD"
$externalHD = (Get-WmiObject win32_logicaldisk | Where-Object{$_.VolumeName -eq $volumeName})
$exepaper   = $externalHD.DeviceID + "\epaper\"
$workDate   = (Get-Date).AddDays(0)
$wc         = New-Object System.Net.WebClient
$pubcodes   = @("NJ", "BO", "CH", "DC", "AT", "NY")

Write-Log -Verb "eppub   " -Noun $eppub -Path $log -Type Short -Status Normal
Write-Log -Verb "epaper  " -Noun $epaper -Path $log -Type Short -Status Normal
Write-Log -Verb "exepaper" -Noun $exepaper -Path $log -Type Short -Status Normal
Write-Line -Length 100 -Path $log

if(Test-Path $exepaper){

    foreach($pubcode in $pubcodes){

        $jsonFileName   = $pubcode + "-" + $workDate.tostring("yyyy-MM-dd") + ".json" 
        $remoteFilePath = $eppub + $pubcode.ToLower() + "/" + $jsonFileName
        $localFilePath  = $exepaper + $jsonFileName
        Write-Log -Verb "jsonFileName  " -Noun $jsonFileName -Path $log -Type Short -Status Normal
        Write-Log -Verb "remoteFilePath" -Noun $remoteFilePath -Path $log -Type Short -Status Normal
        Write-Log -Verb "localFilePath " -Noun $localFilePath -Path $log -Type Short -Status Normal

        Write-Log -Verb "DOWNLOAD FROM" -Noun $remoteFilePath -Path $log -Type Long -Status Normal
        Write-Log -Verb "DOWNLOAD TO" -Noun $localFilePath -Path $log -Type Long -Status Normal

        try{
            $wc.DownloadFile($remoteFilePath, $localFilePath)
            Write-Log -Verb "DOWNLOAD" -Noun $remoteFilePath -Path $log -Type Long -Status Good
        }catch{
            Write-Log -Verb "DOWNLOAD" -Noun $remoteFilePath -Path $log -Type Long -Status Bad
        }

    }
       
}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "CANNOT FIND" -Noun $volumeName -Path $log -Type Long -Status Bad -Output String)
    
}




# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}

# Flag hasError 

if( $false ){
    $hasError = $true
}



###################################################################################

Write-Line -Length 100 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $scriptName + " completed at " + (Get-Date).ToString("HH:mm:ss") + "`n`n" + $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
Emailv2 @emailParam