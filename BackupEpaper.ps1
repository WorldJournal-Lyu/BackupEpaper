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
Write-Line -Length 50 -Path $log

###################################################################################





# Set up variables

$epaper = (Get-WJPath -Name epaper).Path
$eppub  = (Get-WJHTTP -Name epaper_production).Path + "pub/"
$volumeName = "FantomHD"
$externalHD = (Get-WmiObject win32_logicaldisk | Where-Object{$_.VolumeName -eq $volumeName})
$exepaper   = $externalHD.DeviceID + "\epaper\"
$workDate   = (Get-Date).AddDays(0)
$wc         = New-Object System.Net.WebClient
$pubcodes   = @("AT", "BO", "CH", "DC", "NJ", "NY")

Write-Log -Verb "eppub" -Noun $eppub -Path $log -Type Short -Status Normal
Write-Log -Verb "epaper" -Noun $epaper -Path $log -Type Short -Status Normal
Write-Log -Verb "exepaper" -Noun $exepaper -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

if($externalHD.VolumeName -eq $volumeName){

    Write-Log -Verb "HDD CHECK" -Noun $volumeName -Path $log -Type Long -Status Good
    Write-Line -Length 50 -Path $log

    foreach($pubcode in $pubcodes){

        $jsonName = $pubcode + "-" + $workDate.tostring("yyyy-MM-dd") + ".json" 
        $downloadFrom = $eppub + $pubcode.ToLower() + "/" + $jsonName
        $downloadTo   = $exepaper + $jsonName
        Write-Log -Verb "pubcode" -Noun $pubcode -Path $log -Type Short -Status Normal
        Write-Log -Verb "jsonName" -Noun $jsonName -Path $log -Type Short -Status Normal
        Write-Log -Verb "downloadFrom" -Noun $downloadFrom -Path $log -Type Short -Status Normal
        Write-Log -Verb "downloadTo" -Noun $downloadTo -Path $log -Type Short -Status Normal



        # Download json 

        try{

            Write-Log -Verb "DOWNLOAD FROM" -Noun $downloadFrom -Path $log -Type Long -Status Good
            $wc.DownloadFile($downloadFrom, $downloadTo)
            Write-Log -Verb "DOWNLOAD TO" -Noun $downloadTo -Path $log -Type Long -Status Good

            try{

                $json = Get-Content $downloadTo | ConvertFrom-Json
                Write-Log -Verb "JSON CHECK" -Noun $downloadTo -Path $log -Type Long -Status Good
                Write-Log -Verb "pubdatetime" -Noun $json.pubdatetime -Path $log -Type Short -Status Normal

            }catch{

                $mailMsg = $mailMsg + (Write-Log -Verb "JSON CHECK" -Noun $downloadTo -Path $log -Type Long -Status Bad -Output String) + "`n"
                $hasError = $true

            }

        }catch{

            $mailMsg = $mailMsg + (Write-Log -Verb "DOWNLOAD" -Noun $downloadTo -Path $log -Type Long -Status Bad -Output String) + "`n"
            $hasError = $true

        }

        Write-Line -Length 50 -Path $log



        # Backup Jpg

        Get-ChildItem ($epaper + $workDate.ToString("yyyyMMdd") + "\upload") -Filter ($pubcode+$workDate.ToString("yyyyMMdd")+"*.jpg") | ForEach-Object{

            $copyFrom = $_.FullName
            $copyTo   = $exepaper + $_.Name
            Write-Log -Verb "copyFrom" -Noun $copyFrom -Path $log -Type Short -Status Normal
            Write-Log -Verb "copyTo" -Noun $copyTo -Path $log -Type Short -Status Normal

            try{

                Write-Log -Verb "COPY FROM" -Noun $copyFrom -Path $log -Type Long -Status Good
                Copy-Item $copyFrom $copyTo
                Write-Log -Verb "COPY TO" -Noun $copyTo -Path $log -Type Long -Status Good

            }catch{

                $mailMsg = $mailMsg + (Write-Log -Verb "COPY TO" -Noun $copyTo -Path $log -Type Long -Status Bad) + "`n"
                $hasError = $true
           
            }

        }

        Write-Line -Length 50 -Path $log

    }



    # Check available size

    $externalHD = (Get-WmiObject win32_logicaldisk | Where-Object{$_.VolumeName -eq $volumeName})
    $mailMsg = $mailMsg + (Write-Log -Verb "DISK SPACE" -Noun ("{0:N2}" -f ($externalHD.FreeSpace / 1GB) + " GB Available on " + $externalHD.VolumeName + " (" + $externalHD.DeviceID + ")") -Path $log -Type Long -Status Normal -Output String) + "`n"

}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "HDD CHECK" -Noun $volumeName -Path $log -Type Long -Status Bad -Output String) + "`n"
    $hasError = $true

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

Write-Line -Length 50 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
Emailv2 @emailParam