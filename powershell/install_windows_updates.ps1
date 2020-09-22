using namespace System.Runtime.InteropServices

param(
    [int]$RetriesOnReboot=3
)

# A script for installing Windows updates.
# For use cases where installing PSWindowsUpdate module is not ideal/possible.
# Tested only on Windows Server 2016. So, user beware.

Set-Variable -Name RunKey      -Scope Script -Value HKLM:\Software\Microsoft\Windows\CurrentVersion\Run
Set-Variable -Name RunEntry    -Scope Script -value "InstallWindowsUpdates"
Set-Variable -Name ScriptPath  -Scope Script -Value $MyInvocation.MyCommand.Path
Set-Variable -Name MaxAttempts -Scope Script -Value 30

enum OperationResultCode
{
    NotStarted
    InProgress
    Succeeded
    SucceededWithErrors
    Failed
    Aborted
}

$ErrorActionPreference = "Stop"

function Find-Updates
{
    param(
        [__ComObject]$UpdateSession
    )

    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    Write-Host "Searching for updates..."
    $SearchResults  = Start-Search -UpdateSearcher $UpdateSearcher

    if ($SearchResults.ResultCode -eq [OperationResultCode]::Failed)
    {
        Write-Error "Failed to search for updates"

        return @()
    }

    $Updates =  Select-NoUserInput -Updates $SearchResults.Updates.Copy()

    return $Updates
}

function Start-Search
{
    param(
        [__ComObject]$UpdateSearcher
    )

    $Attempts = 0

    while($Attempts -lt $MaxAttempts)
    {
        try
        {
            $SearchResults = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

            return $SearchResults
        }
        catch [COMException]
        {
            $Attempts++
            Write-Error $_.Exception -Silent
            Write-Host "Error searching for updates. Retrying in 30 seconds."
            Start-Sleep -Seconds 30
        }
    }

    Write-Error "Failed to search for updates"
    $SearchResults = @{
        ResultCode = [OperationResultCode]::Failed
    }

    return $SearchResults
}

function Select-NoUserInput
{
    param(
        [Object]$Updates
    )

    $Updates = @($Updates | Where-Object { -not $_.InstallationBehavior.CanRequestUserInput })

    if ($Updates.Count -gt 0)
    {
        Write-Host "The following updates are available:"

        foreach ($Update in $Updates)
        {
            Write-Host " > $($Update.Title)"
        }
    }

    return $Updates
}

function Get-Updates
{
    param(
        [__ComObject]$UpdateSession,
        [Object]$Updates
    )

    $UpdateDownloader = $UpdateSession.CreateUpdateDownloader()
    $UpdatesToDownload = New-Object -ComObject "Microsoft.Update.UpdateColl"

    $Updates | Where-Object { -not $_.IsDownloaded } |
               ForEach-Object { $UpdatesToDownload.Add($_ ) | Out-Null }

    $UpdateDownloader.Updates = $UpdatesToDownload
    $ResultCode = Start-Download -UpdateDownloader $UpdateDownloader
    Write-Host "Download result code: $ResultCode"
    $DownloadedUpdates = Select-Downloaded -Updates $Updates

    return $DownloadedUpdates
}

function Start-Download
{
    param(
        [__ComObject]$UpdateDownloader
    )

    $Attempts = 0

    if ($UpdateDownloader.Updates.Count -ne 0)
    {
        Write-Host "Downloading updates..."

        while ($Attempts -lt $MaxAttempts)
        {
            try
            {
                $DownloadResult = $UpdateDownloader.Download()

                return $DownloadResult.ResultCode
            }
            catch [COMException]
            {
                $Attempts++
                Write-Host $_.Exception
                Write-Host "Error downloading updates. Retrying in 30 seconds."
                Start-Sleep -Seconds 30
            }
        }

        Write-Error "Failed to download updates."

        return [OperationResultCode]::Failed
    }

    Write-Host "No updates require downloading."

    return [OperationResultCode]::Succeeded
}

function Select-Downloaded
{
    param(
        [Object]$Updates
    )

    $DownloadedUpdates = @($Updates | Where-Object { $_.IsDownloaded })

    if ($DownloadedUpdates.Count -gt 0)
    {
        Write-Host "The following updates are downloaded and ready to install:"

        foreach ($DownloadedUpdate in $DownloadedUpdates)
        {
            Write-Host " > $($DownloadedUpdate.Title)"
        }
    }

    return $DownloadedUpdates
}

function Install-Updates
{
    param(
        [__ComObject]$UpdateSession,
        [Object]$Updates
    )

    Write-Host "Installing updates..."
    $UpdateInstaller = $UpdateSession.CreateUpdateInstaller()
    $UpdatesToInstall = New-Object -ComObject "Microsoft.Update.UpdateColl"
    $Updates | ForEach-Object { $UpdatesToInstall.Add($_) | Out-Null }
    $UpdateInstaller.Updates = $UpdatesToInstall
    $InstallationResult = Start-Install -UpdateInstaller $UpdateInstaller

    return $InstallationResult
}

function Start-Install
{
    param(
        [__ComObject]$UpdateInstaller
    )

    try
    {
        $InstallationResult = $UpdateInstaller.Install()
    }
    catch [COMException]
    {
        Write-Error $_.Exception
        Write-Error "Error installing updates."
        $InstallationResult = @{
            ResultCode = [OperationResultCode]::Failed
            RebootRequired = $true
        }
    }

    return $InstallationResult
}

function Add-RestartKey
{
    param(
        [string]$Script,
        [int]$RetriesOnReboot
    )

    $RunValue = "powershell.exe -ExecutionPolicy bypass -File $Script -RetriesOnReboot $RetriesOnReboot" `

    Write-Host "Setting restart registry key."

    Set-ItemProperty -Path $RunKey `
                     -Name $RunEntry `
                     -Value $RunValue `
                     -Force
}

function Remove-RestartKey
{
    $RunScriptEntry = Get-ItemProperty -Path $RunKey -Name $RunEntry -ErrorAction SilentlyContinue

    if ($RunScriptEntry -ne $null)
    {
        Remove-ItemProperty -Path $RunKey -Name $RunEntry -Force
    }
}

function Test-RebootStatus
{
    param(
        $InstallationResult,
        $RetriesOnReboot
    )

    $RebootRequired = $InstallationResult.RebootRequired
    $FailedResult   = $InstallationResult.ResultCode -eq [OperationResultCode]::Failed
    $RetryOnReboot  = $FailedResult -and $RetriesOnReboot -gt 0

    return ($RebootRequired -or $RetryOnReboot)
}

function Main
{
    $UpdateSession = New-Object -ComObject "Microsoft.Update.Session"

    do
    {
        $Updates = Find-Updates -UpdateSession $UpdateSession

        if ($Updates.Count -ne 0)
        {
            $Updates = Get-Updates -UpdateSession $UpdateSession -Updates $Updates
            $InstallationResult = Install-Updates -UpdateSession $UpdateSession `
                                                  -Updates $Updates
        }
        else
        {
            Write-Host "No updates found"
        }

        if (Test-RebootStatus -InstallationResult $InstallationResult -RetriesOnReboot $RetriesOnReboot)
        {
            if ($InstallationResult.ResultCode -eq [OperationResultCode]::Failed)
            {
                $RetriesOnReboot--
                Write-Error "Installation failed."
                Write-Error "Rebooting with $RetriesOnReboot retries remaining."
            }

            Add-RestartKey -Script $ScriptPath -RetriesOnReboot $RetriesOnReboot
            Write-Host "Restarting..."
            Start-Sleep -Seconds 5
            Restart-Computer -Force
        }

    } until ($Updates.Count -eq 0)

    Remove-RestartKey
    Write-Host "Updates complete"
}

#################################################


Main