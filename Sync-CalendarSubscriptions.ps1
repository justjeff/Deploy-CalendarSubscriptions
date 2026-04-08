<#
  .SYNOPSIS
  Utilizes GAM7 to automate Google Calendar Subscriptions for Google Group members.
  .DESCRIPTION
  Reads group/calendar pairs from a config.json file and uses GAM7 to subscribe all user members (including nested groups) to the calendars.
  .PARAMETER Config
  Launches an interactive mode to view/add Groups and Calendars to config.json
  .PARAMETER ConfigPath
  Path to the config.json file. Defaults to config.json in the script directory.
  .PARAMETER EventLogSource
  Windows Event Log source name. Defaults to Add-CalendarForGroupMembers

  .NOTES
  Requires a config.json file and requires GAM7.

#>
[CmdletBinding()]
param (
    [Parameter()]
    [switch]$Config,
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$EventLogSource = "Sync-CalendarSubscriptions"
)

# --- Helper function for Windows Event Logging ---
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Information", "Warning", "Error")]
        [string]$EntryType = "Information"
    )

    if (-not [System.Diagnostics.EventLog]::SourceExists($EventLogSource)) {
        try {
            New-EventLog -LogName Application -Source $EventLogSource -ErrorAction Stop
        } catch {
            Write-Output "[$EntryType] $Message (Source '$EventLogSource' missing)"
            return
        }
    }
    Write-EventLog -LogName Application -Source $EventLogSource -EntryType $EntryType -EventId 1001 -Message $Message
}

# --- Config Manager ---
function Show-ConfigMenu {
  if (Test-Path $ConfigPath) {
    $data = Get-Content $ConfigPath | ConvertFrom-Json
  } else {
    $data = [PSCustomObject]@{ Groups = @(); Calendars = @() }
  }

  # Interactive Menu Display
  while ($true) {
    Clear-Host
    Write-Host "=== Sync Calendar Subscriptions Config ===" -ForegroundColor Cyan
    Write-Host "`nGroups: "
    $data.Groups | ForEach-Object { Write-Host " - $($_.Email) ( $($_.Label) )" }

    Write-Host "`nCalendars: "
    $data.Calendars | ForEach-Object { Write-Host " - $($_.Id) ( $($_.Label) )" }

    Write-Host "`n[1] Add Group [2] Add Calendar [3] Save & Exit [Q] Quit"
    $choice = Read-Host "`nSelection"

    switch ($choice) {
      "1" {
        $email = Read-Host "Enter Group Email"
        $label = Read-Host "Enter Label (e.g. Marketing)"
        $data.Groups += [PSCustomObject]@{ Email = $email; Label = $label }
      }
      "2" {
        $id = Read-Host "Enter Calendar ID"
        $label = Read-Host "Enter Label (e.g. Events)"
        $data.Calendars += [PSCustomObject]@{ Id = $id; Label = $label }
      }
      "3" {
        $data | ConvertTo-Json -Depth 10 | Out-File $ConfigPath
        Write-Host "Saved Config File: $ConfigPath" -ForegroundColor Green
        return
      }
      "Q" {return}
    }
  }
}

# --- Preflight Checks ---
function Start-Preflight {
  # Verify config is present
  if (-not (Test-Path $ConfigPath)) {
    throw "No config found at $ConfigPath. Run with -Config to set up default."
  }
  # Verify GAM accessibility
  if (-not (Get-Command "gam" -ErrorAction SilentlyContinue)) {
    throw "GAM command not found in PATH. Ensure GAM is installed and accessible for the service account."
  }
}

# --- Entry Point ---
# Run w/ -Config param
if ($Config) {
  Write-Host "This tool is for quick adds. `nTo remove items or make major changes, edit the config file directly in any text editor."
  Write-Host "$ConfigPath"
  Start-Sleep -Seconds 5
  Show-ConfigMenu
  exit
}

$tempCsv = [System.IO.Path]::GetTempFileName()

try {
  Start-Preflight
  $Settings = Get-Content $ConfigPath | ConvertFrom-Json

  foreach ($Group in $Settings.Groups) {
    Write-Log "Starting GAM sync: Group $($Group.Label)"
    # GAM: 'redirect' writes directly to csv, ensuring there's no PS pipe formatting/artifacts to contend with
    # 'print group-members group ...' and 'recursive types user' ensures we grab all users that are members of this group and child groups.
    gam redirect csv "$tempCsv" print group-members group "$($Group.Email)" recursive types user

    # VALIDATE CSV CONTENT: Ensure we have more than just a header row
    $csvData = Import-Csv $tempCsv -ErrorAction SilentlyContinue
    if (-not $csvData) {
        throw "No members found for group $($Group.Label) or group does not exist."
    }
    Write-Log "Found $($csvData.Count) users to process."

    foreach ($Calendar in $Settings.Calendars) {
      Write-Log "Processing: Adding $($Calendar.Label) to members of $($Group.Label)."
      gam csv $tempCsv gam user "~email" add calendar "$($calendar.Id)" selected true
    }
    Write-Log "Sync complete for $($Group.Label)."
  }
}
catch {
    Write-Log "CRITICAL ERROR: $($_.Exception.Message)" -EntryType Error
    throw $_
}
finally {
    if (Test-Path $tempCsv) {
        Remove-Item $tempCsv -ErrorAction SilentlyContinue
    }
}
