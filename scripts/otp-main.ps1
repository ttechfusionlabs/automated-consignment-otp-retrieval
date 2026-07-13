<#
===============================================================================
 otp-lookup-workflow
 otp-main.ps1

 PURPOSE
    Watches Outlook 2021 Desktop (Office365 Exchange, COM automation - NO
    Graph API, NO Azure App Registration, NO OAuth) for unread emails that
    carry a single .xlsx attachment with a "CN" column. For each qualifying
    mail it:
       1. Saves the attachment to the shared /files/input folder
       2. Calls the n8n webhook and waits for the OTP lookup to complete
       3. Replies to the SAME email thread with the generated OTP Excel
          attached, using ONLY the required reply body text
       4. Moves the processed mail into a "Processed_OTP" folder
    All steps are logged, retried on transient failure, and wrapped in
    per-mail error handling so one bad mail never blocks the rest.

 REQUIREMENTS
    - Windows 11, PowerShell 5.1
    - Outlook 2021 Desktop MUST be installed and (ideally) already running
      and signed in to the Office365 Exchange account
    - Docker Desktop + n8n stack already up (docker-compose.yml)
    - config/otp-config.json present alongside this script (or path passed
      via -ConfigPath)

 SCHEDULING
    Run via Windows Task Scheduler every 1-5 minutes. See
    docs/02-deployment-guide.md for the exact Task Scheduler setup
    (must run in an interactive session / with "Run only when user is
    logged on" because Outlook COM automation requires a desktop session).

 USAGE
    powershell.exe -ExecutionPolicy Bypass -File otp-main.ps1
    powershell.exe -ExecutionPolicy Bypass -File otp-main.ps1 -ConfigPath "C:\n8n\n8n-otp\scripts\otp-config.json"
===============================================================================
#>

[CmdletBinding()]
param(
    # Allows overriding the config file location from Task Scheduler / CLI.
    # NOTE: We do NOT rely on $PSScriptRoot here - it can resolve to an
    # empty string when this script is launched by Task Scheduler / certain
    # non-interactive contexts, which would crash Join-Path before logging
    # even starts. The fixed deployment path is used as a reliable default;
    # override with -ConfigPath if you ever move the deployment folder.
    [string]$ConfigPath = "C:\n8n\n8n-otp\config\otp-config.json"
)

# =============================================================================
# SECTION 0: STRICT MODE & GLOBAL ERROR PREFERENCE
# =============================================================================
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"      # every uncaught error becomes a terminating error we can catch
$ProgressPreference    = "SilentlyContinue"

# =============================================================================
# SECTION 1: LOAD CONFIGURATION
# =============================================================================
if (-not (Test-Path -LiteralPath $ConfigPath)) {
    Write-Error "Configuration file not found at: $ConfigPath"
    exit 1
}

try {
    $Config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
}
catch {
    Write-Error "Failed to parse configuration JSON at '$ConfigPath': $($_.Exception.Message)"
    exit 1
}

# Convenience shortcuts (kept close to config so nothing is hard-coded below)
$Cfg_InboxFolderName        = $Config.Outlook.InboxFolderName
$Cfg_ProcessedFolderName    = $Config.Outlook.ProcessedFolderName
$Cfg_FailedFolderName       = $Config.Outlook.FailedFolderName
$Cfg_SenderFilter           = $Config.Outlook.SenderFilter
$Cfg_SubjectFilter          = $Config.Outlook.SubjectFilter
$Cfg_RequireXlsxAttachment  = [bool]$Config.Outlook.OnlyProcessWithXlsxAttachment
$Cfg_MarkAsRead             = [bool]$Config.Outlook.MarkAsReadAfterProcessing

$Cfg_SecondMailHour         = [int]$Config.TimeWindow.SecondMailHourThreshold

$Cfg_InputFolder            = $Config.Paths.InputFolder
$Cfg_OutputFolder           = $Config.Paths.OutputFolder
$Cfg_LogFolder              = $Config.Paths.LogFolder
$Cfg_InputFileName          = $Config.Paths.InputFileName

$Cfg_DockerInputFolder      = $Config.DockerFilePaths.InputFolder
$Cfg_DockerOutputFolder     = $Config.DockerFilePaths.OutputFolder

$Cfg_WebhookUrl             = $Config.Webhook.Url
$Cfg_WebhookHeaderName      = $Config.Webhook.SharedSecretHeaderName
$Cfg_WebhookHeaderValue     = $Config.Webhook.SharedSecretHeaderValue
$Cfg_WebhookTimeoutSeconds  = [int]$Config.Webhook.TimeoutSeconds
$Cfg_RetryCount             = [int]$Config.Webhook.RetryCount
$Cfg_RetryDelaySeconds      = [int]$Config.Webhook.RetryDelaySeconds

$Cfg_ReplyBodyText          = $Config.Email.ReplyBodyText
$Cfg_SendAlertOnFailure     = [bool]$Config.Email.SendAlertOnFailure
$Cfg_AlertRecipient         = $Config.Email.AlertRecipient
$Cfg_AlertSubjectPrefix     = $Config.Email.AlertSubjectPrefix

$Cfg_LogFileNamePattern     = $Config.Logging.LogFileNamePattern
$Cfg_MaxLogAgeDays          = [int]$Config.Logging.MaxLogAgeDays
$Cfg_VerboseConsole         = [bool]$Config.Logging.VerboseConsoleOutput

# =============================================================================
# SECTION 2: ENSURE REQUIRED FOLDERS EXIST ON DISK
# =============================================================================
foreach ($folder in @($Cfg_InputFolder, $Cfg_OutputFolder, $Cfg_LogFolder)) {
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

# =============================================================================
# SECTION 3: LOGGING
# =============================================================================
# Single log file per calendar day. Every function below calls Write-Log
# instead of Write-Host directly so behaviour (console + file) stays uniform.
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")][string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logFileName = [string]::Format($Cfg_LogFileNamePattern, (Get-Date))
    $logFilePath = Join-Path $Cfg_LogFolder $logFileName
    $line = "[$timestamp] [$Level] $Message"

    # File output (always)
    try {
        Add-Content -LiteralPath $logFilePath -Value $line -Encoding UTF8
    }
    catch {
        # If logging itself fails, fall back to console only - never let
        # logging failures crash the main process.
        Write-Host "LOGGING FAILURE: $($_.Exception.Message)"
    }

    # Console output (optional, controlled by config)
    if ($Cfg_VerboseConsole) {
        switch ($Level) {
            "ERROR"   { Write-Host $line -ForegroundColor Red }
            "WARN"    { Write-Host $line -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $line -ForegroundColor Green }
            default   { Write-Host $line }
        }
    }
}

function Remove-OldLogs {
    # Housekeeping: delete log files older than Cfg_MaxLogAgeDays
    try {
        Get-ChildItem -LiteralPath $Cfg_LogFolder -Filter "otp-log-*.txt" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$Cfg_MaxLogAgeDays) } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Log rotation cleanup failed: $($_.Exception.Message)" -Level WARN
    }
}

# =============================================================================
# SECTION 4: OUTLOOK COM HELPERS
# =============================================================================

function Get-OutlookApplication {
    <#
        Attaches to an already-running Outlook instance if one exists
        (preferred - avoids extra profile-selection prompts), otherwise
        starts a new one. This is standard COM automation - no Graph API,
        no OAuth, no Azure App Registration involved anywhere.
    #>
    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        Write-Log "Attached to already-running Outlook instance."
    }
    catch {
        Write-Log "No running Outlook instance found - starting a new one." -Level WARN
        $outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Seconds 5   # give Outlook time to fully initialize/sign in
    }
    return $outlook
}

function Get-OrCreateSubFolder {
    <#
        Returns a subfolder of Inbox with the given name, creating it if it
        does not already exist. Used for "Processed_OTP" and "Failed_OTP".
    #>
    param(
        [Parameter(Mandatory)]$ParentFolder,
        [Parameter(Mandatory)][string]$FolderName
    )

    foreach ($f in $ParentFolder.Folders) {
        if ($f.Name -eq $FolderName) {
            return $f
        }
    }

    Write-Log "Sub-folder '$FolderName' not found under '$($ParentFolder.Name)' - creating it."
    return $ParentFolder.Folders.Add($FolderName)
}

function Get-SenderSmtpAddress {
    <#
        MailItem.SenderEmailAddress returns a real SMTP address (like
        "name@domain.com") ONLY for external senders. For senders who are
        INTERNAL to the same Exchange organization (very common - e.g. a
        colleague at the same company), Outlook instead returns a long
        technical "X.500" address like:
            /O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP.../CN=RECIPIENTS/CN=...
        This function resolves the REAL SMTP address in both cases, so
        SenderFilter comparisons work correctly regardless of whether the
        sender is internal or external.
    #>
    param($MailItem)

    try {
        if ($MailItem.SenderEmailType -eq 'EX') {
            # Internal Exchange sender - resolve via GetExchangeUser()
            $exchUser = $MailItem.Sender.GetExchangeUser()
            if ($exchUser -and $exchUser.PrimarySmtpAddress) {
                return $exchUser.PrimarySmtpAddress
            }
            # Fallback: read the PR_SMTP_ADDRESS MAPI property directly
            $PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
            return $MailItem.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS)
        }
        else {
            # External sender - SenderEmailAddress is already the real SMTP address
            return $MailItem.SenderEmailAddress
        }
    }
    catch {
        # Last-resort fallback so a resolution failure never crashes the run
        try {
            $PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
            return $MailItem.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS)
        }
        catch {
            return $MailItem.SenderEmailAddress
        }
    }
}

function Test-HasXlsxAttachment {
    param($MailItem)
    foreach ($att in $MailItem.Attachments) {
        if ($att.FileName -match '\.xlsx$') { return $true }
    }
    return $false
}

function Save-XlsxAttachment {
    <#
        Saves the FIRST .xlsx attachment found on the mail item to the
        configured input path, OVERWRITING any previous input.xlsx.
        Returns the full saved file path, or $null if none found.
    #>
    param($MailItem, [string]$DestinationPath)

    foreach ($att in $MailItem.Attachments) {
        if ($att.FileName -match '\.xlsx$') {
            $att.SaveAsFile($DestinationPath)
            Write-Log "Saved attachment '$($att.FileName)' to '$DestinationPath'."
            return $DestinationPath
        }
    }
    return $null
}

# =============================================================================
# SECTION 5: OUTPUT FILENAME LOGIC (1st mail vs 2nd mail of the day)
# =============================================================================
function Get-OutputFileName {
    <#
        Decides OTP_DDMMYY.xlsx vs OTP_DDMMYY_2nd.xlsx based on the mail's
        ReceivedTime (NOT the time the script happens to run - this makes
        the logic correct even if the script is delayed or catches up on
        a backlog).
    #>
    param([datetime]$ReceivedTime)

    $datePart = $ReceivedTime.ToString("ddMMyy")

    if ($ReceivedTime.Hour -ge $Cfg_SecondMailHour) {
        return "OTP_$($datePart)_2nd.xlsx"
    }
    else {
        return "OTP_$($datePart).xlsx"
    }
}

# =============================================================================
# SECTION 6: WEBHOOK CALL WITH RETRY
# =============================================================================
function Invoke-OtpWebhookWithRetry {
    <#
        Calls the n8n webhook, passing the docker-side input path and the
        desired output filename. Retries transient failures (network
        blips, n8n still starting up, etc.) up to Cfg_RetryCount times.
        Returns the parsed JSON response object on success, or throws
        after exhausting retries.
    #>
    param(
        [Parameter(Mandatory)][string]$DockerInputFilePath,
        [Parameter(Mandatory)][string]$OutputFileName
    )

    $payload = @{
        inputFile      = $DockerInputFilePath
        outputFileName = $OutputFileName
    } | ConvertTo-Json -Compress

    $headers = @{ $Cfg_WebhookHeaderName = $Cfg_WebhookHeaderValue }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Write-Log "Calling n8n webhook (attempt $attempt of $Cfg_RetryCount): $Cfg_WebhookUrl"
            $response = Invoke-RestMethod -Uri $Cfg_WebhookUrl `
                                           -Method Post `
                                           -Body $payload `
                                           -ContentType "application/json" `
                                           -Headers $headers `
                                           -TimeoutSec $Cfg_WebhookTimeoutSeconds

            if (-not $response.success) {
                throw "n8n reported failure: $($response.error)"
            }

            Write-Log "Webhook call succeeded on attempt $attempt." -Level SUCCESS
            return $response
        }
        catch {
            Write-Log "Webhook attempt $attempt failed: $($_.Exception.Message)" -Level WARN
            if ($attempt -ge $Cfg_RetryCount) {
                throw "n8n webhook failed after $Cfg_RetryCount attempts: $($_.Exception.Message)"
            }
            Start-Sleep -Seconds $Cfg_RetryDelaySeconds
        }
    }
}

# =============================================================================
# SECTION 7: REPLY + ATTACH + MOVE
# =============================================================================
function Format-OtpExcel {
    <#
        Opens the generated OTP Excel file via Excel COM automation and
        applies visual formatting:
          - Header row (CN, OTP): bold text + yellow fill
          - Full used range: thin borders around every cell (table look)
          - Auto-fit column widths

        n8n's Spreadsheet File node produces a plain, unstyled .xlsx - this
        step polishes it right before the file gets attached to the reply.
        It uses the Excel installation already present on this machine
        (same Office suite as Outlook) - no extra libraries, no n8n
        changes required.

        Failures here are NON-FATAL: if formatting fails for any reason,
        the original (unstyled) file is still sent rather than blocking
        the whole reply.
    #>
    param([string]$ExcelFilePath)

    $excel     = $null
    $workbook  = $null
    $sheet     = $null

    try {
        Write-Log "Formatting output Excel via Excel COM: '$ExcelFilePath'"

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($ExcelFilePath)
        $sheet    = $workbook.Worksheets.Item(1)

        $usedRange = $sheet.UsedRange
        $rowCount  = $usedRange.Rows.Count
        $colCount  = $usedRange.Columns.Count

        # ---- Header row: bold + yellow fill ----------------------------
        $headerRange = $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item(1,$colCount))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.Color = 65535   # yellow (VBA color code)

        # ---- Full table borders (header + all data rows) ---------------
        $fullRange = $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item($rowCount,$colCount))
        $borders = $fullRange.Borders
        # xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10, xlInsideVertical=11, xlInsideHorizontal=12
        foreach ($borderIndex in 7,8,9,10,11,12) {
            $borders.Item($borderIndex).LineStyle = 1   # xlContinuous
            $borders.Item($borderIndex).Weight = 2      # xlThin
        }

        # ---- Auto-fit columns for a clean look --------------------------
        $usedRange.EntireColumn.AutoFit() | Out-Null

        $workbook.Save()
        Write-Log "Excel formatting applied successfully." -Level SUCCESS
    }
    catch {
        Write-Log "Excel formatting failed (non-fatal - file will still be sent unstyled): $($_.Exception.Message)" -Level WARN
    }
    finally {
        # Always clean up COM objects, even if formatting failed midway,
        # so we never leave an orphaned hidden Excel.exe process running.
        try { if ($workbook) { $workbook.Close($true) } } catch {}
        try { if ($excel)    { $excel.Quit() } } catch {}
        if ($sheet)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)    | Out-Null }
        if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
        if ($excel)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Send-OtpReply {
    <#
        Replies to ALL original recipients + CC (ReplyAll) with ONLY the
        required body text and the generated Excel attached.
    #>
    param($MailItem, [string]$AttachmentFullPath)

    if (-not (Test-Path -LiteralPath $AttachmentFullPath)) {
        throw "Expected output file not found at '$AttachmentFullPath' - cannot reply."
    }

    # ReplyAll() (not Reply()) - the real workflow requires the response to
    # reach every original recipient AND everyone CC'd on the request, not
    # just the original sender.
    $reply = $MailItem.ReplyAll()

    # Overwrite the ENTIRE body (Outlook's ReplyAll() normally quotes the
    # original message below the cursor) - requirement states the reply
    # body must contain ONLY this single line, nothing else.
    $reply.BodyFormat = 1     # olFormatPlain -> guarantees no residual HTML quoting
    $reply.Body = $Cfg_ReplyBodyText

    $reply.Attachments.Add($AttachmentFullPath) | Out-Null
    $reply.Send()

    Write-Log "Reply-all sent with attachment '$AttachmentFullPath'." -Level SUCCESS
}

function Move-MailToFolder {
    param($MailItem, $TargetFolder)
    $MailItem.Move($TargetFolder) | Out-Null
}

# =============================================================================
# SECTION 8: FAILURE ALERTING (separate from the main OTP reply)
# =============================================================================
function Send-AlertMail {
    param([string]$Subject, [string]$Body, $OutlookApp)

    if (-not $Cfg_SendAlertOnFailure) { return }

    try {
        $alert = $OutlookApp.CreateItem(0)   # olMailItem
        $alert.To = $Cfg_AlertRecipient
        $alert.Subject = "$Cfg_AlertSubjectPrefix $Subject"
        $alert.Body = $Body
        $alert.Send()
        Write-Log "Alert email sent to $Cfg_AlertRecipient." -Level WARN
    }
    catch {
        Write-Log "Failed to send alert email itself: $($_.Exception.Message)" -Level ERROR
    }
}

# =============================================================================
# SECTION 9: PER-MAIL PROCESSING (fully isolated try/catch per email)
# =============================================================================
function Invoke-ProcessSingleMail {
    param(
        $MailItem,
        $OutlookApp,
        $ProcessedFolder,
        $FailedFolder
    )

    $subjectForLog = $MailItem.Subject
    Write-Log "----- Processing mail: '$subjectForLog' (received $($MailItem.ReceivedTime)) -----"

    try {
        # --- Filters -------------------------------------------------------
        $senderSmtp = Get-SenderSmtpAddress -MailItem $MailItem
        if ($Cfg_SenderFilter -and $senderSmtp -notlike "*$Cfg_SenderFilter*") {
            Write-Log "Skipped (sender '$senderSmtp' does not match filter)."
            return
        }
        if ($Cfg_SubjectFilter -and $MailItem.Subject -notlike "*$Cfg_SubjectFilter*") {
            Write-Log "Skipped (subject does not match filter)."
            return
        }
        if ($Cfg_RequireXlsxAttachment -and -not (Test-HasXlsxAttachment -MailItem $MailItem)) {
            Write-Log "Skipped (no .xlsx attachment found)."
            return
        }

        # --- Step 1: Save attachment locally ---------------------------------
        $localInputPath = Join-Path $Cfg_InputFolder $Cfg_InputFileName
        $saved = Save-XlsxAttachment -MailItem $MailItem -DestinationPath $localInputPath
        if (-not $saved) {
            throw "No .xlsx attachment could be saved."
        }

        # --- Step 2: Decide output filename (1st vs 2nd mail) ---------------
        $outputFileName = Get-OutputFileName -ReceivedTime $MailItem.ReceivedTime
        Write-Log "Determined output filename: $outputFileName"

        # --- Step 3: Call n8n webhook (with retry) ---------------------------
        $dockerInputPath = "$Cfg_DockerInputFolder/$Cfg_InputFileName"
        $response = Invoke-OtpWebhookWithRetry -DockerInputFilePath $dockerInputPath `
                                                -OutputFileName $outputFileName

        # --- Step 4: Resolve the LOCAL (Windows) path of the generated file --
        # n8n writes to /files/output/<name> inside the container, which is
        # bind-mounted from $Cfg_OutputFolder on the Windows host - so the
        # local path is simply the output folder + filename.
        $localOutputPath = Join-Path $Cfg_OutputFolder $outputFileName

        # Give the filesystem a brief moment in case of any write-flush delay
        $waited = 0
        while (-not (Test-Path -LiteralPath $localOutputPath) -and $waited -lt 10) {
            Start-Sleep -Seconds 1
            $waited++
        }
        if (-not (Test-Path -LiteralPath $localOutputPath)) {
            throw "n8n reported success but output file was not found at '$localOutputPath'."
        }

        # --- Step 5: Apply visual formatting (yellow bold header, table borders)
        Format-OtpExcel -ExcelFilePath $localOutputPath

        # --- Step 6: Reply-all to the SAME mail with the generated Excel -----
        Send-OtpReply -MailItem $MailItem -AttachmentFullPath $localOutputPath

        # --- Step 7: Mark read + move to Processed folder ---------------------
        if ($Cfg_MarkAsRead) { $MailItem.UnRead = $false }
        Move-MailToFolder -MailItem $MailItem -TargetFolder $ProcessedFolder

        Write-Log "Mail '$subjectForLog' processed successfully end-to-end." -Level SUCCESS
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log "ERROR processing mail '$subjectForLog': $errMsg" -Level ERROR

        # Move the failed mail out of Inbox so it doesn't get retried forever
        # and doesn't block processing of subsequent mails.
        try {
            if ($Cfg_MarkAsRead) { $MailItem.UnRead = $false }
            Move-MailToFolder -MailItem $MailItem -TargetFolder $FailedFolder
            Write-Log "Mail moved to '$Cfg_FailedFolderName' for manual review." -Level WARN
        }
        catch {
            Write-Log "Additionally failed to move mail to Failed folder: $($_.Exception.Message)" -Level ERROR
        }

        Send-AlertMail -Subject "Failed to process mail: $subjectForLog" `
                        -Body "Error: $errMsg`r`nReceived: $($MailItem.ReceivedTime)`r`nSee log folder: $Cfg_LogFolder" `
                        -OutlookApp $OutlookApp
    }
}

# =============================================================================
# SECTION 10: MAIN ENTRY POINT
# =============================================================================
function Main {
    Write-Log "=================================================================="
    Write-Log "OTP Automation run started."
    Remove-OldLogs

    $outlook = $null
    try {
        $outlook = Get-OutlookApplication
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)   # 6 = olFolderInbox

        # Resolve the target inbox (supports non-default folder name override)
        if ($Cfg_InboxFolderName -and $Cfg_InboxFolderName -ne "Inbox") {
            $inbox = Get-OrCreateSubFolder -ParentFolder $inbox.Parent -FolderName $Cfg_InboxFolderName
        }

        $processedFolder = Get-OrCreateSubFolder -ParentFolder $inbox -FolderName $Cfg_ProcessedFolderName
        $failedFolder    = Get-OrCreateSubFolder -ParentFolder $inbox -FolderName $Cfg_FailedFolderName

        # Snapshot unread items BEFORE iterating - Outlook collections are
        # live, and moving items during enumeration can skip entries.
        $unreadItems = @($inbox.Items | Where-Object { $_.UnRead -eq $true -and $_.Class -eq 43 })
        # 43 = olMail (guards against meeting requests / receipts appearing unread)

        Write-Log "Found $($unreadItems.Count) unread mail item(s) in '$($inbox.Name)'."

        foreach ($item in $unreadItems) {
            Invoke-ProcessSingleMail -MailItem $item `
                                      -OutlookApp $outlook `
                                      -ProcessedFolder $processedFolder `
                                      -FailedFolder $failedFolder
        }

        Write-Log "OTP Automation run completed." -Level SUCCESS
    }
    catch {
        Write-Log "FATAL ERROR in main run: $($_.Exception.Message)" -Level ERROR
        if ($outlook) {
            Send-AlertMail -Subject "FATAL script error" `
                            -Body "otp-main.ps1 crashed: $($_.Exception.Message)" `
                            -OutlookApp $outlook
        }
        exit 1
    }
    finally {
        # Deliberately NOT calling Marshal.ReleaseComObject / quitting Outlook -
        # since we may have attached to the user's already-running Outlook
        # session. Forcibly releasing/quitting it would close their mailbox.
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Log "=================================================================="
    }
}

# Kick off
Main
