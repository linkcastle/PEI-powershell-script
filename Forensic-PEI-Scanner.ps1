<#
.SYNOPSIS
    PEI Investigation Script for NAME SURNAME (Organized Output)
.DESCRIPTION
    Searches for suspicious files, emails, and communications with separate output folders.
    Creates structured directories for each evidence type.
.NOTES
    Run as Administrator for full access.
    Output saved to C:\PEI_Investigation_Results\ with subfolders.
#>

# --- CONFIGURATION ---
$userProfile = "C:\Users\NAME.SURNAME"
$outputFolder = "C:\PEI_Investigation_Results"
$startDate = "2025-01-01"
$endDate = "2025-09-30"

# Create organized subfolders
$folders = @{
    Files       = "$outputFolder\01_Files"
    Emails      = "$outputFolder\02_Emails"
    OneNote     = "$outputFolder\03_OneNote"
    Browser     = "$outputFolder\04_Browser"
    Reports     = "$outputFolder\05_Reports"
}

# Create all folders
foreach ($folder in $folders.Values) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
    }
}

# German & English keywords
$keywords = @(
    # Fraud/Payments (German)
    "Rechnung", "Faktura", "Zahlung", "Überweisung", "Gefälscht", "Scheinrechnung",
    "Kundenliste", "Bewerberdaten", "Persönlich", "Geld", "Konto", "IBAN",
    # Fraud/Payments (English)
    "invoice", "payment", "fake", "transfer", "client list", "candidate data",
    # Personal Relationship
    "NAME", "SURNAME", "persönlich", "intim", "Vorteil", "Deckung", "nicht weiterleiten",
    # Suspicious Actions
    "löschen", "vertuschen", "geheim", "off the books", "cash"
)

# File extensions to search
$extensions = @("*.pdf", "*.doc*", "*.xls*", "*.csv", "*.msg", "*.png", "*.jpg")

# --- SEARCH & COPY SUSPICIOUS FILES ---
function Search-Files {
    $results = @()
    $copiedFiles = 0

    foreach ($ext in $extensions) {
        Get-ChildItem -Path $userProfile -Recurse -File -Filter $ext |
        Where-Object { $_.LastWriteTime -ge $startDate -and $_.LastWriteTime -le $endDate } |
        ForEach-Object {
            foreach ($keyword in $keywords) {
                if (Select-String -Path $_.FullName -Pattern $keyword -Quiet -ErrorAction SilentlyContinue) {
                    $destPath = Join-Path $folders['Files'] $_.Name
                    if (-not (Test-Path $destPath)) {
                        Copy-Item -Path $_.FullName -Destination $destPath -Force -ErrorAction SilentlyContinue
                        $copiedFiles++
                    }

                    $results += [PSCustomObject]@{
                        FilePath = $_.FullName
                        Keyword  = $keyword
                        Size     = $_.Length / 1KB
                        Modified = $_.LastWriteTime
                        CopiedTo = $destPath
                    }
                }
            }
        }
    }

    $results | Export-Csv -Path "$($folders['Reports'])\Suspicious_Files.csv" -NoTypeInformation
    Write-Host " Copied $copiedFiles suspicious files to $($folders['Files'])" -ForegroundColor Green
    return $results
}

# --- SEARCH OUTLOOK EMAILS ---
function Search-OutlookEmails {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)  # 6 = Inbox
        $sent = $namespace.GetDefaultFolder(5)   # 5 = Sent Items
        $results = @()
        $savedEmails = 0

        # Process Inbox
        foreach ($mail in $inbox.Items) {
            if ($mail.ReceivedTime -ge $startDate -and $mail.ReceivedTime -le $endDate) {
                foreach ($keyword in $keywords) {
                    if ($mail.Subject -match $keyword -or $mail.Body -match $keyword) {
                        $emailFileName = "INBOX_$($mail.ReceivedTime.ToString('yyyyMMdd_HHmmss'))_$($mail.Subject.Replace(' ','_')).msg"
                        $emailFileName = $emailFileName -replace '[^a-zA-Z0-9_\-.]', '_'
                        $destPath = Join-Path $folders['Emails'] $emailFileName
                        $mail.SaveAs($destPath)
                        $savedEmails++

                        $results += [PSCustomObject]@{
                            Folder    = "Inbox"
                            Subject   = $mail.Subject
                            Sender    = $mail.SenderName
                            Keyword   = $keyword
                            Received  = $mail.ReceivedTime
                            HasAttachment = $mail.Attachments.Count -gt 0
                            SavedAs   = $destPath
                        }
                    }
                }
            }
        }

        # Process Sent Items
        foreach ($mail in $sent.Items) {
            if ($mail.SentOn -ge $startDate -and $mail.SentOn -le $endDate) {
                foreach ($keyword in $keywords) {
                    if ($mail.Subject -match $keyword -or $mail.Body -match $keyword) {
                        $emailFileName = "SENT_$($mail.SentOn.ToString('yyyyMMdd_HHmmss'))_$($mail.Subject.Replace(' ','_')).msg"
                        $emailFileName = $emailFileName -replace '[^a-zA-Z0-9_\-.]', '_'
                        $destPath = Join-Path $folders['Emails'] $emailFileName
                        $mail.SaveAs($destPath)
                        $savedEmails++

                        $results += [PSCustomObject]@{
                            Folder    = "Sent Items"
                            Subject   = $mail.Subject
                            Recipient = $mail.To
                            Keyword   = $keyword
                            Sent      = $mail.SentOn
                            HasAttachment = $mail.Attachments.Count -gt 0
                            SavedAs   = $destPath
                        }
                    }
                }
            }
        }
ù
        $outlook.Quit()
        $results | Export-Csv -Path "$($folders['Reports'])\Suspicious_Emails.csv" -NoTypeInformation
        Write-Host "Saved $savedEmails suspicious emails to $($folders['Emails'])" -ForegroundColor Green
        return $results
    }
    catch {
        Write-Warning "Outlook not accessible (run as Beni for full results)"
        "Outlook Error: $_" | Out-File "$($folders['Reports'])\Outlook_Error.txt"
        return $null
    }
}

# --- SEARCH BROWSER HISTORY ---
function Search-BrowserHistory {
    $chromeHistory = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\History"
    $edgeHistory = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\History"
    $results = @()

    # Copy and analyze Chrome history
    if (Test-Path $chromeHistory) {
        $chromeCopy = "$($folders['Browser'])\ChromeHistory.db"
        Copy-Item $chromeHistory $chromeCopy -Force -ErrorAction SilentlyContinue
        $results += Select-String -Path $chromeCopy -Pattern ($keywords -join "|") |
                   Select-Object Line, FileName, LineNumber
    }

    # Copy and analyze Edge history
    if (Test-Path $edgeHistory) {
        $edgeCopy = "$($folders['Browser'])\EdgeHistory.db"
        Copy-Item $edgeHistory $edgeCopy -Force -ErrorAction SilentlyContinue
        $results += Select-String -Path $edgeCopy -Pattern ($keywords -join "|") |
                   Select-Object Line, FileName, LineNumber
    }

    $results | Export-Csv -Path "$($folders['Reports'])\Browser_History_Hits.csv" -NoTypeInformation
    Write-Host " Browser history analysis saved to $($folders['Browser'])" -ForegroundColor Green
    return $results
}

# --- SEARCH ONENOTE ---
function Search-OneNote {
    $oneNotePath = "$env:LOCALAPPDATA\Microsoft\OneNote\16.0"
    $results = @()
    $copiedNotes = 0

    if (Test-Path $oneNotePath) {
        Get-ChildItem -Path $oneNotePath -Recurse -File -Filter "*.one" |
        ForEach-Object {
            foreach ($keyword in $keywords) {
                if (Select-String -Path $_.FullName -Pattern $keyword -Quiet -ErrorAction SilentlyContinue) {
                    $destPath = Join-Path $folders['OneNote'] $_.Name
                    if (-not (Test-Path $destPath)) {
                        Copy-Item -Path $_.FullName -Destination $destPath -Force -ErrorAction SilentlyContinue
                        $copiedNotes++
                    }

                    $results += [PSCustomObject]@{
                        NotePath = $_.FullName
                        Keyword  = $keyword
                        CopiedTo = $destPath
                    }
                }
            }
        }
    }

    $results | Export-Csv -Path "$($folders['Reports'])\OneNote_Hits.csv" -NoTypeInformation
    Write-Host " Copied $copiedNotes suspicious OneNote files to $($folders['OneNote'])" -ForegroundColor Green
    return $results
}

# --- RUN ALL SEARCHES ---
Write-Host " Starting PEI Investigation for Beni Assfalg..." -ForegroundColor Cyan

# 1. Search Files
Write-Host " Searching files..." -ForegroundColor Yellow
$fileResults = Search-Files

# 2. Search Emails
Write-Host " Searching Outlook emails..." -ForegroundColor Yellow
$emailResults = Search-OutlookEmails

# 3. Search Browser History
Write-Host " Searching browser history..." -ForegroundColor Yellow
$browserResults = Search-BrowserHistory

# 4. Search OneNote
Write-Host " Searching OneNote..." -ForegroundColor Yellow
$oneNoteResults = Search-OneNote

# --- GENERATE SUMMARY REPORT ---
$summary = @"
PEI Investigation Summary - Beni Assfalg
===========================================

Timeframe: $startDate to $endDate
User Profile: $userProfile

RESULTS:
-------------------
 Suspicious Files Found: $($fileResults.Count) → $($folders['Files'])
 Suspicious Emails Found: $($emailResults.Count) → $($folders['Emails'])
 Browser History Hits: $($browserResults.Count) → $($folders['Browser'])
 OneNote Hits: $($oneNoteResults.Count) → $($folders['OneNote'])

 All evidence organized in: $outputFolder
 Full reports saved to: $($folders['Reports'])

"@

$summary | Out-File -FilePath "$($folders['Reports'])\PEI_Summary.txt"
Write-Host $summary -ForegroundColor Cyan

Write-Host " Investigation complete! Check $outputFolder for organized results." -ForegroundColor Green
