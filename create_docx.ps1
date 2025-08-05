# create_docx.ps1
# Script to retrieve images and populate a Word doc as a starting point for the 2-4 week outlook
# Smith july 2025

# Usage (Tested on Windows 11 with Office 365):
# 1. A VPN connection is required to grab Lin's images. There's a helpful popup dialog box to remind you, and it can be commented out if you find it annoying.
# 2. Save this script locally and run it. No admin permissions are required.
# 3. The script will create a Word document in your local (not OneDrive) Documents folder with the name CIFFC_SPU_<date>_WK234.docx
# 4. It's possible to have most of this run the in background, but you'll still need to add a classification before it's saved.

# Show dialog box for VPN requirement. Add a block comment if you find this annoying.
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.MessageBox]::Show("This script requires an active connection to the VPN to download images. Click OK to continue.", "VPN Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Create COM object
$docx = New-Object -ComObject Word.Application
$docx.visible = $true # Set to $false if you don't want to see the word window, though you'll still have to add a classification before it's saved.


# New Doc
$word=$docx.Documents.Add()
$docDate = (Get-Date).ToString("yyyyMMdd")
$docPath = "$env:USERPROFILE\Documents\CIFFC_SPU_${docDate}_WK234.docx"

# Insert page numbers in the footer (right-aligned)
$footer = $word.Sections.Item(1).Footers.Item(1)
$footerRange = $footer.Range
$footerRange.Text = "" # Clear any existing text
$footerRange.ParagraphFormat.Alignment = 2 # 2 = wdAlignParagraphRight
# 33 = wdFieldPage
$footerRange.Fields.Add($footerRange, 33) | Out-Null

# ----- Functions ----------------------------------------------------

# Add a heading
function Add-Heading {
    param (
        [string]$text,
        [int]$level = 2,
        [string]$font = "Arial",
        [int]$size = 18,
        [bool]$center = $false,
        [bool]$bold = $false
    )
    $range = $word.Content
    $range.Collapse(0) # Collapse to the end of the document
    $range.Text = "$text`r"
    $range.Font.Name = $font
    $range.Font.Size = $size
    if ($center) {
        $range.ParagraphFormat.Alignment = 1 # wdAlignParagraphCenter
    }
    if ($bold) {
        $range.Font.Bold = $true
    }
}

# Insert table
# Cells are referenced as (row,column) e.g. (1,2) is row 1, column 2
function Add-Table {
    param (
        [int]$rows,
        [int]$cols,
        [hashtable]$cellContents = @{},
        [string]$caption = "",
        [hashtable]$cellColors = @{} # e.g. @{ "1,1" = 255; "1,2" = 0x0000FF }
    )
    $range = $word.Content
    $range.Collapse(0)
    $table = $word.Tables.Add($range, $rows, $cols)
    $table.Range.Font.Name = "Arial"
    $table.Range.Font.Size = 10
    $table.Borders.Enable = $true
    foreach ($cell in $cellContents.Keys) {
        $indices = $cell -split ","
        $r = [int]$indices[0]
        $c = [int]$indices[1]
        $cellObj = $table.Cell($r, $c)
        $cellObj.Range.Text = $cellContents[$cell]
        $cellObj.Range.ParagraphFormat.Alignment = 1 # wdAlignParagraphCenter
        # Set text color if specified
        if ($cellColors.ContainsKey($cell)) {
            $cellObj.Range.Font.Color = $cellColors[$cell]
        }
        # Set background to light grey (RGB 242,242,242)
        $cellObj.Shading.BackgroundPatternColor = 12632256
    }
    if ($caption -ne "") {
        $table.Range.InsertCaption("Table", ". ${caption}", 0, 0)
    }
    #$table.Rows.Add() | Out-Null
    $range.InsertParagraphAfter()
}

# Add and resize image from web. Add a captoion that includes the source URL and retrieval date.
function Add-Image {
    param (
        [string]$url,
        [string]$ref,
        [string]$caption = "I forgot to insert a caption. Please edit me.",
        [int]$width = 300,
        [int]$height = 200,
        [double]$cropBottomPercent = 0 # 0-1, e.g. 0.15 for 15% crop from bottom
    )
    # download temp file
    $tempImage = Join-Path $env:TEMP ("\image_" + [guid]::NewGuid().ToString() + ".png" )
    # We need to be a little greasy here as Tropical tidbits blocks some automated requests.
    Invoke-WebRequest -Uri $url -OutFile $tempImage -Headers @{
    "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0 Safari/537.36"
    "Referer" = $ref
    "Accept" = "image/webp,image/apng,image/*,*/*;q=0.8"
}

    $range = $word.Content
    $range.Collapse(0)
    $shape = $range.InlineShapes.AddPicture($tempImage)
    $shape.LockAspectRatio = $true
    $shape.Width = $width
    $shape.Height = $height
    # Crop bottom if requested
    if ($cropBottomPercent -gt 0 -and $cropBottomPercent -lt 1) {
        $cropAmount = $shape.Height * $cropBottomPercent
        $shape.PictureFormat.CropBottom = $cropAmount
    }
    if ($caption -ne "") {
        $sourceText = "Retrieved on $(Get-Date -Format 'yyyy-MM-dd'): $url"
        $fullCaption = "$caption ($sourceText)"
        $range.Collapse(0)
        $range.InsertCaption("Figure", ". ${fullCaption}", 0, 0)
    }
    $range.InsertParagraphAfter()
}

function getToday {
    $formattedDate = (Get-Date).ToString("MMMM dd yyyy")
    return $formattedDate
}

# Get the next Monday or Sunday date, optionally weeks ahead
# weeksAhead = 1 for next week, 2 for the week after, etc.
function getMonday {
    param (
        [int]$weeksAhead = 1
    )

    if ($weeksAhead -lt 1) {
        throw "Weeks ahead must be at least 1"
    }

    $today = Get-Date
    $targetDay = [DayofWeek]::Monday
    $daysUntilMonday = ([int]$targetDay - [int]$today.DayofWeek + 7) % 7
    if ($daysUntilMonday -eq 0) { $daysUntilMonday = 7 }
    $thisMonday = $today.AddDays($daysUntilMonday)
    $targetMonday = $thisMonday.AddDays(7 * ($weeksAhead - 1))
    return $targetMonday.ToString("MMMM dd")
}

# Get the next Sunday date, optionally weeks ahead
# weeksAhead = 1 for next week, 2 for the week after, etc.
function getSunday {
    param (
        [int]$weeksAhead = 1
    )

    if ($weeksAhead -lt 1) {
        throw "Weeks ahead must be at least 1"
    }
     
     $today = Get-Date
     $targetDay = [DayofWeek]::Sunday
     $daysUntilSunday = ([int]$targetDay - [int]$today.DayofWeek + 7) % 7
     if ($daysUntilSunday -eq 0) { $daysUntilSunday = 7}
     $thisSunday = $today.AddDays($daysUntilSunday)
     $targetSunday= $thisSunday.AddDays(7 *($weeksAhead - 1))

    return $targetSunday.ToString("MMMM dd")
}

# Create the base URL for TT CFSv2 images:
# The images have the pattern: https://www.tropicaltidbits.com/analysis/models/cfs-avg/<YYYYMMDDHH>/cfs-avg_z500aMean_namer_<1|2|3|4>.png where <YYYYMMDDHH> is the initialization time
# and <1|2|3|4} is the week. For consistency we'll use the Monday 12z run so that images roughly line up with the M-S forecast week.
function ttURL {
    param (
        [int]$week = 2
    )

    if ($week -lt 1 -or $week -gt 4) {
        throw "Week must be between 1 and 4"
    }

    # Find the Monday of the current week
    $today = Get-Date
    $daysSinceMonday = ([int]$today.DayOfWeek - [int][System.DayOfWeek]::Monday + 7) % 7
    $monday = $today.AddDays(-$daysSinceMonday)
    # Use 12Z (12:00 UTC) as the initialization hour
    $dateStr = $monday.ToString("yyyyMMdd") + "12"
    $baseUrl = "https://www.tropicaltidbits.com/analysis/models/cfs-avg/$dateStr/cfs-avg_z500aMean_namer_${week}.png"
    return $baseUrl
}

# Helper to add hyperlink to a found name
function Add-InlineMailto {
    param($doc, $name, $email)
    $find = $doc.Content.Find
    $find.Text = $name
    $find.Forward = $true
    $find.Wrap = 1 # wdFindContinue
    if ($find.Execute()) {
        $foundRange = $find.Parent.Duplicate
        $doc.Hyperlinks.Add($foundRange, "mailto:$email", $null, $null, $name, $null) | Out-Null
    }
}


# ---- End of functions ------------------------------------------------------------------

# --- Insert the title info -----
Add-Heading -text "$(getToday), Week 2/3/4 Significant Fire Weather Outlook"  -font "Arial" -size 18 -center $true -bold $true
Add-Heading -text "Week 2: $(getMonday -weeksahead 1) - $(getSunday -weeksAhead 2)" -font "Arial" -size 14 -center $true
Add-Heading -text "Week 3: $(getMonday -weeksahead 2) - $(getSunday -weeksAhead 3)" -font "Arial" -size 14 -center $true
Add-Heading -text "Week 4: $(getMonday -weeksahead 3) - $(getSunday -weeksAhead 4)" -font "Arial" -size 14 -center $true


# ----- Week 2 headings, tables and images -----------
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()
Add-Heading -text "Week 2: $(getMonday -weeksahead 1) - $(getSunday -weeksAhead 2): " -font "Arial" -size 11 -center $false -bold $true
Add-Heading -text "<placeholder text here>" -font "Arial" -size 11 -center $false -bold $false
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()


$ttWeek2 = ttURL -week 2
Add-Image -url $ttWeek2 -width 450 -height 300 -ref "https://www.tropicaltidbits.com/" -caption "Week 2 500 hPa mean forecast from CFSv2. "

$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

Add-Image -url "https://hpfx.science.gc.ca/~lin001/forecastsMon/combine-2.jpeg" -width 650 -height 600 -ref "https://hpfx.science.gc.ca" -caption "GEPS Week 2 2m temperature and precipitation anomalies (top) and probabilities (bottom)." -cropBottomPercent 0.35

Add-Table -rows 10 -cols 3 -caption "Week 2 risk summary" -cellContents @{
    "1,1" = "Geographic area"
    "1,2" = "Precipitation"
    "1,3" = "Temperature"
    "2,1" = "BC"
    "3,1" = "YT"
    "4,1" = "NWT"
    "5,1" = "AB"
    "6,1" = "SK"
    "7,1" = "MB"
    "8,1" = "ON"
    "9,1" = "PQ"
    "10,1" = "ATL"
}  -cellColors @{
    "1,2" = 0xFFA500 # Orange
    "1,3" = 255 # Red
}

$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()
Add-Heading -text "Week 2 trends:" -font "Arial" -size 11 -center $false -bold $false
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

# ----- Week 3 headings, tables and images -----------
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()
Add-Heading -text "Week 3: $(getMonday -weeksahead 2) - $(getSunday -weeksAhead 3): " -font "Arial" -size 11 -center $false -bold $true
Add-Heading -text "<placeholder text here>" -font "Arial" -size 11 -center $false -bold $false
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

$ttWeek3 = ttURL -week 3
Add-Image -url $ttWeek3 -width 450 -height 300 -ref "https://www.tropicaltidbits.com/" -caption "Week 3 500 hPa mean forecast from CFSv2. "

Add-Image -url "https://hpfx.science.gc.ca/~lin001/forecastsMon/combine-3.jpeg" -width 650 -height 600 -ref "https://hpfx.science.gc.ca" -caption "GEPS Week 3 2m temperature and precipitation anomalies (top) and probabilities (bottom)." -cropBottomPercent 0.35

Add-Table -rows 10 -cols 3 -caption "Week 3 risk summary" -cellContents @{
    "1,1" = "Geographic area"
    "1,2" = "Precipitation"
    "1,3" = "Temperature"
    "2,1" = "BC"
    "3,1" = "YT"
    "4,1" = "NWT"
    "5,1" = "AB"
    "6,1" = "SK"
    "7,1" = "MB"
    "8,1" = "ON"
    "9,1" = "PQ"
    "10,1" = "ATL"
}  -cellColors @{
    "1,2" = 0xFFA500 # Orange
    "1,3" = 255 # Red
}

$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()
Add-Heading -text "Week 3 trends:" -font "Arial" -size 11 -center $false -bold $false
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

# ----- Week 4 headings, tables and images -----------
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()
Add-Heading -text "Week 4: $(getMonday -weeksahead 3) - $(getSunday -weeksAhead 4): " -font "Arial" -size 11 -center $false -bold $true
Add-Heading -text "<placeholder text here>" -font "Arial" -size 11 -center $false -bold $false
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

$ttWeek4 = ttURL -week 4
Add-Image -url $ttWeek4 -width 450 -height 300 -ref "https://www.tropicaltidbits.com/" -caption "Week 4 500 hPa mean forecast from CFSv2. "

Add-Image -url "https://hpfx.science.gc.ca/~lin001/forecastsMon/combine-4.jpeg" -width 650 -height 600 -ref "https://hpfx.science.gc.ca" -caption "GEPS Week 4 2m temperature and precipitation anomalies (top) and probabilities (bottom)." -cropBottomPercent 0.35

Add-Table -rows 10 -cols 3 -caption "Week 4 risk summary" -cellContents @{
    "1,1" = "Geographic area"
    "1,2" = "Precipitation"
    "1,3" = "Temperature"
    "2,1" = "BC"
    "3,1" = "YT"
    "4,1" = "NWT"
    "5,1" = "AB"
    "6,1" = "SK"
    "7,1" = "MB"
    "8,1" = "ON"
    "9,1" = "PQ"
    "10,1" = "ATL"
}  -cellColors @{
    "1,2" = 0xFFA500 # Orange
    "1,3" = 255 # Red
}
# -----Summary text for all 3 weeks -----------
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()
Add-Heading -text "Week 2-4 Summary: $(getMonday -weeksahead 1) - $(getSunday -weeksAhead 4): " -font "Arial" -size 11 -center $false -bold $true
Add-Heading -text "<placeholder text here>" -font "Arial" -size 11 -center $false -bold $false


# Add contact info as plain text, then convert names to hyperlinks using Find
$range = $word.Content
$range.Collapse(0)
$contactLine = "If you have any questions about the forecast, interpretations or terminology please contact one or all of the WIPS weather team: Richard Carr, Liam Buchart or Mike Smith."
$range.Text = $contactLine
$range.Font.Name = "Arial"
$range.Font.Size = 11
$range.Font.Bold = $false
$range.ParagraphFormat.Alignment = 0
$range.InsertParagraphAfter()


# Add hyperlinks for each name
Add-InlineMailto $word "Richard Carr" "richard.carr@nrcan-rnca.gc.ca"
Add-InlineMailto $word "Liam Buchart" "liam.buchart@nrcan-rnca.gc.ca"
Add-InlineMailto $word "Mike Smith" "michael.smith2@nrcan-rnca.gc.ca"

# disclaimer
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()
Add-Heading -text "Disclaimer: " -font "Arial" -size 11 -center $false -bold $true
Add-Heading -text "This document is intended for internal use by CIFFC and its partners. This outlook is updated once per week and amendments are not issued. It provides a high-level outlook of significant fire weather conditions for the next 2-4 weeks based on the latest forecast model data. The information is subject to change as new data becomes available and should not be used for operational decision-making. The focus of this outlook is on meteorological conditions pertinent to wildfire behavior, and other potentially high-impact weather is not considered or included." -font "Arial" -size 11 -center $false -bold $false

$word.SaveAs([ref]$docPath)
# Uncomment the 2 lines below if you want to close the document after it's created and saved.
<# $word.Close()
$docx.Quit() #>

Write-Host "Week 234 outlook saved to $docPath"
