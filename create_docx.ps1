
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
        # Shade first row and first column light grey, and set bold/non-italic
        if ($r -eq 1 -or $c -eq 1) {
            $cellObj.Shading.BackgroundPatternColor = 12632256
            $cellObj.Range.Font.Bold = $true
            $cellObj.Range.Font.Italic = $false
        } else {
            $cellObj.Range.Font.Bold = $false
            $cellObj.Range.Font.Italic = $false
        }
    }
    if ($caption -ne "") {
        $table.Range.InsertCaption("Table", ". ${caption}", 0, 0)
    }
    #$table.Rows.Add() | Out-Null
    $range.InsertParagraphAfter()
}

# Add and resize image from web. Add a caption that includes the source URL and retrieval date.
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

# Add two images side by side and add a single caption
# The Word COM object doesn't appear to support inserting, cropping and resizing images in this manner, so any cropping will have to be done manually in Word after the document is created.
function Add-ImagesSideBySide {
    param (
        [string]$url1,
        [string]$ref1,
        [string]$url2,
        [string]$ref2,
        [int]$width = 300,
        [int]$height = 200,
        [string]$caption = "I forgot to insert a caption. Please edit me."
    )
    # Download temp files
    $tempImage1 = Join-Path $env:TEMP ("image_" + [guid]::NewGuid().ToString() + "_1.png")
    $tempImage2 = Join-Path $env:TEMP ("image_" + [guid]::NewGuid().ToString() + "_2.png")
    Invoke-WebRequest -Uri $url1 -OutFile $tempImage1 -Headers @{
        "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0 Safari/537.36"
        "Referer" = $ref1
        "Accept" = "image/webp,image/apng,image/*,*/*;q=0.8"
    }
    Invoke-WebRequest -Uri $url2 -OutFile $tempImage2 -Headers @{
        "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0 Safari/537.36"
        "Referer" = $ref2
        "Accept" = "image/webp,image/apng,image/*,*/*;q=0.8"
    }
    $range = $word.Content
    $range.Collapse(0)
    # Insert a 1-row, 2-column table
    $table = $word.Tables.Add($range, 1, 2)
    $table.Rows.Height = $height
    $table.Columns.Width = $width
    # Insert first image
    $cell1 = $table.Cell(1,1)
    $img1 = $cell1.Range.InlineShapes.AddPicture($tempImage1)
    $img1.LockAspectRatio = $true
    $img1.Width = $width
    $img1.Height = $height
    # Insert second image
    $cell2 = $table.Cell(1,2)
    $img2 = $cell2.Range.InlineShapes.AddPicture($tempImage2)
    $img2.LockAspectRatio = $true
    $img2.Width = $width
    $img2.Height = $height
    # Add caption below the table
    $table.Range.Collapse(0)
    if ($caption -ne "") {
        $table.Range.InsertCaption("Figure", ". ${caption}", 0, 0)
    }
    $table.Range.InsertParagraphAfter()
}

function getToday {
    $formattedDate = (Get-Date).ToString("MMMM dd yyyy")
    return $formattedDate
}

# Get the next Monday or Sunday date, optionally weeks ahead
# weeksAhead = 1 for next week, 2 for the week after, etc.
function getMonday {
param (
    [int]$weeksAhead = 1,
    [string]$format = "MMMM dd"
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
    return $targetMonday.ToString($format)
}

# Get the next Sunday date, optionally weeks ahead
# weeksAhead = 1 for next week, 2 for the week after, etc.
function getSunday {
param (
    [int]$weeksAhead = 1,
    [string]$format = "MMMM dd"
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

    return $targetSunday.ToString($format)
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

# NAEFS POP over time. The init date can be the current date, but the range depends on the day of the week. There's a hard coded assumption that this will be run on a Tuesday, hence hours will be from hour 168 to hour 312. These as well as the thresholds used can be modifed below as needed.
$hourStart = 168
$hourEnd = 312
$hourInit = 00
$datetimeInit = (Get-Date).ToString("yyyyMMdd") + $hourInit.ToString("D2")
$threshold1=10
$threshold2=25
$url1="https://collaboration.cmc.ec.gc.ca/cmc/ensemble/produits/data/produits/PR-1ACC/GT0.0$threshold1/CMC_NCEP/${datetimeInit}_${hourStart}_${hourEnd}.gif"
$url2="https://collaboration.cmc.ec.gc.ca/cmc/ensemble/produits/data/produits/PR-1ACC/GT0.0$threshold2/CMC_NCEP/${datetimeInit}_${hourStart}_${hourEnd}.gif"
$ref1 = "https://collaboration.cmc.ec.gc.ca/cmc/ensemble/produits/index_f.html"
$ref2 = "https://collaboration.cmc.ec.gc.ca/cmc/ensemble/produits/index_f.html"

Add-ImagesSideBySide -width 250 -height 200 -url1 $url1 -ref1 $ref1 -url2 $url2 -ref2 $ref2 -caption "NAEFS Probability of at least ${threshold1}mm (left) and ${threshold2}mm (right) of total precipitation for Week 2. Retrieved $(getToday). Source: CMC."

$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

# CWFIS BUI and HFI for the end of week 2. Need the URLs, the current year and the date of the final day of week 2 in order to retrieve the images
$nextSunday = (getSunday -weeksAhead 2 -format "yyyyMMdd")
$thisYear = (Get-Date).ToString("yyyy")
$url1="https://cwfis.cfs.nrcan.gc.ca/data/maps/fwi_fbp/${thisYear}/xf/bui${nextSunday}.png"
$url2="https://cwfis.cfs.nrcan.gc.ca/data/maps/fwi_fbp/${thisYear}/xf/hfi${nextSunday}.png"

Add-ImagesSideBySide -width 250 -height 200 -url1 $url1 -ref1 "https://cwfis.cfs.nrcan.gc.ca" -url2 $url2 -ref2 "https://cwfis.cfs.nrcan.gc.ca" -caption "Projected BUI (left) and HFI (right) valid ${nextSunday}. Retrieved $(getToday). Source: CWFIS."

$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

# Leave a space and prompt for the week 2 annotated map
Add-Heading -text "<Insert week 2 annotated map of Canada here>" -font "Arial" -size 11 -center $false -bold $false
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

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
Add-Heading -text "Week 2 trends:" -font "Arial" -size 11 -center $false -bold $true
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
Add-Heading -text "Week 3 trends:" -font "Arial" -size 11 -center $false -bold $true
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
Add-Heading -text "Week 2-4 Summary: $(getMonday -weeksahead 1) - $(getSunday -weeksAhead 4): " -font "Arial" -size 11 -center $false -bold $true
Add-Heading -text "<placeholder text here>" -font "Arial" -size 11 -center $false -bold $false 
Add-Heading -text "<text here>" -font "Arial" -size 11 -center $false -bold $false 


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
