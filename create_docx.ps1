# create_docx.ps1
# Script to retrieve images and populate a Word doc as a starting point for the 2-4 week outlook
# Smith july 2025

# Usage:
#

# Create COM object
$docx = New-Object -ComObject Word.Application
$docx.visible = $true # For testing. Set to $false for production

# New Doc
$word=$docx.Documents.Add()
$docDate = (Get-Date).ToString("yyyyMMdd")
$docPath = "$env:USERPROFILE\Documents\CIFFC_SPU_${docDate}_WK234.docx"

# ----- Functions -------

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
        if ($cellColors.ContainsKey($cell)) {
            $cellObj.Range.Font.Color = $cellColors[$cell]
        }
    }
    if ($caption -ne "") {
        $table.Range.InsertCaption("Table", ". ${caption}", 0, 0)
    }
    $table.Rows.Add() | Out-Null
    $range.InsertParagraphAfter()
}

# Add and resize image from web
function Add-Image {
    param (
        [string]$url,
        [string]$ref,
        [string]$caption = "",
        [int]$width = 300,
        [int]$height = 200
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
    if ($caption -ne "") {
        $range.InsertCaption("Figure", ". ${caption}", 0, 0)
    }
    $range.InsertParagraphAfter()
}

function getToday {
    $formattedDate = (Get-Date).ToString("MMMM dd yyyy")
    return $formattedDate
}

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

function getFriday {
    param (
        [int]$weeksAhead = 1
    )

    if ($weeksAhead -lt 1) {
        throw "Weeks ahead must be at least 1"
    }
     
     $today = Get-Date
     $targetDate = [DayofWeek]::Friday
     $daysUntilFriday = ([int]$targetDay - [int]$today.DayofWeek + 7) % 7
     if ($daysUntilFriday -eq 0) { $daysUntilFriday = 7}
     $thisFriday = $today.AddDays($daysUntilFriday)
     $targetFriday = $thisFriday.AddDays(7 *($weeksAhead - 1))

    return $targetFriday.ToString("MMMM dd")
}

# ---- End of functions -----

# --- Insert the title info -----
Add-Heading -text "$(getToday), Week 2/3/4 Significant Fire Weather Outlook"  -font "Arial" -size 18 -center $true -bold $true
Add-Heading -text "Week 2: $(getMonday -weeksahead 1) - $(getFriday -weeksAhead 2)" -font "Arial" -size 14 -center $true
Add-Heading -text "Week 3: $(getMonday -weeksahead 2) - $(getFriday -weeksAhead 3)" -font "Arial" -size 14 -center $true
Add-Heading -text "Week 4: $(getMonday -weeksahead 3) - $(getFriday -weeksAhead 4)" -font "Arial" -size 14 -center $true


# ----- Week 2 headings, tables and images -----------
Add-Heading -text "Week 2: $(getMonday -weeksahead 1) - $(getFriday -weeksAhead 2): " -font "Arial" -size 14 -center $false -bold $true
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()
Add-Heading -text "<placeholder text here>" -font "Arial" -size 11 -center $false -bold $false
$range = $word.Content
$range.Collapse(0)
$range.InsertParagraphAfter()

#Add-Image -url "https://www.tropicaltidbits.com/analysis/models/cfs-avg/2025072212/cfs-avg_z500aMean_namer_1.png" -width 450 -height 300 -ref="https://www.tropicaltidbits.com/"

Add-Table -rows 10 -cols 3 -caption "Week 2 risk summary" -cellContents @{
    "1,1" = "Little or no activity"
    "1,2" = "Placeholder text here"
    "1,3" = "Other placeholder text here"
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
    "1,1" = 0x00FF00 # Green
    "1,2" = 0xFFFF00 # Yellow
    "1,3" = 255 # Red
  <#   "4,1" = 0xFFA500 # Orange
    "5,1" = 0x0000FF # Blue
    "6,1" = 0x800080 # Purple
    "7,1" = 0xFFC0CB # Pink
    "8,1" = 0xA52A2A # Brown
    "9,1" = 0x808080 # Gray
    "10,1" = 0xFFFFFF # White #>
}

$word.SaveAs([ref]$docPath)
<# $word.Close()
$docx.Quit() #>

Write-Host "Week 234 outlook saved to $docPath"