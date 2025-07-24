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

# ----- Functions -------

# Add a heading
function Add-Heading {
    param (
        [string]$text,
        [int]$level=1,
        [string]$font = "Arial",
        [int]$size = 18
    )
    $range = $word.Content
    $range.Collapse(0) # Collapse to the end of the document
    $range.Text = "$text`r"
    $range.Font.Name = $font
    $range.Font.Size = $size
    $range.set_Style("Heading $level")
    $ange.InsertParagraphAfter()
}
