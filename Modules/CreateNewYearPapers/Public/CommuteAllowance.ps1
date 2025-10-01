param(
    [int]$Year,
    [string]$DataSourcePath
)

# Load Word Interop assembly
[System.Reflection.Assembly]::LoadFrom(
    [System.Environment]::GetEnvironmentVariable("OfficeAssemblies_Word", [System.EnvironmentVariableTarget]::User)
) | Out-Null

# Load Word COM
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $word.Documents.Add()

# Manual line-break (Shift+Enter)
$lb = [char]11

# Title lines
$line1 = "Hírös Sport Nonprofit Kft."
$line2 = "Saját személygépkocsival történő"
$line3 = "Utazási költségtérítés elszámolása"

# Gray shading for title paragraph (15%)
$gray15 = 14277081

# --- Mail Merge Setup ---
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$DataSourcePath;Mode=Read;Extended Properties=`"HDR=YES;IMEX=1;`";"
$sqlStatement = "SELECT * FROM [Munka1$]"

$MailMerge = $doc.MailMerge
$MailMerge.OpenDataSource(
    $DataSourcePath,
    [ref][Microsoft.Office.Interop.Word.WdOpenFormat]::wdOpenFormatAuto,
    [ref]$false, # ConfirmConversions
    [ref]$true,  # ReadOnly
    [ref]$true,  # LinkToSource
    [ref]$false, # AddToRecentFiles
    [Type]::Missing, # PasswordDocument
    [Type]::Missing, # PasswordTemplate
    [ref]$false, # Revert
    [Type]::Missing, # WritePasswordDocument
    [Type]::Missing, # WritePasswordTemplate
    [ref]$connectionString, # Connection
    [ref]$sqlStatement,    # SQLStatement
    [Type]::Missing         # SQLStatement1
)

# Employee info field headers
$infoFields = @(
    "Munkavállaló neve",
    "Állandó lakóhelye",
    "Anyja neve",
    "Születési helye, ideje",
    "Gépjármű típusa",
    "Forgalmi rendszáma"
)

# --- Loop over months ---
for ($month=1; $month -le 12; $month++) {
    $selection = $word.Selection
    $selection.EndKey(6) # wdStory

    if ($month -gt 1) { $selection.InsertBreak(7) } # Page break

    # --- TITLE ---
    $fullTitle = $line1 + $lb + $line2 + $lb + $line3
    $selection.TypeText($fullTitle)
    $para = $selection.Paragraphs($selection.Paragraphs.Count)
    $rangeTitle = $para.Range
    $rangeTitle.Font.Name = "Times New Roman"
    $rangeTitle.Font.Size = 14
    $rangeTitle.Font.Bold = $true

    # Italic first line, AllCaps for second/third
    $rng1 = $rangeTitle.Duplicate; $rng1.Find.Text = $line1; if ($rng1.Find.Execute()) { $rng1.Font.Italic = $true }
    $rng2 = $rangeTitle.Duplicate; $rng2.Find.Text = $line2; if ($rng2.Find.Execute()) { $rng2.Font.AllCaps = $true }
    $rng3 = $rangeTitle.Duplicate; $rng3.Find.Text = $line3; if ($rng3.Find.Execute()) { $rng3.Font.AllCaps = $true }

    # Paragraph formatting
    $para.Alignment = 1                     # Centered
    $para.Format.LineSpacingRule = 0       # Single spacing
    $para.Format.SpaceBefore = 0
    $para.Format.SpaceAfter = 0

    # Borders (all sides)
    for ($i=1; $i -le 4; $i++) {
        $b = $para.Borders.Item($i)
        $b.LineStyle = 1
        $b.LineWidth = 2
        $b.Color = 0
    }

    # Shading
    $para.Shading.Texture = 0
    $para.Shading.ForegroundPatternColor = $gray15
    $para.Shading.BackgroundPatternColor = $gray15

    # --- MONTH HEADER ---
    $selection.TypeParagraph()
    $selection.Style = "Normál"
    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
    $selection.TypeText("$Year. $monthName")
    $monthPara = $selection.Paragraphs($selection.Paragraphs.Count)
    $monthPara.Range.Font.Name = "Times New Roman"
    $monthPara.Range.Font.Size = 11
    $monthPara.Range.Font.Bold = $true
    $monthPara.Range.Font.Italic = $false
    $monthPara.Format.SpaceBefore = 8
    $monthPara.Format.SpaceAfter = 8
    $monthPara.Alignment = 1   # Centered

   # --- FIRST EMPLOYEE INFO BLOCK ---
$selection.TypeParagraph()
$startRange = $selection.Range.Duplicate

$infoLabels = @(
    "A munkavállaló neve",
    "Állandó lakóhelye",
    "Anyja neve",
    "Születési helye, ideje",
    "Gépjármű típusa",
    "Forgalmi rendszáma"
)

$infoFields = @(
    "Munkavállaló_neve",
    "Állandó_lakhelye",
    "Anyja_neve",
    "Születési_helye_ideje",
    "Gépjármű_típusa",
    "Forgalmi_rendszáma"
)

for ($i = 0; $i -lt $infoLabels.Count; $i++) {
    $label = $infoLabels[$i]
    $field = $infoFields[$i]

    # Use single quotes outside so PowerShell won't try to interpolate $label:
    $selection.TypeText($label + ":`t")

    $para = $selection.Paragraphs($selection.Paragraphs.Count)
    $para.Alignment = 0
    $para.Format.LineSpacingRule = 0
    $para.Format.SpaceBefore = 0
    $para.Format.SpaceAfter = 0
    $para.TabStops.ClearAll()
    $para.TabStops.Add(4.5 * 28.35, 0) # Left-aligned

    # Insert merge field
    $range = $selection.Range.Duplicate
    $range.Collapse(0)
    $selection.Fields.Add($range, [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldMergeField, $field, $false)

    if ($i -lt $infoLabels.Count - 1) { $selection.TypeParagraph() }
}

# Close the first block explicitly before creating range
$selection.TypeParagraph()
$endRange = $selection.Range.Duplicate
$blockRange = $doc.Range($startRange.Start, $endRange.End)

# Font formatting
$blockRange.Font.Name = "Times New Roman"
$blockRange.Font.Size = 11

# Space after the first block
$lastPara = $blockRange.Paragraphs($blockRange.Paragraphs.Count)
$lastPara.Format.SpaceAfter = 16

# Outer border only
$blockRange.Borders.Enable = 0
foreach ($i in 1..4) {
    $border = $blockRange.Borders.Item($i)
    $border.LineStyle = 1
    $border.LineWidth = 2
    $border.Color = 0
}

# --- SECOND WORKPLACE INFO BLOCK ---
$selection.TypeParagraph()
$startRange2 = $selection.Range.Duplicate

$workLabels = @(
    "Munkahely címe",
    "Lakóhely címe",
    "A két közigazgatási egység közötti távolság",
    "Egy munkanapra eső költségtérítés",
    "Munkahelyemen munkavégzés céljából az alábbi napokon jelentem meg"
)

$workFields = @(
    "",                       # Literal text, type directly
    "Állandó_lakhelye",
    "távolság",
    "napi_térítés",
    ""                        # Manual marking, no merge field
)

$tabStops = @(4.5, 4.5, 7.5, 7.5, 4.5) # Tab stops in cm

for ($i = 0; $i -lt $workLabels.Count; $i++) {
    $label = $workLabels[$i]
    $field = $workFields[$i]
    $tabCm = $tabStops[$i]

    $selection.TypeText($label + ":`t")

    $para = $selection.Paragraphs($selection.Paragraphs.Count)
    $para.Alignment = 0
    $para.Format.LineSpacingRule = 0
    $para.Format.SpaceBefore = 0
    $para.Format.SpaceAfter = 0
    $para.TabStops.ClearAll()
    $para.TabStops.Add($tabCm * 28.35, 0) # Left-aligned

    if ($field -ne "") {
        $range = $selection.Range.Duplicate
        $range.Collapse(0)
        $selection.Fields.Add($range, [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldMergeField, $field, $false)
    } elseif ($i -eq 0) {
        $selection.TypeText("6000 Kecskemét, Csabay Géza krt. 5.")
    }

    if ($i -lt $workLabels.Count - 1) { $selection.TypeParagraph() }
}

# Close second block explicitly before creating range
$selection.TypeParagraph()
$endRange2 = $selection.Range.Duplicate
$blockRange2 = $doc.Range($startRange2.Start, $endRange2.End)

# Font formatting
$blockRange2.Font.Name = "Times New Roman"
$blockRange2.Font.Size = 11

# Space after second block
$lastPara2 = $blockRange2.Paragraphs($blockRange2.Paragraphs.Count)
$lastPara2.Format.SpaceAfter = 16

# Outer border only
$blockRange2.Borders.Enable = 0
foreach ($i in 1..4) {
    $border = $blockRange2.Borders.Item($i)
    $border.LineStyle = 1
    $border.LineWidth = 2
    $border.Color = 0
}



}
