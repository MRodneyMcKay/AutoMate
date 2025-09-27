<#  
    This file is part of AutoMate.  

    AutoMate is free software: you can redistribute it and/or modify  
    it under the terms of the GNU General Public License as published by  
    the Free Software Foundation, either version 3 of the License, or  
    (at your option) any later version.  

    This program is distributed in the hope that it will be useful,  
    but WITHOUT ANY WARRANTY; without even the implied warranty of  
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the  
    GNU General Public License for more details.  

    You should have received a copy of the GNU General Public License  
    along with this program. If not, see <https://www.gnu.org/licenses/>.  
#>

function New-TIGAttendanceSheet {
    [CmdletBinding()]
    param(
        [int]$Year,

        [string]$CsvPath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\TIG csoportok.csv",

        [string]$OutputPath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\TIG jelenléti ív_$Year.docx"
    )

    # Convert cm to points
    $cellPaddingPts = 1  # 0.05 cm ≈ 1 pt

    # Load Word COM
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()

    # Set margins
    $doc.PageSetup.LeftMargin   = [int][math]::Round(2 * 28.35)  # 2 cm
    $doc.PageSetup.RightMargin  = [int][math]::Round(2 * 28.35)  # 2 cm
    $doc.PageSetup.BottomMargin = [int][math]::Round(0.52 * 28.35)  # 0.52 cm

    # Open mail merge source
    $doc.MailMerge
    $doc.MailMerge.OpenDataSource($CsvPath)

    # Days of week mapping
    $daysOfWeek = @("V","H","K","Sze","Cs","P","Szo")  # Sunday=V

    for ($month = 1; $month -le 12; $month++) {
        $selection = $word.Selection
        $selection.EndKey(6) # wdStory

        if ($month -gt 1) { $selection.InsertBreak(7) } # wdPageBreak

        # Header
        $selection.TypeText("Teljesítésigazolás – jelenléti ív")
        $selection.Style = "Cím"
        $selection.TypeParagraph()

        $selection.TypeText("A Kecskeméti Fürdőben megjelent csoportok létszámáról")
        $selection.Style = "Címsor 1"
        $selection.TypeParagraph()

        # Merge field
        $selection.TypeText("Csoport neve: ")
        $range = $selection.Range
        $range.Collapse(0)
        $selection.Fields.Add($range, [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldMergeField, "Csoportok", $false)
        $selection.TypeParagraph()

        # Month/year
        $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
        $selection.TypeText("Igénybevétel ideje: $Year. $monthName")
        $selection.TypeParagraph()

        # Table
        $daysInMonth = [DateTime]::DaysInMonth($Year, $month)
        $table = $doc.Tables.Add($selection.Range, $daysInMonth + 1, 5)
        $table.Borders.Enable = 1
        $table.TopPadding = $cellPaddingPts
        $table.BottomPadding = $cellPaddingPts
        $table.LeftPadding = $cellPaddingPts
        $table.RightPadding = $cellPaddingPts

        # Center table horizontally
        $table.Rows.Alignment = [Microsoft.Office.Interop.Word.WdRowAlignment]::wdAlignRowCenter

        # Table header
        $headers = @("Dátum","Nap","Létszám","Igénybe vevő képviselő aláírása","K. Fürdő képviselő aláírása")
        for ($col=1; $col -le 5; $col++) {
            $cell = $table.Cell(1,$col)
            $cell.Range.Text = $headers[$col-1]
            $cell.VerticalAlignment = 1
            $cell.Range.ParagraphFormat.SpaceBefore = 0
            $cell.Range.ParagraphFormat.SpaceAfter = 0
            $cell.Range.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpaceSingle
            $cell.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
        }

        # Table body
        for ($day=1; $day -le $daysInMonth; $day++) {
            $row = $day + 1
            $date = Get-Date -Year $Year -Month $month -Day $day

            # Correct day mapping snippet
            $dayStr = "{0:D2}" -f $day
            $dowStr = "$($daysOfWeek[$date.DayOfWeek.value__])"
            $values = @($dayStr, $dowStr, "", "", "")

            for ($col=1; $col -le 5; $col++) {
                $cell = $table.Cell($row,$col)
                $cell.Range.Text = $values[$col-1]
                $cell.VerticalAlignment = 1
                $cell.Range.ParagraphFormat.SpaceBefore = 0
                $cell.Range.ParagraphFormat.SpaceAfter = 0
                $cell.Range.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpaceSingle
                $cell.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter

                if ($dowStr -eq "V") { $cell.Shading.BackgroundPatternColor = 12632256 } # Sunday 25% gray
            }
        }

        # Auto-fit columns
        $table.AutoFitBehavior([Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent)
        foreach ($col in $table.Columns) {
            $col.PreferredWidthType = [Microsoft.Office.Interop.Word.WdPreferredWidthType]::wdPreferredWidthAuto
            $col.AutoFit()
        }
    }

    # Save and quit
    $doc.SaveAs([ref]$OutputPath)
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
