[System.Reflection.Assembly]::LoadFrom("C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll ") 

function RGB($r, $g, $b) { $r + ($g * 256) + ($b * 256 * 256) }

function Create-Template {
    param (
        $muszakok,
        [bool]
        $fillMuszakok,
        [string]
        $header
    )
    
    $excel = New-Object -comobject Excel.Application
    $excel.visible =$false
    $workbook = $excel.Workbooks.Add()
    $workbook.WorkSheets.item(1).Name = "Munka1"
    $worksheet= $workbook.Sheets.Item(1)

    $MergeCells = $worksheet.Range(“A2:H2”)
    $MergeCells.Select() | Out-Null
    $MergeCells.MergeCells = $true 
    $MergeCells.Value2 = $header
    $worksheet.Range("A2").Font.Size=16
    $worksheet.Range("A2").Font.Name="Calibri"
    $worksheet.Range("A2").Font.Bold=$true
    $worksheet.Range("A2").HorizontalAlignment = [Microsoft.Office.Interop.Excel.XlHAlign]::xlHAlignCenter
    $worksheet.Range("A2").VerticalAlignment = [Microsoft.Office.Interop.Excel.XlVAlign]::xlVAlignCenter

    $worksheet.Range("B3:H3").RowHeight = 45.75
    $worksheet.Range("B3:H3").Font.Size=11
    $worksheet.Range("B3:H3").Font.Name="Calibri"
    $worksheet.Range("B3:H3").Font.Bold=$true
    $worksheet.Range("B3:H3").HorizontalAlignment = [Microsoft.Office.Interop.Excel.XlHAlign]::xlHAlignCenter
    $worksheet.Range("B3:H3").VerticalAlignment = [Microsoft.Office.Interop.Excel.XlVAlign]::xlVAlignCenter
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A3:H3").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic

    $worksheet.Range("B3").Value2 = "Név"
    $worksheet.Range("B3").ColumnWidth = 23
    $worksheet.Range("C3").Value2 = "Igényelt `n műszak `n kezd-vég."
    $worksheet.Range("C3").ColumnWidth = 11.7
    $worksheet.Range("D3").Value2 = "Igényelt `n óra"
    $worksheet.Range("D3").ColumnWidth = 8.5
    $worksheet.Range("E3").Value2 = "Teljesített `n műszak `n kezd-vég."
    $worksheet.Range("E3").ColumnWidth = 11.7
    $worksheet.Range("F3").Value2 = "Teljesített `n óra"
    $worksheet.Range("F3").ColumnWidth = 9.7
    $worksheet.Range("G3").Value2 = "Igazolt `n távollét"
    $worksheet.Range("G3").ColumnWidth = 9.7
    $worksheet.Range("H3").Value2 = "Igazolatlan `n távollét"
    $worksheet.Range("H3").ColumnWidth = 9.7
    $worksheet.Range("I3").ColumnWidth = 3.71
    $worksheet.Range("I:I").Font.Bold=$true
    $worksheet.Range("A4:H4").Font.Bold=$true
    $worksheet.Range("B4:H4").RowHeight = 21
    $worksheet.Range("A4:H4").Font.Size=11
    $worksheet.Range("A4:H4").Font.Name="Calibri"
    $worksheet.Range("A4").HorizontalAlignment = [Microsoft.Office.Interop.Excel.XlHAlign]::xlHAlignRight
    $worksheet.Range("B4:H4").HorizontalAlignment = [Microsoft.Office.Interop.Excel.XlHAlign]::xlHAlignCenter
    $worksheet.Range("A4:H4").VerticalAlignment = [Microsoft.Office.Interop.Excel.XlVAlign]::xlVAlignCenter
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlThin
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic 
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A4:H4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $worksheet.Range("A4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
    $worksheet.Range("A4").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
    $worksheet.Range("A4").ColumnWidth = 16.3
    $worksheet.Range("A4").Value2 = "Összesen"
    $worksheet.Range("A4").Interior.Color = RGB -r 217 -g 217 -b 217
    $worksheet.Range("G4").Font.Color = RGB -r 54 -g 96 -b 146
    $worksheet.Range("H4").Font.Color = RGB -r 255 -g 0 -b 0



    $row = 4

    foreach($key in $muszakok.keys) {
        for ($i = $muszakok[$key].Count - 1; $i -ge 0; $i-- ) {
            $worksheet.Range("A$row").EntireRow.Insert([Microsoft.Office.Interop.Excel.XlInsertShiftDirection]::xlShiftDown) | Out-Null
            $worksheet.Range("B$($row):C$row").Font.Bold=$false
            $worksheet.Range("E$row").Font.Bold=$false
            $worksheet.Range("G$row").Font.Color = RGB -r 54 -g 96 -b 146
            $worksheet.Range("H$row").Font.Color = RGB -r 255 -g 0 -b 0
            $worksheet.Range("B$row").RowHeight = 15.75  
            if ($fillMuszakok) {    
                $worksheet.Range("C$row").Value2 = $muszakok[$key][$i].definition
                $worksheet.Range("D$row").Value2 = [string] $muszakok[$key][$i].hours  
            }
            else {
                 $worksheet.Range("D$row").Formula = "=IFERROR(IF(NUMBERVALUE(TEXT(((RIGHT(C$row,LEN(C$row)-SEARCH(`"-`",C$row)))-(LEFT(C$row,SEARCH(`"-`",C$row,1)-1)))*24,`"normál`")) > 0,NUMBERVALUE(TEXT(((RIGHT(C$row,LEN(C$row)-SEARCH(`"-`",C$row)))-(LEFT(C$row,SEARCH(`"-`",C$row,1)-1)))*24,`"normál`")),NUMBERVALUE(TEXT(((RIGHT(C$row,LEN(C$row)-SEARCH(`"-`",C$row)))-(LEFT(C$row,SEARCH(`"-`",C$row,1)-1)))*24,`"normál`")) + 24),`"`")"
            }
            $worksheet.Range("F$row").Formula = "=IFERROR(IF(NUMBERVALUE(TEXT(((RIGHT(E$row,LEN(E$row)-SEARCH(`"-`",E$row)))-(LEFT(E$row,SEARCH(`"-`",E$row,1)-1)))*24,`"normál`")) > 0,NUMBERVALUE(TEXT(((RIGHT(E$row,LEN(E$row)-SEARCH(`"-`",E$row)))-(LEFT(E$row,SEARCH(`"-`",E$row,1)-1)))*24,`"normál`")),NUMBERVALUE(TEXT(((RIGHT(E$row,LEN(E$row)-SEARCH(`"-`",E$row)))-(LEFT(E$row,SEARCH(`"-`",E$row,1)-1)))*24,`"normál`")) + 24),`"`")"
            $worksheet.Range("I$row").Formula = "=IF(COUNTIF(J:J,B$row) > 0,`"EFO`",`"`")"
        }
        $worksheet.Range("A$($row):A$($row + $muszakok[$key].Count -1)").MergeCells = $true
        if ("Gyógymedence`nfelügyelet" -eq $key.trim() -and $muszakok[$key].Count -eq 1) {
            $worksheet.Range("A$row").Value2 = "Gyógymedence"
        }
        else {
            $worksheet.Range("A$row").Value2 = $key.trim()
        }      
        
        $worksheet.Range("A$row").Font.Bold=$true
        $worksheet.Range("A$row").HorizontalAlignment = [Microsoft.Office.Interop.Excel.XlHAlign]::xlHAlignRight
        $worksheet.Range("A$row").VerticalAlignment = [Microsoft.Office.Interop.Excel.XlVAlign]::xlVAlignCenter 
        $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideHorizontal).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
        $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideHorizontal).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlThin
        $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideHorizontal).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic  
        $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
        $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlThin
        $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic    
        $worksheet.Range("A$($row):A$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
        $worksheet.Range("A$($row):A$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
        $worksheet.Range("A$($row):A$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
        if ($key -ne $($muszakok.keys)[-1]) {
            $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlDouble
            
        }
        else {
            $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
            $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium
        }
        $worksheet.Range("A$($row):H$($row + $muszakok[$key].Count -1)").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).ColorIndex = [Microsoft.Office.Interop.Excel.XlColorIndex]::xlColorIndexAutomatic
        $row +=  $muszakok[$key].Count
    }

   
    $worksheet.Range("D$($row)").Value2 = "=SUM(D4:D$($row -1))"
    $worksheet.Range("F$($row)").Value2 = "=SUM(F4:F$($row -1))"
    $worksheet.Range("G$($row)").Value2 = "=SUM(G4:G$($row -1))"
    $worksheet.Range("H$($row)").Value2 = "=SUM(H4:H$($row -1))"
    $condition = $worksheet.Range("E$($row)").FormatConditions.Add([Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlExpression, 
     [Type]::Missing, 
     "=D$($row)<>SZUM(F$($row):H$($row))", [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing)
    $condition.Interior.Color = RGB -r 255 -g 0 -b 0
    
    $efo = $("Fazekas Éva", "Keller Dániel", "Karádiné Diószegi Anita", "Keresztes Alexandra", "Matuszka Máté", "Nagy Alexandra", "Sántha Tibor", "Tóth Anett", "Kolonics Anita", "Tóth Annamária Szilvia", "Domokos Enikő", "Magó Dorina", "Ölvödi Petra", "Tóth György", "Bukovics Ferenc", "Zsadányi Zsolt")
    $worksheet.Range("J1").Value2 = "EFO névsor"
    $row = 2
    foreach ($nev in $efo) {
        $worksheet.Range("J$row").Value2 = $nev
        $row++
    }
    $worksheet.Columns(10).Hidden=$true
    return [pscustomobject] @{
        "Application" = $excel
        "Workbook"    = $workbook
        "Worksheet"   = $worksheet
    }
    
}