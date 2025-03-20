Import-Module "$PSScriptRoot\createTemplate.psm1"
Add-Type -AssemblyName PresentationFramework
[System.Reflection.Assembly]::LoadFrom("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll")
[System.Reflection.Assembly]::LoadFrom("C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll ") 

function getHours {
    param (
    [string]
    #String, mely az igényekből tartalmazza a műszak leírását
    $definition
    ) 
    $definition = $definition.Replace(' ', '')
    if ($definition -match '^\d{2}:\d{2}-\d{2}:\d{2}' -or
        $definition -match '^\d{1}:\d{2}-\d{2}:\d{2}' -or
        $definition -match '^\d{2}:\d{2}-\d{1}:\d{2}') {
        $start = $yesterday.Date
        $start = $start.AddHours($definition.Split("-")[0].Split(":")[0])
        $start = $start.AddMinutes($definition.Split("-")[0].Split(":")[1])

        $end = $yesterday.Date
        $end = $end.AddHours($definition.Split("-")[1].Split(":")[0])
        $end = $end.AddMinutes($definition.Split("-")[1].Split(":")[1])
        if ($end -lt $start) {
            $end = $end.AddDays(1)
        }
        $diff = $end.Subtract($start)
        $shift = [PSCustomObject]@{
            definition     = $start.ToString("HH:mm") + "-" +  $end.ToString("HH:mm")
            hours = ($diff.TotalMinutes % 60 -eq 0) ? $diff.Hours : $diff.TotalMinutes/60  
        }
        return $shift
    }
    elseif ($definition -eq "12óra") {
        getHours("08:30-20:30")
    }
    elseif ($definition -eq "2+4é") {
        getHours ('13:30-01:30')
    }
    elseif ([regex]::Matches($definition, '(?<=\().+?(?=\))' ).Success) {
        getHours(([regex]::Matches($definition, '(?<=\().+?(?=\))' ).Value))
    }
    else {
        throw "Rossz formátum"
    }
     <#
        .SYNOPSIS

        Értelmezi az igényben található munkaidő leírást

        .DESCRIPTION

        Párt képez, melíy tartalmazza, hogy mettől meddig tart egy műszak (String), és hogy az hány órát jelent (Int)
        Rekurzív függvény, mely első körben a leírás stringből kiszedi azt, hogy meddig tart egy műszak a következő formátumban ÓÓ:PP-ÓÓ:PP, majd második körben értelmezi azt.
    #>
}

function getIgenyekHelper {
    param (
        $Map,
        [pscustomobject]
        $igeny
    )
    $total = 0    
    $limit = $igeny.Worksheet.Cells.Find("Mindösszesen").row
    for ($i = 4; $i -lt $limit; $i++)
    {
        $hours = $igeny.Worksheet.Cells($i, 2 + $yesterday.Day).Value2
        if ($null -ne $hours)
        {   
            $role =  [string] $igeny.Worksheet.Cells($i, 1).MergeArea().Value2
            try {
                $pair = getHours $igeny.Worksheet.Cells($i, 2).Value2
            }
            catch {
                $igeny.Workbook.close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($igeny.Application)
                exit
            }
            for ($j = 0; $j -lt $hours / $pair.hours; $j++) {
                $Map[$role]+=,$pair
                $total += $pair.hours
            }                 
        }
    }
    if ($igeny.Worksheet.Cells($limit, 2 + $yesterday.Day).Value2 -ne $total) {
        $Map
        Write-Output "Valami nem jó, nem egyezik az elszámolásba beírt órák száma az igénnyel"         
        $igeny.Application.visible=$true
        [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "$total Valami nem jó, nem egyezik az elszámolásba beírt órák száma az igénnyel. `nKérlek nézd át!", "Hiba", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Error)
        exit
    }
    $igeny.Workbook.close($false)
}

function getIgenyek {
    param (
        $muszakok,
        [string]
        $UszomesterIgenyPath,
        [string]
        $FrontOfficeIgenyPath
    )
    $excelHelper = New-Object -comobject Excel.Application
    $excelHelper.visible = $false

    getIgenyekHelper  $muszakok @{
        "Application" = $excelHelper
        "Workbook"    = $excelHelper.Workbooks.Open($UszomesterIgenyPath)
        "Worksheet"   = $excelHelper.Workbooks.Open($UszomesterIgenyPath).Sheets.Item(1)
    } | Out-Null
    getIgenyekHelper $muszakok  @{
        "Application" = $excelHelper
        "Workbook"    = $excelHelper.Workbooks.Open($FrontOfficeIgenyPath)
        "Worksheet"   = $excelHelper.Workbooks.Open($FrontOfficeIgenyPath).Sheets.Item(1)
    } | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelHelper)
}

function open-DiakElszamolas {
    param (
        [boolean]
        #Teljesen legyen-e kitöltve az elszámolás
        $fillCompletely,
        [datetime] $yesterday = (Get-Date).AddDays(-1)
    )
    $base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\'
        $honap = $yesterday.ToString('MMMM')
        $elszamolas = "$($base)Diák elszámolás\"
        $saveAsDir= "$($elszamolas)$($yesterday.ToString('yyyy.MM')) - $honap\"
        if (-Not ([System.IO.Directory]::Exists($SaveAsDir)))
        {
            mkdir -p $SaveAsDir
        }
        $savaAsFurdo= "$($saveAsDir)$($yesterday.ToString('yyyy.MM.dd.')).xlsx"
    
        $igeny = [ordered]@{}
    
        $patternfu = '^F.*(' + [regex]::escape($honap) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month) +').*.*ürdő.*\.xlsx$'
        $items = Get-ChildItem C:\Users\Hirossport\Desktop\ | Where-Object {$_.Name.Normalize() -match ($patternfu)}| Where-Object Name -match ("^((?!Gazd).)*$") | Where-Object Name -match ("^((?!Front).)*$") | Sort-Object -Property Name -Descending | Select-Object -First 1
        $PathFU = ""
        if ($items.Length -ne 0)
        {
            $PathFU =  $items.FullName
        }
        elseif ($fillCompletely) {
            [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "Nem találom a Fürdő úszómester igényeit", "Hiba", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Error)
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory = "C:\Users\Hirossport\Desktop\"
            $OpenFileDialog.filter = "Minden Excel fájl|*.xls;*.xlsx;*.xlsm"
            $OpenFileDialog.Title = "Fürdő úszómester igényei"
            $OpenFileDialog.ShowDialog() |  Out-Null
            $items = Get-ChildItem C:\Users\Hirossport\Desktop\ | Where-Object Name -match $OpenFileDialog.SafeFileName
            $PathFU =  $items.FullName
        }
        $patternfp =  '^F.*(' + [regex]::escape($honap) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month) +').*Fürdő.*(azd|ront).*xlsx$'
        $items = Get-ChildItem C:\Users\Hirossport\Desktop\ | Where-Object {$_.Name.Normalize() -match ($patternfp)} | Sort-Object -Property Name -Descending | Select-Object -First 1
        $PathFP = ""
        if ($items.Length -ne 0)
        {
            $PathFP =  $items.FullName
        }
        elseif ($fillCompletely) {
            [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "Nem találom a Fürdő front office igényeit", "Hiba", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Error)
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory = "C:\Users\Hirossport\Desktop\"
            $OpenFileDialog.filter = "Minden Excel fájl|*.xls;*.xlsx;*.xlsm"
            $OpenFileDialog.Title = "Fürdő front ofice igényei"
            $OpenFileDialog.ShowDialog() |  Out-Null
            $items = Get-ChildItem C:\Users\Hirossport\Desktop\ | Where-Object Name -match $OpenFileDialog.SafeFileName
            $PathFP =  $items.FullName
        }
        if ("" -eq $PathFU -or  "" -eq $PathFP) {
            $empty = @{definition = ""; hours = ""}
            $igeny = [ordered] @{"Úszómester" = @($empty;$empty); "Csúszda" = @($empty); "Gyógymedence`nfelügyelet" = @($empty;$empty); "Öltöző szolgálat" = @($empty;$empty); Gyermekmegőrző = @($empty); "Pénztáros" = @($empty;$empty); Recepció = @($empty);}
        }
        else {
            getIgenyek $igeny $PathFU $PathFP | Out-Null
        }
        
        
        if (0 -ne $igeny.Count) {
            $furdosablon = Create-Template $igeny $fillCompletely $("Napi diákmunka elszámolás " + $yesterday.ToString("yyyy. MMMM dd."))
            if (-Not ([System.IO.File]::Exists($savaAsFurdo)))
            {
                $furdosablon.Workbook.SaveAs($savaAsFurdo,[Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
            }
            $furdosablon.Application.visible=$true
            #Set-ItemProperty -Path HKCU:\Software\Script -Name YesterdaysWorkingHours -value $Today.Day
        }
    
    
    <#
        $igeny = [ordered]@{}
        $patternsu = '^F.*(' + [regex]::escape($honap) + '|' + [regex]::escape((Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month)) +').*Strand.*\.xlsx$'
        $items = Get-ChildItem C:\Users\Hirossport\Desktop\ | Where-Object {$_.Name.Normalize() -match ($patternsu)} | Where-Object Name -match ("^((?!Gazd).)*$") | Where-Object Name -match ("^((?!Front).)*$") | Sort-Object -Property Name -Descending | Select-Object -First 1
        $PathSU=""
        if ($items.Length -ne 0)
        {
            $PathSU =  $items.FullName
        }
        elseif ($fillCompletely) {
            [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "Nem találom a Strand úszómester igényeit", "Hiba", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Error)
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory = "C:\Users\Hirossport\Desktop\"
            $OpenFileDialog.filter = "Minden Excel fájl|*.xls;*.xlsx;*.xlsm"
            $OpenFileDialog.Title = "Strand úszómester igényei"
            $OpenFileDialog.ShowDialog() |  Out-Null
            $items = Get-ChildItem C:\Users\Hirossport\Desktop\ | Where-Object Name -match $OpenFileDialog.SafeFileName
            $PathSU =  $items.FullName
        }
        $patternsp = '^F.*(' + [regex]::escape($honap) + '|' + [regex]::escape((Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month)) +').*Strand.*(azd|ront).*xlsx$'
        $items = Get-ChildItem C:\Users\Hirossport\Desktop\ | Where-Object {$_.Name.Normalize() -match ($patternsp)} | Sort-Object -Property Name -Descending | Select-Object -First 1
        $PathSP = ""
        if ($items.Length -ne 0)
        {
            $PathSP =  $items.FullName
        }
        elseif ($fillCompletely) {
            [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "Nem találom a Strand pénztár igényeit", "Hiba", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Error)
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory = "C:\Users\Hirossport\Desktop\"
            $OpenFileDialog.filter = "Minden Excel fájl|*.xls;*.xlsx;*.xlsm"
            $OpenFileDialog.Title = "Strand front office igényei"
            $OpenFileDialog.ShowDialog() |  Out-Null
            $items = Get-ChildItem C:\Users\Hirossport\Desktop\ | Where-Object Name -match $OpenFileDialog.SafeFileName
            $PathSP = $items.FullName
        }    
        $saveAsStrand= "$($saveAsDir)$($yesterday.ToString('yyyy.MM.dd.')) - STRAND.xlsx"  
        if ("" -eq $PathFU -or "" -eq $PathFP) {
            $empty = @{definition = ""; hours = ""}
            $igeny = [ordered] @{Úszómester = @($empty;$empty;$empty); Csúszda = @($empty); "Vízőr (pancsoló)" = @($empty); "Gyógymedence" = @($empty); "Öltöző női" = @($empty); "Öltöző férfi" = @($empty); Kisegítő = @($empty); Pénztáros = @($empty)}
        }
        else {
            getIgenyek $igeny $PathSU $PathSP | Out-Null
        }
        
        if (0 -ne $igeny.Count) {
            $strandsablon = Create-Template $igeny $fillCompletely $("Napi diákmunka elszámolás - STRAND - " + $yesterday.ToString("yyyy. MMMM dd."))
            if (-Not ([System.IO.File]::Exists($saveAsStrand)))
            {
                $strandsablon.Workbook.SaveAs($saveAsStrand,[Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
            }
            $strandsablon.Application.visible = $true
        }
        Set-ItemProperty -Path HKCU:\Software\Script -Name YesterdaysWorkingHours -value $Today.Day
    #>    
    }