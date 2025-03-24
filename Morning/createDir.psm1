function create-Directories {
    param (
        [datetime]$Today = (Get-Date)
    ) 
    $Monday = $Today.AddDays($Today.DayOfWeek.value__ -eq 0 ? -6 : 1 - $Today.DayOfWeek.value__) 
    $Sunday = $Monday.AddDays(6)
    $Week = $(get-date $Today -UFormat %V)
    #HIBA mappa
    $base="C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\HIBA\HIBA $($Today.Year)\"
    $dir = "$base$Week. hét ($($Monday.ToString("MM.dd.")) - $($Sunday.ToString("MM.dd.")))"
    if (-Not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force
    }  
}