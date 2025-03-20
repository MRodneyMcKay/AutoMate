function create-Directories {
    param (
        [datetime]$Today
    )
    if ($Today.DayOfWeek.value__ -eq 0)
    {
        $Monday = $Today.AddDays(-6 - $Today.DayOfWeek.value__)
        $Sunday = $Monday.AddDays(6)
        $Week = $(get-date $Today -UFormat %V)
    }
    else {
        $Monday = $Today.AddDays(1 - $Today.DayOfWeek.value__)
        $Sunday = $Monday.AddDays(6)
        $Week = $(get-date $Today -UFormat %V)
    }
    #HIBA mappa
    $base="C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\HIBA\HIBA $($Today.Year)\"
    $dir = "$base$Week. hét ($($Monday.ToString("MM.dd.")) - $($Sunday.ToString("MM.dd.")))"
    if (-Not ([System.IO.Directory]::Exists($dir)))
    {
        mkdir -p $dir
    }   
}