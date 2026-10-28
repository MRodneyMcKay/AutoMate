class NationalityZipStatisticViewModel : WPFToolkit.MVVM.ObservableObject {
    hidden [string] $DefaultOutputDirectory

    [System.Collections.ObjectModel.ObservableCollection[string]] $NationalityFiles
    [System.Collections.ObjectModel.ObservableCollection[string]] $ZipFiles
    [System.Collections.ObjectModel.ObservableCollection[string]] $SelectedNationalityFiles
    [System.Collections.ObjectModel.ObservableCollection[string]] $SelectedZipFiles
    [WPFToolkit.MVVM.ObservableProperty[string]] $OutputDirectory
    [WPFToolkit.MVVM.ObservableProperty[string]] $Status
    [System.Windows.Window] $Window
    [scriptblock] $StatisticsAction

    [WPFToolkit.MVVM.RelayCommand] $BrowseNationalityFilesCommand
    [WPFToolkit.MVVM.RelayCommand] $BrowseZipFilesCommand
    [WPFToolkit.MVVM.RelayCommand] $BrowseOutputDirectoryCommand
    [WPFToolkit.MVVM.RelayCommand] $CalculateStatisticsCommand
    [WPFToolkit.MVVM.RelayCommand] $DeleteSelectedFilesCommand

    NationalityZipStatisticViewModel([scriptblock] $statisticsAction) {
        $this.StatisticsAction = $statisticsAction
        $this.NationalityFiles = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
        $this.ZipFiles = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
        $this.SelectedNationalityFiles = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
        $this.SelectedZipFiles = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
        $this.DefaultOutputDirectory = 'C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\irányítószám statisztika\'
        $this.OutputDirectory = [WPFToolkit.MVVM.ObservableProperty[string]]::new($this.DefaultOutputDirectory)
        $this.Status = [WPFToolkit.MVVM.ObservableProperty[string]]::new("Készen áll.")
        $viewModel = $this

        $this.BrowseNationalityFilesCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                $viewModel.BrowseFiles("Nationality", "Nemzetiségi CSV-fájlok")
            }.GetNewClosure()
        )
        $this.BrowseZipFilesCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                $viewModel.BrowseFiles("Zip", "Irányítószám CSV-fájlok")
            }.GetNewClosure()
        )
        $this.BrowseOutputDirectoryCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                $viewModel.BrowseOutputDirectory()
            }.GetNewClosure()
        )
        $this.CalculateStatisticsCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                $viewModel.CalculateStatistics()
            }.GetNewClosure(),
            [System.Predicate[object]] {
                return $viewModel.HasMatchingPairs()
            }.GetNewClosure()
        )
        $this.DeleteSelectedFilesCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                $viewModel.DeleteSelectedFiles()
            }.GetNewClosure(),
            [System.Predicate[object]] {
                return $viewModel.SelectedNationalityFiles.Count -gt 0 -or
                    $viewModel.SelectedZipFiles.Count -gt 0
            }.GetNewClosure()
        )
    }

    [void] BrowseFiles(
        [string] $FileType,
        [string] $title
    ) {
        $dialog = [Microsoft.Win32.OpenFileDialog]::new()
        $dialog.Title = $title
        $dialog.Filter = "CSV-fájlok (*.csv)|*.csv"
        $dialog.Multiselect = $true

        if ($dialog.ShowDialog() -eq $true) {
            $this.AddCsvFiles($FileType, $dialog.FileNames)
        }
    }

    [void] BrowseOutputDirectory() {
        $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
        $dialog.Description = "Válassza ki a jelentések mentési mappáját:"
        $dialog.ShowNewFolderButton = $true
        $dialog.SelectedPath = $this.OutputDirectory.Value

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $this.OutputDirectory.Value = $dialog.SelectedPath
            [System.Windows.Input.CommandManager]::InvalidateRequerySuggested()
        }
    }

    [void] AddNationalityFiles([string[]] $paths) {
        $this.AddCsvFiles("Nationality", $paths)
    }

    [void] AddZipFiles([string[]] $paths) {
        $this.AddCsvFiles("Zip", $paths)
    }

    hidden [void] AddCsvFiles(
        [string] $FileType,
        [string[]] $paths
    ) {
        if ($FileType -eq "Nationality") {
            $files = $this.NationalityFiles
        } elseif ($FileType -eq "Zip") {
            $files = $this.ZipFiles
        } else {
            throw "Ismeretlen fájltípus: $FileType"
        }
        $requiredNamePart = if ($FileType -eq "Nationality") { "nemzet" } else { "irányító" }
        $rejectedFiles = [System.Collections.Generic.List[string]]::new()
        $addedFiles = [System.Collections.Generic.List[string]]::new()

        foreach ($path in $paths) {
            $fileName = [System.IO.Path]::GetFileName($path)
            if ([System.IO.Path]::GetExtension($path) -ine ".csv" -or
                $fileName.IndexOf($requiredNamePart, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                $rejectedFiles.Add($fileName)
                continue
            }
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                $rejectedFiles.Add("$fileName (nem található)")
                continue
            }
            if ($files.Contains($path)) {
                continue
            }

            $files.Add($path)
            $addedFiles.Add($fileName)
        }

        if ($rejectedFiles.Count -gt 0) {
            $this.Status.Value = "Hiba: nem megfelelő fájl (CSV és '$requiredNamePart' szükséges): $($rejectedFiles -join ', ')"
        } elseif ($addedFiles.Count -eq 1) {
            $this.Status.Value = "$($addedFiles[0]) hozzáadva."
        } elseif ($addedFiles.Count -gt 1) {
            $this.Status.Value = "$($addedFiles.Count) fájl hozzáadva."
        }

        [System.Windows.Input.CommandManager]::InvalidateRequerySuggested()
    }

    [string] GetFilePrefix([string] $path, [string] $fileType) {
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($path)
        $marker = if ($fileType -eq "Nationality") { "nemzet" } else { "irányító" }
        $markerIndex = $fileName.IndexOf($marker, [System.StringComparison]::OrdinalIgnoreCase)
        if ($markerIndex -lt 0) {
            return ""
        }

        return $fileName.Substring(0, $markerIndex).TrimEnd([char[]]@(' ', '-', '_', '.'))
    }

    [string] GetPairedFile([string] $path, [string] $fileType) {
        $sourcePrefix = $this.GetFilePrefix($path, $fileType)
        if ([string]::IsNullOrWhiteSpace($sourcePrefix)) {
            return ""
        }

        $targetFiles = if ($fileType -eq "Nationality") {
            $this.ZipFiles
        } elseif ($fileType -eq "Zip") {
            $this.NationalityFiles
        } else {
            throw "Ismeretlen fájltípus: $fileType"
        }
        $targetType = if ($fileType -eq "Nationality") { "Zip" } else { "Nationality" }
        foreach ($targetPath in $targetFiles) {
            if ($this.GetFilePrefix($targetPath, $targetType) -ceq $sourcePrefix) {
                return $targetPath
            }
        }

        return ""
    }

    [void] UpdateSelection([string] $fileType, [string[]] $selectedFiles) {
        if ($fileType -eq "Nationality") {
            $target = $this.SelectedNationalityFiles
        } elseif ($fileType -eq "Zip") {
            $target = $this.SelectedZipFiles
        } else {
            throw "Ismeretlen fájltípus: $fileType"
        }

        $target.Clear()
        foreach ($path in $selectedFiles) {
            if (-not $target.Contains($path)) {
                $target.Add($path)
            }
        }
        [System.Windows.Input.CommandManager]::InvalidateRequerySuggested()
    }

    [void] ClearSelection([string] $fileType) {
        $this.UpdateSelection($fileType, [string[]]@())
    }

    [void] DeleteSelectedFiles() {
        $nationalityFilesToDelete = [string[]]@($this.SelectedNationalityFiles)
        $zipFilesToDelete = [string[]]@($this.SelectedZipFiles)
        $deletedFiles = [System.Collections.Generic.List[string]]::new()
        foreach ($path in $nationalityFilesToDelete) {
            if ($this.NationalityFiles.Remove($path)) {
                $deletedFiles.Add([System.IO.Path]::GetFileName($path))
            }
        }
        foreach ($path in $zipFilesToDelete) {
            if ($this.ZipFiles.Remove($path)) {
                $deletedFiles.Add([System.IO.Path]::GetFileName($path))
            }
        }
        $this.SelectedNationalityFiles.Clear()
        $this.SelectedZipFiles.Clear()
        if ($deletedFiles.Count -eq 1) {
            $this.Status.Value = "$($deletedFiles[0]) törölve."
        } elseif ($deletedFiles.Count -gt 1) {
            $this.Status.Value = "$($deletedFiles.Count) fájl törölve."
        }
        [System.Windows.Input.CommandManager]::InvalidateRequerySuggested()
    }

    [bool] HasMatchingPairs() {
        if ($this.NationalityFiles.Count -eq 0 -or
            $this.NationalityFiles.Count -ne $this.ZipFiles.Count -or
            [string]::IsNullOrWhiteSpace($this.OutputDirectory.Value)) {
            return $false
        }

        $nationalityPrefixes = [System.Collections.Generic.List[string]]::new()
        foreach ($path in $this.NationalityFiles) {
            $prefix = $this.GetFilePrefix($path, "Nationality")
            if ([string]::IsNullOrWhiteSpace($prefix) -or
                $nationalityPrefixes -contains $prefix) {
                return $false
            }
            $nationalityPrefixes.Add($prefix)
        }

        $zipPrefixes = [System.Collections.Generic.List[string]]::new()
        foreach ($path in $this.ZipFiles) {
            $prefix = $this.GetFilePrefix($path, "Zip")
            if ([string]::IsNullOrWhiteSpace($prefix) -or
                $zipPrefixes -contains $prefix) {
                return $false
            }
            $zipPrefixes.Add($prefix)
        }

        foreach ($prefix in $nationalityPrefixes) {
            if ($zipPrefixes -notcontains $prefix) {
                return $false
            }
        }

        return $true
    }

    [void] RefreshStatus() {
        if ($null -ne $this.Window) {
            $targetWindow = $this.Window
            $targetWindow.Dispatcher.Invoke(
                [System.Action] { $targetWindow.UpdateLayout() }.GetNewClosure(),
                [System.Windows.Threading.DispatcherPriority]::Render
            )
        }
    }

    [void] CalculateStatistics() {
        try {
            $nationalityPaths = [string[]]@($this.NationalityFiles)
            $zipPaths = [string[]]@($this.ZipFiles)
            $viewModel = $this
            $onProgress = {
                param($prefix, $index, $total)
                $viewModel.Status.Value = "Számítás: {0} ({1}/{2})" -f $prefix, $index, $total
                $viewModel.RefreshStatus()
            }.GetNewClosure()
            $actionOutput = @(& $this.StatisticsAction `
                $nationalityPaths `
                $zipPaths `
                $this.OutputDirectory.Value `
                $onProgress)
            $results = @(
                $actionOutput | Where-Object {
                    $null -ne $_ -and
                    $null -ne $_.PSObject.Properties['PairCount'] -and
                    $null -ne $_.PSObject.Properties['Prefixes']
                }
            )
            if ($results.Count -eq 0) {
                throw "A statisztikai művelet nem adott vissza feldolgozási eredményt."
            }

            $result = $results[-1]
            $prefixes = [string]::Join(", ", [string[]]$result.Prefixes)
            $this.Status.Value = "Elkészült {0} pár: {1}" -f [int]$result.PairCount, $prefixes
        }
        catch {
            $this.Status.Value = "Hiba: $($_.Exception.Message)"
        }
        $this.RefreshStatus()
    }
}
