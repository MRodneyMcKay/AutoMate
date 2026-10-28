class NationalityZipStatisticView {
    [System.Windows.Window] $Window
    [NationalityZipStatisticViewModel] $ViewModel

    NationalityZipStatisticView([NationalityZipStatisticViewModel] $viewModel) {
        $this.ViewModel = $viewModel
        $xamlPath = Join-Path $PSScriptRoot "NationalityZipStatisticView.xaml"
        $xaml = Get-Content -LiteralPath $xamlPath -Raw
        $this.Window = [System.Windows.Markup.XamlReader]::Parse($xaml)
        $this.Window.DataContext = $this.ViewModel
        $this.ViewModel.Window = $this.Window
        $this.SetupDragDrop()
        $this.SetupFileSelection()
    }

    hidden [void] SetupDragDrop() {
        $view = $this
        $dropTargets = @(
            @{ Name = "NationalityPane"; FileType = "Nationality" },
            @{ Name = "ZipPane"; FileType = "Zip" }
        )

        foreach ($dropTarget in $dropTargets) {
            $dropArea = $this.Window.FindName($dropTarget.Name)
            $fileType = $dropTarget.FileType
            $normalBackground = $dropArea.Background

            $dropArea.Add_DragOver({
                    param($sender, $e)
                    $hasFiles = $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)
                    $e.Effects = if ($hasFiles) {
                        [System.Windows.DragDropEffects]::Copy
                    } else {
                        [System.Windows.DragDropEffects]::None
                    }
                    $e.Handled = $true
                }.GetNewClosure())

            $dropArea.Add_DragEnter({
                    param($sender, $e)
                    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
                        $sender.Background = $view.Window.FindResource("DropHighlightBrush")
                    }
                    $e.Handled = $true
                }.GetNewClosure())

            $dropArea.Add_DragLeave({
                    param($sender, $e)
                    $sender.Background = $normalBackground
                }.GetNewClosure())

            $dropArea.Add_Drop({
                    param($sender, $e)
                    $sender.Background = $normalBackground
                    if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
                        return
                    }

                    $files = [string[]]$e.Data.GetData([System.Windows.DataFormats]::FileDrop)
                    if ($fileType -eq "Nationality") {
                        $view.ViewModel.AddNationalityFiles($files)
                    } else {
                        $view.ViewModel.AddZipFiles($files)
                    }
                    $e.Handled = $true
                }.GetNewClosure())
        }
    }

    hidden [void] SetupFileSelection() {
        $view = $this
        $nationalityList = $this.Window.FindName("NationalityFileList")
        $zipList = $this.Window.FindName("ZipFileList")
        $selectionSync = @{ IsSyncing = $false }

        [WPFToolkit.Behaviors.BehaviorBridge]::SetAction(
            $nationalityList,
            { $nationalityList.UnselectAll() }.GetNewClosure()
        )
        [WPFToolkit.Behaviors.BehaviorBridge]::SetAction(
            $zipList,
            { $zipList.UnselectAll() }.GetNewClosure()
        )

        $selectionChanged = {
            param($sender, $e)
            if ($selectionSync.IsSyncing) {
                return
            }

            $selectionSync.IsSyncing = $true
            try {
                $sourceType = if ([object]::ReferenceEquals($sender, $nationalityList)) {
                    "Nationality"
                } else {
                    "Zip"
                }
                $targetList = if ($sourceType -eq "Nationality") { $zipList } else { $nationalityList }

                foreach ($path in $e.AddedItems) {
                    $pairedPath = $view.ViewModel.GetPairedFile([string]$path, $sourceType)
                    if (-not [string]::IsNullOrWhiteSpace($pairedPath) -and
                        -not $targetList.SelectedItems.Contains($pairedPath)) {
                        [void]$targetList.SelectedItems.Add($pairedPath)
                    }
                }

                foreach ($path in $e.RemovedItems) {
                    $pairedPath = $view.ViewModel.GetPairedFile([string]$path, $sourceType)
                    if (-not [string]::IsNullOrWhiteSpace($pairedPath) -and
                        $targetList.SelectedItems.Contains($pairedPath)) {
                        [void]$targetList.SelectedItems.Remove($pairedPath)
                    }
                }
            }
            finally {
                $selectionSync.IsSyncing = $false
            }

            $view.ViewModel.UpdateSelection("Nationality", [string[]]@($nationalityList.SelectedItems))
            $view.ViewModel.UpdateSelection("Zip", [string[]]@($zipList.SelectedItems))
        }.GetNewClosure()

        $nationalityList.Add_SelectionChanged($selectionChanged)
        $zipList.Add_SelectionChanged($selectionChanged)

        $nationalityPane = $this.Window.FindName("NationalityPane")
        $zipPane = $this.Window.FindName("ZipPane")
        $this.Window.Add_PreviewMouseDown({
                param($sender, $e)
                $source = $e.OriginalSource
                $activePane = $null
                while ($null -ne $source) {
                    if ([object]::ReferenceEquals($source, $nationalityPane)) {
                        $activePane = $nationalityPane
                        break
                    }
                    if ([object]::ReferenceEquals($source, $zipPane)) {
                        $activePane = $zipPane
                        break
                    }
                    if ($source -is [System.Windows.Media.Visual] -or
                        $source -is [System.Windows.Media.Media3D.Visual3D]) {
                        $source = [System.Windows.Media.VisualTreeHelper]::GetParent($source)
                    } else {
                        break
                    }
                }

                $view.SetActivePane($activePane)
            }.GetNewClosure())

        $this.Window.Add_PreviewKeyDown({
                param($sender, $e)
                if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
                    if ($view.ViewModel.SelectedNationalityFiles.Count -gt 0 -or
                        $view.ViewModel.SelectedZipFiles.Count -gt 0) {
                        $view.ViewModel.DeleteSelectedFiles()
                        $e.Handled = $true
                    }
                } elseif ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                    $nationalityList.UnselectAll()
                    $zipList.UnselectAll()
                    $e.Handled = $true
                }
            }.GetNewClosure())
    }

    hidden [void] SetActivePane([System.Windows.Controls.Border] $activePane) {
        $nationalityPane = $this.Window.FindName("NationalityPane")
        $zipPane = $this.Window.FindName("ZipPane")
        $normalBrush = $this.Window.FindResource("BorderBrush")
        $activeBrush = $this.Window.FindResource("AccentActiveBrush")

        foreach ($pane in @($nationalityPane, $zipPane)) {
            $pane.BorderBrush = $normalBrush
            $pane.BorderThickness = [System.Windows.Thickness]::new(1)
        }

        if ($null -ne $activePane) {
            $activePane.BorderBrush = $activeBrush
            $activePane.BorderThickness = [System.Windows.Thickness]::new(2)
        }
    }

    [void] Show() {
        $this.Window.Show()
    }
}
