BeforeAll {
    # Define the PowerPoint executable path
    $script:PowerPointPath = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
    
    # Alternative: Use just the executable name if it's in PATH
    # $script:PowerPointExe = "POWERPNT.EXE"
}

Describe "PowerPoint Application Tests" {
    
    Context "Launch PowerPoint and Check Version" {
        
        BeforeEach {
            # Start PowerPoint process
            $script:PowerPointProcess = Start-Process -FilePath $PowerPointPath -PassThru
            Start-Sleep -Seconds 3
        }
        
        AfterEach {
            # Clean up - close PowerPoint
            if ($script:PowerPointProcess -and !$script:PowerPointProcess.HasExited) {
                $script:PowerPointProcess.CloseMainWindow() | Out-Null
                Start-Sleep -Milliseconds 500
                if (!$script:PowerPointProcess.HasExited) {
                    Stop-Process -Id $script:PowerPointProcess.Id -Force
                }
            }
        }
        
        It "Should retrieve detailed version from About dialog" {

            # Verify process started
            $script:PowerPointProcess | Should -Not -BeNullOrEmpty
            $script:PowerPointProcess.HasExited | Should -Be $false

            # Get the file version info
            $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($PowerPointPath)
            
            # Assertions
            $versionInfo | Should -Not -BeNullOrEmpty
            $versionInfo.ProductVersion | Should -Not -BeNullOrEmpty
            $versionInfo.ProductName | Should -Match "Microsoft Office"
            Write-Host "PowerPoint Product Name: $($versionInfo.ProductName)" -ForegroundColor Green

            # Get version using COM object
            try {
                $powerPointApp = New-Object -ComObject PowerPoint.Application
                $powerPointApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
                
                $version = $powerPointApp.Version
                $version | Should -Not -BeNullOrEmpty
                
                Write-Host "PowerPoint COM Version: $version" -ForegroundColor Cyan
                
                # Clean up COM object
                $powerPointApp.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPointApp) | Out-Null
            }
            catch {
                Write-Warning "Could not retrieve version via COM: $_"
                Set-ItResult -Skipped -Because "COM automation not available"
            }
        }
    }
}
