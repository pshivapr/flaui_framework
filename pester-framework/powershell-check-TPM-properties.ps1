BeforeAll {
    # No setup required for TPM tests
}

Describe "System Verify TPM Properties Tests" {
    
    Context "Get-Tpm Tests" {

        It "Should have TPM present and enabled" {
            $tpm = Get-Tpm
            $tpm.TpmPresent | Should -Be $true
            $tpm.TpmReady | Should -Be $true
        }

        It "Should have a ManufacturerId and ManufacturerVersion" {
            $tpm = Get-Tpm
            $tpm.ManufacturerId | Should -Not -BeNullOrEmpty
            $tpm.ManufacturerVersion | Should -Not -BeNullOrEmpty
            Write-Host "TPM ManufacturerId: $($tpm.ManufacturerId)" -ForegroundColor Cyan
            Write-Host "TPM ManufacturerVersion: $($tpm.ManufacturerVersion)" -ForegroundColor Cyan
        }

        It "Should have LockoutHealTime and LockoutCount properties" {
            $tpm = Get-Tpm
            $tpm.LockoutHealTime | Should -Be "10 minutes"
            $tpm.LockoutCount | Should -Be "0"
            $tpm.LockoutMax | Should -Be "31"
            Write-Host "TPM LockoutHealTime: $($tpm.LockoutHealTime)" -ForegroundColor Cyan
            Write-Host "TPM LockoutCount: $($tpm.LockoutCount)" -ForegroundColor Cyan
            Write-Host "TPM LockoutMaximum: $($tpm.LockoutMax)" -ForegroundColor Cyan
        }
    }
}

AfterAll {
    Write-Host "`n===System TPM Test Summary ===" -ForegroundColor Green
    Write-Host "All TPM verification tests completed" -ForegroundColor Green
}
