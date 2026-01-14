BeforeAll {
    # Check if running with administrator privileges
    $script:IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

Describe "SMB1 Protocol Status Test" {
    
    Context "SMB1 Protocol Client Feature Status" {
        
        It "Should have SMB1Protocol client feature disabled" {
            # Get the Windows Optional Feature status for SMB1Protocol
            $smb1Feature = Get-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol" -ErrorAction SilentlyContinue
            
            # Assert that the feature exists
            $smb1Feature | Should -Not -BeNullOrEmpty -Because "SMB1Protocol feature should exist on Windows"
            
            # Assert that SMB1 is Disabled (not Enabled)
            $smb1Feature.State | Should -Be "Disabled" -Because "SMB1 Protocol is insecure and should be disabled for security reasons"
            
            Write-Host "SMB1 Protocol Client Status: $($smb1Feature.State)" -ForegroundColor Green
            Write-Host "Feature Name: $($smb1Feature.FeatureName)" -ForegroundColor Cyan
        }
    }
}

AfterAll {
    Write-Host "`n=== SMB1 Security Test Summary ===" -ForegroundColor Green
    Write-Host "All SMB1 protocol status test completed" -ForegroundColor Green
}
