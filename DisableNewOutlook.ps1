$user_sids = Get-ChildItem -Path Registry::HKU -ErrorAction silentlycontinue

Write-Host "-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-"

$eligible_count = 0
$user_sids | ForEach-Object {
    $user_sid = $_.Name
    $base_reg_path = "Registry::$user_sid\Software\Microsoft\Office\16.0\Outlook"

    if (Test-Path -Path $base_reg_path) {
        $eligible_count++

        Write-Host "SID:", $user_sid, "Is eligble... Clearing"

        if (Test-Path -Path "$base_reg_path\Preferences") {
            Write-Host "Trying to create: $reg_path\Preferences"
            New-ItemProperty -Path "$base_reg_path\Preferences" -Name "UseNewOutlook" -Value 0 -PropertyType DWORD -ErrorAction SilentlyContinue
        } else {
            Write-Path "Failed to create DWORD in $base_reg_path\Preferences, Path does not exist."
        }

        if (Test-Path -Path "$base_reg_path\Options\General") {
            Write-Host "Trying to create: $reg_path\Options\General"
            New-ItemProperty -Path "$reg_path\Options\General" -Name "HideNewOutlookToggle" -Value 1 -PropertyType DWORD -ErrorAction SilentlyContinue
        } else {
            Write-Host "Failed to create DWORD in $reg_path\Options\General, Path does not exist."
        }
        Write-Host "-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-"
    }
}
Write-Host "Amount of new Outlooks disabled:", $eligible_count