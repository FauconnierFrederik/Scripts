Install-Module MSOnline

Connect-MsolService




$Results = @()
$CSV = "C:\temp\mfaregistrationdetails.csv"

$AllUsers = Get-MsolUser -All -ErrorAction SilentlyContinue


ForEach ($User in $AllUsers) {
    Write-Host User Details for $User.UserPrincipalName
    #Write-Host "Self-Service Password Feature (SSP)..: " -NoNewline;
    $UserPrincipalName = $User.UserPrincipalName
    $DisplayName = $User.DisplayName
    If ($User.StrongAuthenticationUserDetails) {  
        $SSPR = "Enabled"
    }
    Else {
        $SSPR = "Not Configured"
    }

    #Write-Host "MFA Feature (Portal) ................: " -NoNewline;
    If ((($User | Select-Object -ExpandProperty StrongAuthenticationRequirements).State) -ne $null) {
         $PerUserMFA = "Enabled! This overrides Conditional Access"
        }
    Else { 
        $PerUserMFA = "No"
    }

    #-Host "MFA Feature (Conditional)............: " -NoNewline;
    If ($User.StrongAuthenticationMethods) {
        $MFACapable = "Yes"
    #Write-host "Authentication Methods:"
    ForEach ($Type in $User.StrongAuthenticationMethods) {
        If ($Type.IsDefault -eq $True) {
                $Method = $Type.MethodType
            }
        }

    $PhoneNumber = $User.StrongAuthenticationUserDetails.PhoneNumber
    $AlternativePhoneNumber = $User.StrongAuthenticationUserDetails.AlternativePhoneNumber
    $Email = $User.StrongAuthenticationUserDetails.Email
    }
    Else {
        $Method = "Not Configured"
        $MFACapable = "Not Configured"
        $PhoneNumber = "Not Configured"
        $AlternativePhoneNumber = "Not Configured"
        $Email = "Not Configured"
    }

    $MFAs = @{
        DisplayName = $DisplayName
        UserPrincipalName = $UserPrincipalName
        ObjectID = $User.ObjectID
        SSPR = $SSPR
        PerUserMFA = $PerUserMFA
        Method = $Method
        MFACapable = $MFACapable
        PhoneNumber = $PhoneNumber
        AlternativePhoneNumber = $AlternativePhoneNumber
        AlternativeEmail = $Email
        Licensed = $User.IsLicensed
        BlockCredential = $User.BlockCredential

    }

    $Results += New-Object psobject -Property $MFAs
}

$Results = $Results | Sort-Object DisplayName

$Results | Select DisplayName, UserPrincipalName, ObjectID, Licensed, BlockCredential, PerUserMFA, MFACapable, Method, PhoneNumber, AlternativePhoneNumber, AlternativeEmail, SSPR | Export-CSV $CSV -Encoding UTF8 -NoTypeInformation
Write-Host "CSV output to $CSV" -ForegroundColor Green
