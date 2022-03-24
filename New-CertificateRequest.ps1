param (
    # option to provide $upn at command line
    [string]$upn
)

# warning message 
$prompt = @"


*************************************************************************
* THIS SCRIPT IS INTENDED FOR USE ON OPTIMIZELY OWNED COMPUTERS ONLY.   *
* EXECUTING THIS SCRIPT ON A NON-OPTIMIZELY OWNED DEVICE IS A VIOLATION *
* OF THE OPTIMIZELY CUSTOMER SECURITY POLICY.                           *
* ***********************************************************************

Please type "OPTIMIZELY" to confirm you understand: 
"@

# prompt to confirm understanding of warning message
if ($(Read-Host -Prompt $prompt) -ne 'OPTIMIZELY') {
    Break
}
else {

    # prompt user for Okta username
    if (-not $upn) {
        $upn = Read-Host 'Enter your Okta username. For example: firstName.lastName@optimizely.com'
    }

    # base template for cert request
    $template = @"
[Version]
Signature = $Windows NT$

[NewRequest]
Subject    = "CN=$upn,DC=ad,DC=optimizely,DC=net"
KeySpec    = 1
KeyLength    = 2048
Exportable = False
ExportableEncrypted = False
MachineKeySet = False
ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
RequestType    = PKCS10
KeyUsage =    "CERT_DIGITAL_SIGNATURE_KEY_USAGE | CERT_KEY_ENCIPHERMENT_KEY_USAGE"

[EnhancedKeyUsageExtension]
OID = 1.3.6.1.5.5.7.3.2

[Extensions]
2.5.29.17  = "{text}"
_continue_ = "EMail=$upn&"
"@

    # save the updated request template
    $template | Out-File "$env:temp\request.ini"
    # run certreq to generate the base64 encoded request
    & certreq.exe -new "$env:temp\request.ini" "$env:temp\request.txt"

    # check for error generating certificate
    if (-not $?) {
        Write-Verbose 'Error generating cert, please double check your input.' -Verbose
        Break
    }
    else {
        # get the encoded request
        $req = Get-Content "$env:temp\request.txt"
        
        # clean up temp files
        Remove-Item "$env:temp\request.ini" -Force:$true
        Remove-Item "$env:temp\request.txt" -Force:$true
        
        # copy encoded request to clipboard
        Set-Clipboard -Value $req

        # dump encoded request to console
        $req

        # notify user
        Write-Verbose 'Your request has been generated and copied to the clipboard.' -Verbose
    }
}

# SIG # Begin signature block
# MIIJcQYJKoZIhvcNAQcCoIIJYjCCCV4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU1VQ1hij9dN5DhipoQWcNcmZ/
# IZOgggbfMIIG2zCCBcOgAwIBAgITXgAAAgVUlTZhfue45wACAAACBTANBgkqhkiG
# 9w0BAQsFADBEMRIwEAYKCZImiZPyLGQBGRYCc2UxEjAQBgoJkiaJk/IsZAEZFgJl
# cDEaMBgGA1UEAxMRRVBpU2VydmVyLWVwaWNhMDEwHhcNMjIwMzI0MTcyMTI5WhcN
# MjMwMzI0MTcyMTI5WjCBjTESMBAGCgmSJomT8ixkARkWAnNlMRIwEAYKCZImiZPy
# LGQBGRYCZXAxEjAQBgNVBAsTCUVQaVNlcnZlcjEWMBQGA1UECxMNTWFuYWdlZCBV
# c2VyczEfMB0GA1UECxMWSW5mb3JtYXRpb24gVGVjaG5vbG9neTEWMBQGA1UEAxMN
# VGltb3RoeSBCcm9jazCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMHa
# 1U92Wsl3UmHSEKzx/kOX3qfwFPzZ/T4a9VesRhT6f/nJ/DFqx7DN/IOzwf7pcgr4
# zECS1n6yRzt5+csYxY9S3fA6cowvJTA0c9dzLXMKOPSwR951DCqQH6yzOCM90nto
# JgeyPy7u2jWlMAdU/oPMzOMDLsZiJEWl83rZumCk+ciYj+6ZPWuQkS/gOHY28iLU
# lUSWCUqEMRhadf6TkWkeoBraOHvBtBfOfrMgbPD3YOwbsjCBZenkXjv+oShG8E24
# 4GP0U40Uf8rpXmPzxdSkNAuC5oh9bFMfKrY8GaILZX2bq+qCMYPyu3k2iFefYqqG
# AW3cr4G44rcRdw/2k40CAwEAAaOCA3owggN2MD0GCSsGAQQBgjcVBwQwMC4GJisG
# AQQBgjcVCIeRzDOG05x2hp2dCIGQuxKH2NZNbIPrszCCsPMuAgFkAgELMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIHgDAbBgkrBgEEAYI3FQoEDjAM
# MAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQ9JOFh8ZepkdrjhE8hHhhlDPlktjBVBgNV
# HREETjBMgRx0aW1vdGh5LmJyb2NrQG9wdGltaXplbHkuY29toCwGCisGAQQBgjcU
# AgOgHgwcdGltb3RoeS5icm9ja0BvcHRpbWl6ZWx5LmNvbTAfBgNVHSMEGDAWgBTE
# //EWZJIHUdc9b+QzsASFE3Ol+TCCAT8GA1UdHwSCATYwggEyMIIBLqCCASqgggEm
# hoG0bGRhcDovLy9DTj1FUGlTZXJ2ZXItZXBpY2EwMSxDTj1lcGljYXJvb3QsQ049
# Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNv
# bmZpZ3VyYXRpb24sREM9ZXAsREM9c2U/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlz
# dD9iYXNlP29iamVjdENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50hjdodHRwOi8v
# ZXBpY2Fyb290LmVwLnNlL0NlcnRFbnJvbGwvRVBpU2VydmVyLWVwaWNhMDEuY3Js
# hjRodHRwOi8vbWFpbC5lcGlzZXJ2ZXIuY29tL2NybGQvRVBpU2VydmVyLWVwaWNh
# MDEuY3JsMIIBFwYIKwYBBQUHAQEEggEJMIIBBTCBqgYIKwYBBQUHMAKGgZ1sZGFw
# Oi8vL0NOPUVQaVNlcnZlci1lcGljYTAxLENOPUFJQSxDTj1QdWJsaWMlMjBLZXkl
# MjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPWVwLERD
# PXNlP2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9u
# QXV0aG9yaXR5MFYGCCsGAQUFBzAChkpodHRwOi8vZXBpY2Fyb290LmVwLnNlL0Nl
# cnRFbnJvbGwvZXBpY2Fyb290LmVwLnNlX0VQaVNlcnZlci1lcGljYTAxKDIpLmNy
# dDANBgkqhkiG9w0BAQsFAAOCAQEAjnw+8xJ+D6UJfFdEha0MxcsyYOqH2R5BRRnU
# f9NKtf3IxucajuWeic7gqcG6zCJO5WnvVEoJw2Pc22IvJLDaxbILNlIbrjDqOUaR
# Wm3WuSHH+38Jpnk0RiTebsYVJM3J4H0Hj/WvgLcqQRncgwCUV0qQIsUd2gkT70TW
# ZmOsulYY8u6LyFPO2WgpquxaK0TGAP6Lkbj+7gdcfwq2kG21g9f6gqaJf4UoIZHi
# Lu0s7ByFjCW7OwFtcNvmRHGb0/DPyX01QoD0qrCB3oZTnPVNh9Hno0zY8PnXC8oT
# 2eg+VTWXFkOEtJXOYYUjZLbEXGH9kwk/NzZFqYqJSwZOGQZ2EDGCAfwwggH4AgEB
# MFswRDESMBAGCgmSJomT8ixkARkWAnNlMRIwEAYKCZImiZPyLGQBGRYCZXAxGjAY
# BgNVBAMTEUVQaVNlcnZlci1lcGljYTAxAhNeAAACBVSVNmF+57jnAAIAAAIFMAkG
# BSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJ
# AzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMG
# CSqGSIb3DQEJBDEWBBRa9Y1opYwfXLFAIMHJtNCjv0B7pzANBgkqhkiG9w0BAQEF
# AASCAQAgODyKNus9nyVYBAxfWYNf2qPJSkIuljDz40RPO3k69cMsKggh3mJuIrtV
# JpRW/hOv8yhPFW620HnMSrjwulNFOZzlw0Qs93D2iOj1SZ/n31lXO0q7mAx8mayp
# XsA1O9pgKwSZsxApFsS21wrlsF4DAoCpElz68TSk02NntL07lhsGwkpTGRz0iukJ
# uj7tuKBcZsUFT71fXeDnjyPfHTQ4qpvLhBqnl5iTSkj3Xjs0FK0tNFt0Qgt4o5k7
# TOLmfHE+mAC4Qcyk7vWf5mNFQZAwr+u1eF4xRzibHyxy7QBk8t7mv85vpVQDBukN
# XvlXCTIUbYGm64xjorKkVp0KVR4c
# SIG # End signature block
