# prompt user for Okta username
$upn = Read-Host 'Enter your Okta username. For example: timothy.brock@optimizely.com'

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

# dump encoded request to the console
Get-Content "$env:temp\request.txt"

Remove-Item "$env:temp\request.ini" -Force:$true
Remove-Item "$env:temp\request.txt" -Force:$true
