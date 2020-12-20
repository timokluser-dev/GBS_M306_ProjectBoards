## Microsoft Azure AD User
$m365User = "<M365 Admin User>"
$m365Password = "<M365 Admin Password>"
$securePassword = $m365Password | ConvertTo-SecureString -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $m365User, $securePassword

$azureAd = Connect-AzureAD -Credential $credentials

## Microsoft Graph Application
$TenantId = "<M365 Tentant Id>"
$ClientId = "<M365 Client Id>"
$ClientSecret = "<M365 Client Secret>"

$AuthResponse = Get-MsalToken -DeviceCode -ClientId $ClientId -TenantId $TenantId -RedirectUri "http://localhost"
