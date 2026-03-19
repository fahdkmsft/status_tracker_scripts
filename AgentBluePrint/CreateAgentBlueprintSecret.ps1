param(
[string]$AgentObjectId,
[string]$DisplayName = "My Secret",
[string]$EndDateTime = "2060-08-05T23:59:59Z"
)

# Prompt if not provided

if (-not $AgentObjectId) {
    $AgentObjectId = Read-Host "Enter Agent Blueprint (Application) Object ID"
}

if (-not $AgentObjectId) {
    Write-Error "AgentObjectId is required."
    exit 1
}

Write-Host "Logging in via Azure CLI..."
az login | Out-Null

Write-Host "Obtaining Microsoft Graph access token..."
$token = az account get-access-token --resource-type ms-graph --query accessToken -o tsv

if (-not $token) {
    Write-Error "Failed to obtain access token."
    exit 1
}

Write-Host "Creating client secret for application (Agent Blueprint)..."

$body = @{
    passwordCredential = @{
        displayName = $DisplayName
        endDateTime = $EndDateTime
    }
} | ConvertTo-Json -Depth 5

try {
    $params = @{
        Method  = "POST"
        Uri     = "https://graph.microsoft.com/beta/applications/$AgentObjectId/addPassword"
        Headers = @{
            Authorization = "Bearer $token"
            "Content-Type" = "application/json"
        }
        Body    = $body
    }

    $response = Invoke-RestMethod @params

    Write-Host "`nClient secret created successfully!" -ForegroundColor Green

    Write-Host "`nIMPORTANT: Save this secret now. It will not be shown again.`n" -ForegroundColor Yellow
    Write-Host "Client Secret:"
    Write-Output $response.secretText
}
catch {
    Write-Error "Request failed:"
    Write-Error $_
}
