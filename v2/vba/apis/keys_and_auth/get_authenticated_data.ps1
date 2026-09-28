Set-StrictMode -Version 1.0

function Get-AuthenticatedData {

    $url = "https://httpbin.org/bearer"

    $keyPath = "$env:USERPROFILE\demo-api-key.txt"

    if (-not (Test-Path $keyPath)) {
        Write-Host "No key file found at $keyPath. Run the save command first."
        return
    }

    $secureKey = Get-Content $keyPath | ConvertTo-SecureString
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
    $apiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

    $headers = @{ Authorization = "$apiKey" }

    try {
        $response = Invoke-WebRequest -Uri $url -Method Get -Headers $headers -UseBasicParsing
        $status = [int]$response.StatusCode
        $body = $response.Content
    }
    catch {
        if ($_.Exception.Response) {
            $status = [int]$_.Exception.Response.StatusCode
            $body = ""
        }
        else {
            Write-Host "The request did not reach the server: $($_.Exception.Message)"
            return
        }
    }

    Write-Host "HTTP Status: $status"
    Write-Host "Raw Response: $body"

    if ($status -eq 401) {
        Write-Host "401 Unauthorized. Check the key is present, correct, and formatted as Bearer followed by a space."
        return
    }
    elseif ($status -ne 200) {
        Write-Host "The API returned status $status"
        return
    }

}

Get-AuthenticatedData

