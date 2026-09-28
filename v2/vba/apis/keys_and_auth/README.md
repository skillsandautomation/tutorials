API Keys & Authentication in Excel VBA
This project demonstrates how to call an API that requires authentication from both Excel VBA and PowerShell.
The VBA example uses WinHttp.WinHttpRequest.5.1 to call the httpbin Bearer authentication endpoint. It first demonstrates how a request fails with 401 Unauthorized when no valid Bearer header is supplied, then adds the Authorization header in the correct format.
The API key is read from a Windows environment variable instead of being hardcoded into the workbook. This keeps the key outside the Excel file, which is safer when the workbook is shared.
The PowerShell example covers the same request from a different angle. The API key is stored in an encrypted file using Windows DPAPI, then decrypted at runtime and passed in the Bearer Authorization header.
Main concepts covered:
- API keys and authentication
- Bearer tokens
- Authorization headers
- 401 Unauthorized
- 401 vs 403
- Windows environment variables
- Reading environment variables from VBA
- Avoiding hardcoded secrets
- PowerShell SecureString
- Windows DPAPI encryption
- Invoke-WebRequest
- Basic API error handling
The test endpoint used in the video is:
https://httpbin.org/bearer
httpbin accepts any Bearer token, making it useful for learning and testing request structure without needing a real API key.
Excel VBA Code
Option Explicit

Sub Get_Authenticated_Data()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Auth")

    Dim apiKey As String
    apiKey = Environ("DEMO_API_KEY")

    If Len(apiKey) = 0 Then
        MsgBox "No API key found. Add DEMO_API_KEY to your environment variables, " & _
               "then restart Excel."
        Exit Sub
    End If

    Dim url As String
    url = "https://httpbin.org/bearer"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.Send

    ws.Range("B5").Value = http.Status
    ws.Range("B6").Value = http.responseText

    If http.Status = 401 Then
        MsgBox "401 Unauthorized. Check the key is present, correct, and formatted as " & _
               "Bearer followed by a space."
        Exit Sub

    ElseIf http.Status <> 200 Then
        MsgBox "The API returned status " & http.Status & vbNewLine & http.responseText
        Exit Sub
    End If

End Sub

PowerShell: Save the API Key
Run this once in Windows PowerShell ISE:
Read-Host "Enter your API key" -AsSecureString |
    ConvertFrom-SecureString |
    Set-Content "$env:USERPROFILE\demo-api-key.txt"

PowerShell: Read the Key Back
$secure = Get-Content "$env:USERPROFILE\demo-api-key.txt" | ConvertTo-SecureString

[System.Runtime.InteropServices.Marshal]::PtrToStringBSTR(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
)

PowerShell: Complete API Request
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

    $headers = @{
        Authorization = "Bearer $apiKey"
    }

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
