# cURL for Beginners — Call APIs from Windows CMD

**Skills and Automation**

This repository contains the commands and scripts demonstrated in the cURL for Beginners video tutorial.

Learn how to communicate with APIs directly from Windows Command Prompt, send and retrieve data, inspect HTTP responses, pass authentication headers, and automate API requests using Windows batch files.

## Prerequisites

- Windows 10 or Windows 11
- Windows Command Prompt (CMD)
- cURL
- Postman (optional, for comparing API requests)

The examples use [Postman Echo](https://postman-echo.com), a public API testing service that returns information about the requests it receives.

No API key or account is required for these examples.

> All commands in this tutorial are written for Windows Command Prompt (CMD). Quoting rules may differ in PowerShell, Bash and other shells.

---

## 1. Check Your cURL Installation

Open Windows Command Prompt and run:

```cmd
curl.exe --version
```

This displays the installed cURL version, supported protocols and available features.

Modern versions of Windows normally include cURL.

We explicitly use `curl.exe` throughout this tutorial to avoid confusion with the `curl` alias found in older versions of Windows PowerShell.

---

## 2. Your First GET Request

Retrieve data from an API using a GET request.

```cmd
curl.exe "https://postman-echo.com/get?report=sales&month=11"
```

The URL contains two query parameters:

| Parameter | Value |
|---|---|
| report | sales |
| month | 11 |

cURL uses the GET method by default when no request method or body is specified for this example.

The response is returned in JSON format.

### Important: Double Quotes

Place the entire URL inside double quotes.

The ampersand (`&`) has a special meaning in Windows CMD. Without double quotes, CMD may interpret part of the URL as a separate command.

### Equivalent Request in Postman

- Method: GET
- URL: `https://postman-echo.com/get`

Under the **Params** tab, enter:

| Key | Value |
|---|---|
| report | sales |
| month | 11 |

Click **Send** to execute the request.

---

## 3. HTTP Status and Response Headers

Use the `-i` option to include HTTP response headers in the output.

```cmd
curl.exe -i "https://postman-echo.com/get?report=sales"
```

The output includes:

- HTTP status
- Response headers
- Response body

### Common HTTP Status Code Categories

| Status Code | Meaning |
|---|---|
| 2xx | Successful response |
| 4xx | Client-side error |
| 5xx | Server-side error |

### Test an HTTP Error

The following endpoint deliberately returns an HTTP 404 response.

```cmd
curl.exe -i "https://postman-echo.com/status/404"
```

By default, receiving an HTTP error response does not necessarily cause cURL to report a command failure.

This is particularly important when automating API requests.

---

## 4. Send JSON with POST

In this example, we will send a JSON request body to an API.

### Step 1: Create a Working Folder

Run the following commands:

```cmd
mkdir "%USERPROFILE%\curl-demo"
cd /d "%USERPROFILE%\curl-demo"
```

### Step 2: Create a JSON File

Open Notepad from CMD:

```cmd
notepad request.json
```

Enter the following JSON:

```json
{
    "report": "Monthly Sales",
    "department": "Finance"
}
```

Save the file as `request.json` and close Notepad.

### Step 3: Send the POST Request

```cmd
curl.exe -X POST -H "Content-Type: application/json" --data-binary "@request.json" "https://postman-echo.com/post"
```

### Understanding the Command

| Option | Description |
|---|---|
| `-X POST` | Explicitly specifies the HTTP POST method |
| `-H` | Adds a request header |
| `Content-Type: application/json` | Indicates that the request body contains JSON |
| `--data-binary` | Sends data without altering its contents |
| `@request.json` | Reads the request body from the specified file |

Postman Echo returns information about the request it received, including the JSON data we supplied.

### Equivalent Request in Postman

- Method: POST
- URL: `https://postman-echo.com/post`
- Body: raw
- Body format: JSON

Enter the same JSON in the request body and click **Send**.

---

## 5. Authentication Headers

Many business APIs require authentication.

One common approach is to send a bearer token using the Authorization header.

```cmd
curl.exe -H "Authorization: Bearer DEMO_ONLY" "https://postman-echo.com/headers"
```

The header consists of:

| Component | Value |
|---|---|
| Header name | Authorization |
| Authentication scheme | Bearer |
| Token | DEMO_ONLY |

The Postman Echo service returns information about the request headers it received.

### Equivalent Request in Postman

- Method: GET
- URL: `https://postman-echo.com/headers`
- Authorization Type: Bearer Token
- Token: `DEMO_ONLY`

> **Security Note:** DEMO_ONLY is a fictional token used for demonstration purposes. This example does not perform actual authentication. Never expose real API tokens in public repositories, recordings or shared batch files.

---

## 6. Generate cURL from Postman

Postman can generate cURL commands from an existing API request.

1. Configure an API request in Postman.
2. Click the **Code** option in the right sidebar.
3. Select **cURL** from the available languages.
4. Review the generated command.
5. Adapt the command for Windows CMD if required.

For our GET request, the simplified CMD-compatible command is:

```cmd
curl.exe "https://postman-echo.com/get?report=sales&month=11"
```

### Things to Check

Generated cURL commands may contain:

- Single quotes instead of double quotes
- Shell-specific line continuation characters
- Unnecessary headers
- Cookies
- Sensitive authentication information

Always review generated commands before copying, executing or sharing them.

---

## 7. Automate an API Request with a Batch File

In this example, we will create a Windows batch file that:

- Calls an API
- Saves the response to a JSON file
- Checks whether the request succeeded
- Reports success or failure

### Step 1: Open the Working Folder

```cmd
cd /d "%USERPROFILE%\curl-demo"
```

### Step 2: Create the Batch File

```cmd
notepad fetch-report.bat
```

### Step 3: Write the Batch Script

File: `fetch-report.bat`

```bat
@echo off

curl.exe --fail --silent --show-error --output sales-response.json "https://postman-echo.com/get?report=sales"

if errorlevel 1 (
    echo API request failed.
    exit /b 1
)

echo API response saved to sales-response.json
```

Save the file and close Notepad.

### Understanding the cURL Options

| Option | Description |
|---|---|
| `--fail` | Reports most HTTP 4xx and 5xx responses as command failures |
| `--silent` | Suppresses the progress display |
| `--show-error` | Displays error messages even when silent mode is enabled |
| `--output` | Saves the response body to a specified file |

### Understanding the Batch File

| Command | Description |
|---|---|
| `@echo off` | Prevents commands from being displayed as they execute |
| `if errorlevel 1` | Checks whether the exit code is greater than or equal to 1 |
| `exit /b 1` | Exits the batch file with an exit code of 1 |
| `echo` | Displays a message in the console |

### Step 4: Execute the Batch File

```cmd
fetch-report.bat
```

Expected success message:

```text
API response saved to sales-response.json
```

### Step 5: Inspect the Output

Check that the file exists:

```cmd
dir sales-response.json
```

Open the saved response in Notepad:

```cmd
notepad sales-response.json
```

The file contains the JSON response returned by the API.

---

## 8. Test the Batch File Error Handling

Open the batch file:

```cmd
notepad fetch-report.bat
```

Temporarily replace the original cURL command with the following:

```bat
curl.exe --fail --silent --show-error --output sales-response.json "https://postman-echo.com/status/404"
```

Save the file and execute it again:

```cmd
fetch-report.bat
```

The API deliberately returns an HTTP 404 response.

cURL reports a failure, and the batch file displays:

```text
API request failed.
```

The batch file terminates with a nonzero exit code.

Once testing is complete, restore the original cURL command:

```bat
curl.exe --fail --silent --show-error --output sales-response.json "https://postman-echo.com/get?report=sales"
```

> A production automation should also validate the response contents and account for authentication, retries and other possible failures. A file left over from a previous successful run does not indicate that the latest request succeeded.

---

## 9. Schedule the Batch File

The batch file can be scheduled to run automatically using Windows Task Scheduler.

For example, it could retrieve data from an API every morning before the start of the business day.

When configuring the scheduled task, ensure the **Start in** folder points to the working directory containing the batch file.

The overall workflow is:

```text
Windows Task Scheduler
        |
        v
fetch-report.bat
        |
        v
     cURL
        |
        v
      API
        |
        v
sales-response.json
```

Refer to the Windows Task Scheduler for Beginners video for a detailed walkthrough of creating scheduled tasks.

---

## 10. Final Folder Structure

After completing the tutorial, your working folder should contain the following files:

```text
curl-demo/
│
├── request.json
├── fetch-report.bat
└── sales-response.json
```

| File | Purpose |
|---|---|
| request.json | JSON request body used in the POST example |
| fetch-report.bat | Batch script that retrieves and saves API data |
| sales-response.json | Response file generated by the batch script |

---

## cURL Quick Reference

| Command or Option | Purpose |
|---|---|
| `curl.exe --version` | Display cURL version and features |
| `curl.exe "URL"` | Send a GET request |
| `-i` | Include response headers |
| `-X POST` | Specify the POST method |
| `-H` | Add a request header |
| `--data-binary "@file.json"` | Send the contents of a file |
| `--output filename` | Save the response to a file |
| `--fail` | Report most HTTP errors as command failures |
| `--silent` | Suppress the progress display |
| `--show-error` | Display error messages |

---

## When Should You Use Each Tool?

| Tool | Typical Use Case |
|---|---|
| Postman | Exploring, testing and troubleshooting APIs through a graphical interface |
| cURL | Quick API requests from a terminal and simple batch file automation |
| PowerShell | Scripting, processing JSON and more complex automation workflows |
| Excel VBA | Integrating APIs with Excel workbooks and business processes |

These are different tools for communicating with the same APIs.

---

## Useful Resources

- [Official cURL Website](https://curl.se/)
- [cURL Command-Line Documentation](https://curl.se/docs/manpage.html)
- [Postman Echo Documentation](https://www.postman.com/postman/postman-echo/overview)
- [Postman](https://www.postman.com/)
- [Skills and Automation](https://skillsandautomation.com/)

---

## Related Tutorials

Part of the Skills and Automation Windows automation series:

- Windows CMD for Beginners
- Windows Batch Files for Beginners
- Windows Task Scheduler for Beginners
- cURL for Beginners — Call APIs from Windows CMD

---

**Skills and Automation**

Practical technology, programming and automation for business professionals.
