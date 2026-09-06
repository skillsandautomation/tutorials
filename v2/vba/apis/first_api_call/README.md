Your First API Call in Excel: VBA and Power Query

Pull a live exchange rate from a public API into a worksheet cell, twice — once with VBA and once with Power Query. No add-ins, no API key, nothing to install.

The API

Frankfurter — free, open source, central bank data. No API key, no signup, no daily or monthly quota, free for commercial use. Requests are rate-limited.

Endpoint

https://api.frankfurter.dev/v2/rate/EUR/USD

Response

json
{"date":"2026-08-26","base":"EUR","quote":"USD","rate":1.1675}
Requirements

Excel on Windows. The VBA uses WinHTTP, which ships with Windows and is created here with late binding — so there is no reference to tick under Tools > References, and the workbook runs on any Windows machine you send it to.

Save as .xlsm, not .xlsx, or the macro will not survive the save.

Postman is optional. If your work machine will not allow it, Bruno is a good offline alternative with no account needed, and Hoppscotch runs in a browser.

Part 1 — VBA
Sheet setup

Sheet named Rates:

Cell	Column A	Column B
Row 2	Base Currency	EUR
Row 3	Quote Currency	USD
Row 5	Rate	filled by the macro
Row 6	As At	homework
Row 8	HTTP Status	filled by the macro
Row 9	Raw Response	filled by the macro

Alt + F11 to open the editor, then Insert > Module, and paste. Run with F5, or wire it to a button on the sheet.

The code
vba
Option Explicit

Sub Get_Exchange_Rate()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Rates")

    ' Read the currency codes from the sheet - nothing hard-coded
    Dim baseCode As String, quoteCode As String
    baseCode = ws.Range("B2").Value
    quoteCode = ws.Range("B3").Value

    ' Build the endpoint. This is our question, written as text
    Dim url As String
    url = "https://api.frankfurter.dev/v2/rate/" & baseCode & "/" & quoteCode

    ' The messenger. Late binding - no references to set
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' GET   = please give me something
    ' False = synchronous, so VBA waits here for the answer
    http.Open "GET", url, False
    http.Send

    Dim responseText As String
    responseText = http.responseText

    ' Show the raw status and response so we can see what came back
    ws.Range("B8").Value = http.Status
    ws.Range("B9").Value = responseText

    ' Never assume the response is usable. Check the status first.
    ' Scoped to this request - elsewhere 201 and 204 also mean success,
    ' so there you would check the whole 200 range, not a single number.
    If http.Status <> 200 Then
        MsgBox "The API returned status " & http.Status & vbNewLine & responseText
        Exit Sub
    End If

    ' Pull the rate out of the JSON string.
    ' Fragile by design - a proper JSON parser is a later video.
    '
    ' Val, not CDbl. JSON always writes its numbers with a full stop as the
    ' decimal point, and Val always reads a full stop as the decimal point.
    ' CDbl follows Windows regional settings, so on a machine set up to use
    ' a comma the same number can come back as something else entirely.
    Dim startPos As Long, endPos As Long
    startPos = InStr(responseText, """rate"":") + 7
    endPos = InStr(startPos, responseText, "}")

    ws.Range("B5").Value = Val(Mid(responseText, startPos, endPos - startPos))

End Sub
The whole thing in four lines

Everything above is checking, parsing and tidying up. This is the shape underneath it, and it does not change — a weather API, a courier's tracking API, your company's internal system, the OpenAI API.

vba
Sub Minimal_Api_Call()

    Dim url As String
    url = "https://api.frankfurter.dev/v2/rate/EUR/USD"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False
    http.Send

    Debug.Print http.Status
    Debug.Print http.responseText

End Sub
Part 2 — Power Query
Sheet setup

Sheet named Rates PQ:

Cell	Column A	Column B
Row 2	Base Currency	EUR — named BaseCode
Row 3	Quote Currency	USD — named QuoteCode
Row 7	—	query output loads here

To name a cell: select it, type the name into the Name Box to the left of the formula bar, press Enter. Power Query cannot address a bare cell, so this step is not optional.

The simple version

This is what Data > Get Data > From Other Sources > From Web generates for you, after Anonymous > Connect > Into Table > Close & Load.

Worth reading rather than skipping. It built a URL, sent a GET, and parsed the JSON — exactly what the VBA does, with every step hidden from view. The same request, just a different messenger.

powerquery
let
    Source = Json.Document(
        Web.Contents("https://api.frankfurter.dev/v2/rate/EUR/USD")
    ),
    #"Converted to Table" = Record.ToTable(Source)
in
    #"Converted to Table"

The catch is the URL. It is fixed, and nothing on the sheet can change it.

The version driven from the sheet

Data > Get Data > From Other Sources > Blank Query, then Home > Advanced Editor, and paste. Then Close & Load To > Existing worksheet > =$A$7.

powerquery
let
    // Excel.CurrentWorkbook hands us everything Power Query can see inside
    // this file. A named single cell still arrives as a little one-row
    // table, hence {0}[Column1].
    Base  = Excel.CurrentWorkbook(){[Name="BaseCode"]}[Content]{0}[Column1],
    Quote = Excel.CurrentWorkbook(){[Name="QuoteCode"]}[Content]{0}[Column1],

    // The host stays a constant and the changing part goes into
    // RelativePath. Concatenating the whole URL will often work on your own
    // machine with a manual refresh, but Power Query likes to know what its
    // data sources are before it runs anything, and an address assembled
    // out of cell values cannot be worked out in advance - which is what
    // the dynamic data source complaint is about. Same request on the
    // wire, far better behaved query.
    Source = Json.Document(
        Web.Contents(
            "https://api.frankfurter.dev",
            [RelativePath = "v2/rate/" & Base & "/" & Quote]
        )
    ),

    // The API returns a single record: one date, one base, one quote, one
    // rate. Excel wants a table, so turn it into a one-row table.
    Result = Table.FromRecords({Source})
in
    Result
Refreshing

Data > Refresh All, or Ctrl + Alt + F5, or right-click the output table > Refresh.

Editing a currency cell does not trigger a refresh on its own — Power Query has no live dependency on it the way a formula would. To refresh on open: right-click the query in the Queries & Connections pane > Properties > Refresh data when opening the file.

If you want a Refresh button, it needs a line of VBA, which puts you back into a macro-enabled workbook:

vba
Sub Refresh_Rate()
    ThisWorkbook.RefreshAll
End Sub
If something goes wrong

Status 404 with a message about the currency. The code in B2 or B3 is not one the API knows. That is the error handling working correctly.

The rate comes back as a strange number, or the last line errors. You are on a comma-decimal machine and something has been switched to CDbl. Use Val.

Power Query complains about privacy levels. That is the Formula Firewall. File > Options and settings > Query Options > Current Workbook > Privacy > Ignore the Privacy Levels. Workbook-scoped, so it changes nothing else on your machine.

Power Query complains about a dynamic data source. The URL has been concatenated into one string. Split it — host as a constant, changing part into RelativePath.

The credential dialog never appears. Excel remembers per host. Data > Get Data > Data Source Settings, find the api.frankfurter.dev entry, Clear Permissions.

Homework
Swap in a different pair of currency codes and confirm the sheet really is driving the whole thing.
The date is sitting in the same response, right next to the rate. Pull it into B6 using the same InStr and Mid technique. About three lines.
A note on this code

This is teaching code, written to match what happens on screen rather than to be the most robust version of itself. Pulling a number out of JSON with InStr and Mid is fragile — adding one field to that response could break it. A proper parser comes later in the series.

Don't lift the parsing into production. Do lift the shape of it.

The series
Your First API Call in Excel: VBA and Power Query ← you are here
GET vs POST
API Keys
Parsing JSON Properly
Handling a Response with PowerShell
Access Tokens and Expiry
Handling API Errors Gracefully

Each one stands on its own, so take them in whatever order suits you.
