Paste this into your README.md or video page:
# GET vs POST in Excel VBA

In this tutorial, we build two API requests in Excel VBA that send the same business data to the same service in two different ways.

The first is a GET request.

The second is a POST request.

Along the way, we also look at:

- Query strings
- Request bodies
- Postman Echo
- URL encoding
- Percent encoding
- Application.EncodeURL
- JSON request bodies
- Content-Type
- HTTP status codes
- Debug.Print
- WinHttp.WinHttpRequest.5.1

---

## The Example

Our Excel worksheet contains:

| Cell | Description | Example |
|---|---|---|
| B2 | Product Code | SKU-1044 |
| B3 | Quantity | 25 |
| B6 | HTTP Status | Filled by VBA |
| B7 | Raw Response | Filled by VBA |

We create two macros:

- Check_Stock_GET
- Submit_Order_POST

Both send the same product code and quantity.

The difference is how the request is constructed.

---

# GET Request

A GET request commonly sends input values as part of the URL.

For example:

https://postman-echo.com/get?product=SKU-1044&quantity=25

Everything after the question mark is the query string.

The structure is:

?name=value&name=value

The first parameter starts after a question mark.

Additional parameters are separated using an ampersand.

---

## Building the GET Request in VBA

We read the values from Excel:

```text
Product Code: SKU-1044
Quantity: 25

Then build the URL.
The request is sent using:
WinHttp.WinHttpRequest.5.1

The basic pattern is:
Open the request
Send the request
Read the status
Read the response

For a GET request, Send does not need a request body.
Debug.Print
Before sending a URL that we have constructed manually, it is useful to inspect it.
The example therefore uses:
Debug.Print url

The URL appears in the VBA Immediate Window.
This makes it much easier to see what our code actually constructed before trying to diagnose a problem at the API.
The GET Problem: Special Characters
Suppose the product code changes from:
SKU-1044

to:
SKU&1044

The ampersand already has a structural meaning inside a query string.
It separates parameters.
So a URL containing:
product=SKU&1044

may be interpreted incorrectly by the server.
The ampersand that belongs to our data can be mistaken for the start of another parameter.
This can be particularly dangerous because the HTTP request itself may still return a successful status such as 200.
The request succeeded technically, but the server received the wrong data.
URL Encoding
The solution is URL encoding, also known as percent encoding.
Characters that have a special meaning inside a URL are represented using an encoded value.
For example:
&

becomes:
%26

A space commonly becomes:
%20

A quotation mark becomes:
%22

So:
SKU&1044

becomes:
SKU%261044

The server understands this encoding and converts the value back when it receives the request.
Application.EncodeURL
Excel provides a built-in function that can encode values for us:
Application.EncodeURL

The important rule is:
Encode the value, not the entire URL.
For example:
Application.EncodeURL(productCode)

We do not wrap the entire finished URL in EncodeURL.
The URL contains structural characters such as:
:
/
?
=
&

Those characters have a job to do.
We only encode the values being inserted into that structure.
POST Request
The second macro sends the same product code and quantity using POST.
Instead of adding the values to the URL, we send them inside a request body.
The URL stays simple:
https://postman-echo.com/post

The data is placed inside JSON.
For example:
{
  "product": "SKU-1044",
  "quantity": 25
}

JSON Data Types
Notice that the product code is surrounded by quotation marks:
"product": "SKU-1044"

It is text.
The quantity is not surrounded by quotation marks:
"quantity": 25

It is a number.
For that reason, the VBA example stores quantity as a Long when building the POST body.
Debugging the JSON Body
Just as we inspect the URL before sending the GET request, we inspect the JSON body before sending the POST request.
The example uses:
Debug.Print jsonBody

This allows us to check the exact JSON VBA created before it is sent.
Content-Type
When we send JSON, we also tell the server what format we are sending.
We do this using the Content-Type header:
Content-Type: application/json

In VBA:
http.setRequestHeader "Content-Type", "application/json"

The body contains the data.
The Content-Type tells the server how that data should be interpreted.
These are two separate jobs.
Sending the POST Body
For the GET request, the call to Send is empty:
http.Send

For the POST request, we pass the JSON body:
http.Send jsonBody

That is the point at which the JSON becomes the request body.
GET vs POST
For the pattern demonstrated in this tutorial:
GET
Values are placed in the query string.
https://example.com/get?product=SKU-1044&quantity=25

VBA:
http.Open "GET", url, False
http.Send

Values inserted into the URL should be URL encoded.
POST
The URL remains clean.
https://example.com/post

The data is sent inside the request body.
VBA:
http.Open "POST", url, False
http.setRequestHeader "Content-Type", "application/json"
http.Send jsonBody

Postman Echo
The examples use Postman Echo:
https://postman-echo.com

Postman Echo is useful because it returns information about the request it received.
This allows us to inspect things such as:
- Query parameters
- Headers
- Request body
- URL
It effectively acts as a mirror for an HTTP request.
This makes echo endpoints particularly useful when debugging API integrations.
Important Lesson
When consuming somebody else's API, we normally do not decide whether an operation should use GET or POST.
The API documentation tells us which method to use.
GET is generally associated with reading information.
POST is commonly used when submitting content for processing.
However, this is not the same as saying:
GET = SELECT
POST = INSERT

APIs do not map directly to database operations.
For example, an API can legitimately use POST for an operation that ultimately returns information.
Always follow the API documentation.
Security Note
Avoid putting sensitive information such as API keys, access tokens or personal information into query strings unless the API specifically requires it.
URLs can appear in places such as:
- Server logs
- Proxy logs
- Monitoring systems
- Application logs
Authentication information normally has its own appropriate header or authentication mechanism.
Complete VBA Code
The complete VBA used in this tutorial is included below.
The code uses late binding:
CreateObject("WinHttp.WinHttpRequest.5.1")

This means you do not need to manually add a VBA reference for WinHTTP.
This example is intended for Windows Excel because WinHTTP is a Windows component.
Notes
Application.EncodeURL is the short form used in the video.
Excel also exposes the function as:
WorksheetFunction.EncodeURL

When writing API responses directly into an Excel cell, remember that an Excel cell can hold a maximum of 32,767 characters.
For larger API responses, keep the response in a VBA variable or process it before writing data to the worksheet.

### 2. Complete raw VBA code

You can paste this directly into a `.bas`, `.txt`, or GitHub code file:

```vb
Option Explicit

Sub Check_Stock_GET()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Orders")

    Dim productCode As String
    Dim quantity As String

    productCode = CStr(ws.Range("B2").Value)
    quantity = CStr(ws.Range("B3").Value)

    ' GET - the inputs travel in the address,
    ' after a "?", separated by "&"
