Code allows convert text to byte array and vice versa for most popular code pages.

## Features:
* Convert text to bytes
* Convert bytes to text
* Full code page list with filtering
* Live preview when code page is changing
* Scan-codes AT and XT
* Base64 encoding

## Usage:

```vbnet
Dim cpc As New CodepageConverter()

cpc.TextEncoding = CodepageConverter.TextEncodings.Ascii

Dim s As String = cpc.GetStringEncoded()
```
