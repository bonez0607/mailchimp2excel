Attribute VB_Name = "Base64Encoder"
' **Code taken from this stackoverflow answer https://stackoverflow.com/a/40118072
' Base64-encodes the specified string.
' Parameter fAsUtf16LE determines how the input text is encoded at the
' byte level before Base64 encoding is applied.
' * Pass False to use UTF-8 encoding.
' * Pass True to use UTF-16 LE encoding.
Function Base64Encode(ByVal sText)

    ' Use an aux. XML document with a Base64-encoded element.
    ' Assigning the byte stream (array) returned by StrToBytes() to .NodeTypedValue
    ' automatically performs Base64-encoding, whose result can then be accessed
    ' as the element's text.
    With CreateObject("Msxml2.DOMDocument").CreateElement("aux")
        .DataType = "bin.base64"
        .NodeTypedValue = StrToBytes(sText, "utf-8", 3)

        Base64Encode = .Text
    End With

End Function

' Returns a binary representation (byte array) of the specified string in
' the specified text encoding, such as "utf-8" or "utf-16le".
' Pass the number of bytes that the encoding's BOM uses as iBomByteCount;
' pass 0 to include the BOM in the output.
Function StrToBytes(ByVal sText, ByVal sTextEncoding, ByVal iBomByteCount)

    ' Create a text string with the specified encoding and then
    ' get its binary (byte array) representation.
    With CreateObject("ADODB.Stream")
        ' Create a stream with the specified text encoding...
        .Type = 2  ' adTypeText
        .Charset = sTextEncoding
        .Open
        .WriteText sText
        ' ... and convert it to a binary stream to get a byte-array
        ' representation.
        .Position = 0
        .Type = 1  ' adTypeBinary
        .Position = iBomByteCount ' skip the BOM
        StrToBytes = .Read
        .Close
    End With

End Function

' Returns a string that corresponds to the specified byte array, interpreted
' with the specified text encoding, such as "utf-8" or "utf-16le".
Function BytesToStr(ByVal byteArray, ByVal sTextEncoding)

    If LCase(sTextEncoding) = "utf-16le" Then
        ' UTF-16 LE happens to be VBScript's internal encoding, so we can
        ' take a shortcut and use CStr() to directly convert the byte array
        ' to a string.
        BytesToStr = CStr(byteArray)
    Else ' Convert the specified text encoding to a VBScript string.
        ' Create a binary stream and copy the input byte array to it.
        With CreateObject("ADODB.Stream")
            .Type = 1 ' adTypeBinary
            .Open
            .Write byteArray
            ' Now change the type to text, set the encoding, and output the
            ' result as text.
            .Position = 0
            .Type = 2 ' adTypeText
            .Charset = sTextEncoding
            BytesToStr = .ReadText
            .Close
        End With
    End If

End Function
