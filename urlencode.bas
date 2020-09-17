Option Base 0
Option Explicit

'' Convert an unicode integer into a utf-8 byte array
' https://en.wikipedia.org/wiki/UTF-8
' @method Unicode2Utf8
' @param {Long} uni unicode code to convert to utf8
' @return Byte() byte array
' for numbers greater than &H10FFFF&, return an empty array
'
' Note: vba does not have bitwise shift operators,
' use division by 2 ^ n instead
'
Public Function Unicode2Utf8(uni As Long) As Byte()

    Dim nBytes As Long
    Dim utf8Buf(3) As Byte
    
    Select Case uni
        Case 0 To &H7F
            utf8Buf(0) = uni
            
        Case &H80 To &H7FF
        ' 110xxxxx    10xxxxxx
            utf8Buf(0) = &HC0 Or (uni \ (2 ^ 6))
            utf8Buf(1) = &H80 Or (&H3F And uni)
            
        Case &H800 To &HFFFF&
        ' 1110xxxx    10xxxxxx    10xxxxxx
            utf8Buf(0) = &HE0 Or (uni \ (2 ^ 12))
            utf8Buf(1) = &H80 Or (&H3F And (uni \ (2 ^ 6)))
            utf8Buf(2) = &H80 Or (&H3F And uni)
            
        Case &H10000 To &H10FFFF
        ' 11110xxx  10xxxxxx    10xxxxxx    10xxxxxx
            utf8Buf(0) = &HF0 Or (uni \ (2 ^ 18))
            utf8Buf(1) = &H80 Or (&H3F And (uni \ (2 ^ 12)))
            utf8Buf(2) = &H80 Or (&H3F And (uni \ (2 ^ 6)))
            utf8Buf(3) = &H80 Or (&H3F And uni)
            
        Case Else
            ' do nothing
    End Select
    
    Unicode2Utf8 = utf8Buf
End Function

'' URL encode
' ALPHA / DIGIT / "-" / "." / "_" / "*" / " " are not encoded
' space is encoded as %20
' characters with the ascii code up to 255 (ff) are encoded as %xy
' characters with ascii codes between 256 and 65535 (ffff) are first converted
' to utf-8, then each byte is encoded as %xy
' This does not work correctly for unicode greater than 65535 (ffff)
' I did not find a way to make it work in vba.
' Neither AscW(c) nor StrConv(c, vbUnicode) work correctly for characters 
' represented by more than two bytes
' @method UrlEncode
' @param {Variant} Text Text to encode
' @return {String} Encoded string
'
''
Public Function UrlEncode(Text As Variant) As String

    Dim url As String
    Dim ch, chHex As String
    Dim chCode As Long
    Dim utf8Buf() As Byte
    Dim i As Long
    Dim j As Integer
    
    url = CStr(Text)
    
    UrlEncode = ""
    
    For i = 1 To Len(url)
        ' Get character and its 2-byte code
        ' This does not work correctly for characters respresented by more than two bytes
        ch = Mid$(url, i, 1)
        chCode = CLng(AscW(ch)) And &HFFFF&

        Select Case chCode
            Case 65 To 90, 97 To 122
            ' alpha
                UrlEncode = UrlEncode & ch
                
            Case 48 To 57
            ' digit
                UrlEncode = UrlEncode & ch
                
            Case 45, 46, 95, 126
                ' "-" / "." / "_" / "~"
                UrlEncode = UrlEncode & ch
                
            Case 32
                ' space
                UrlEncode = UrlEncode & "%20"
    
            Case 0 To 127
                ' all other ascii
                ' Hex() does not return the leading 0
                UrlEncode = UrlEncode & "%0" & Hex(chCode)
                
            Case 127 To 255
                ' extended ascii
                UrlEncode = UrlEncode & "%" & Hex(chCode)
            
            Case Else
                ' convert to utf-8 first
                utf8Buf = Unicode2Utf8(chCode)                               
                
                For j = 0 To 3
                    If utf8Buf(j) Then
                        UrlEncode = UrlEncode & "%" & Hex(utf8Buf(j))
                    Else
                        Exit For
                    End If
                Next j
          
        End Select
    Next i
    
End Function



