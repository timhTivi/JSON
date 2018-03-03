Attribute VB_Name = "modUtils"
Option Explicit

Public Function StringFormat(strFormat As String, ParamArray args() As Variant) As String

On Error GoTo errHandler

Dim i As Integer
Dim strOut As String

    strOut = strFormat
    
    For i = 0 To UBound(args)
    
        strOut = Replace(strOut, "{" & i & "}", CStr(args(i)))
    
    Next i
    
    StringFormat = strOut
    
Exit Function

errHandler:

    StringFormat = ""

End Function

Public Function SafeUbound(var() As Variant) As Integer

On Error GoTo errHandler

    SafeUbound = UBound(var)

Exit Function

errHandler:

    SafeUbound = -1

End Function

Public Function Wrap(strLeftDelim As String, strRightDelim As String, strToBeWrapped As String) As String

    Wrap = strLeftDelim & strToBeWrapped & strRightDelim

End Function

Public Function WrapSq(strToBeWrapped As String) As String

    WrapSq = "[" & strToBeWrapped & "]"

End Function

Public Function WrapBr(strToBeWrapped As String) As String

    WrapBr = "{" & strToBeWrapped & "}"

End Function



