VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNameValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrName As String
Private mstrValue As String
Private mstrDelim As String

'Base name value pair
Public Property Let Name(str As String)
    
    mstrName = str
    
End Property

Public Property Get Name() As String

    Name = mstrName

End Property

Public Property Let ValueStr(str As String)
    
    mstrValue = str
    mstrDelim = """"
    
End Property

Public Property Let ValueDt(dt As Date)

    mstrValue = Format(dt, "yyyy-dd-MMThh:mm:ss")
    mstrDelim = """"
    
End Property

Public Property Let Value(var As Variant)

    mstrValue = CStr(var)
    mstrDelim = ""
    
End Property

Public Property Get Value() As String

    Value = mstrValue

End Property

Public Property Let Delim(Value As String)
    
    mstrDelim = Value
    
End Property

Public Property Get Delim() As String

    Delim = mstrDelim

End Property

Public Function JSON() As String

    JSON = StringFormat("{0}:{1}{2}{1}", Name, mstrDelim, mstrValue)

End Function
