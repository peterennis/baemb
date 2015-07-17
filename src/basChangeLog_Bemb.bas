Option Compare Database
Option Explicit

' Constants for settings of "Bemb"
Public Const gstrPROJECT_BEMB As String = "Bemb"
Private Const mstrVERSION_BEMB As String = "0.5.0.0"
Private Const mstrDATE_BEMB As String = "July 17, 2015"

Public Const THE_SOURCE_FOLDER = "C:\ae\bemb\src\"
Public Const THE_BACK_END_SOURCE_FOLDER = "C:\ae\bemb\srcbe\"
Public Const THE_XML_FOLDER = "C:\ae\bemb\src\xml\"
Public Const THE_BACK_END_XML_FOLDER = "C:\ae\bemb\srcbe\xml\"
Public Const THE_BACK_END_DB1 = "NONE"
'

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = mstrVERSION_BEMB
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = mstrDATE_BEMB
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gstrPROJECT_BEMB
End Function

Public Sub BEMB_EXPORT(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    If Not IsMissing(varDebug) Then
        aegitClassTest varDebug:="varDebug", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlData:=THE_XML_FOLDER, varBackEndDb1:=THE_BACK_END_DB1
    Else
        aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlData:=THE_XML_FOLDER, varBackEndDb1:=THE_BACK_END_DB1
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure SVIPDB_EXPORT"
    Resume Next

End Sub

'=============================================================================================================================
' Tasks:
' %005 -
' %004 -
' %003 -
' %002 -
' %001 -
' #################
' Issues:
' #005 -
' #004 -
' #003 -
' #002 -
' #001 -
'=============================================================================================================================
'
'
'20150717 - v0500 -
    ' Initial setup for code export