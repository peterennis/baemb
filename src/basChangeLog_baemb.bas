Option Compare Database
Option Explicit

' Constants for settings of "Baemb"
Public Const BAEMB_DEMO_PATH = "\demo"
Public Const BEMB_SCRIPT_PATH = "\bemb"
Public Const BAEMB_SCRIPT_PATH = "\baemb"
Public Const BAEMBPIE_SCRIPT_PATH = "\baembPie"
Public gstrBAEMB_DEMO As String

Public Const gstrPROJECT_BAEMB As String = "baemb"
Private Const mstrVERSION_BAEMB As String = "0.5.0.5"
Private Const mstrDATE_BAEMB As String = "August 1, 2015"

Public Const THE_SOURCE_FOLDER = "C:\ae\baemb\src\"
Public Const THE_BACK_END_SOURCE_FOLDER = "C:\ae\baemb\srcbe\"
Public Const THE_XML_FOLDER = "C:\ae\baemb\src\xml\"
Public Const THE_BACK_END_XML_FOLDER = "C:\ae\baemb\srcbe\xml\"
Public Const THE_BACK_END_DB1 = "NONE"
'

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = mstrVERSION_BAEMB
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = mstrDATE_BAEMB
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gstrPROJECT_BAEMB
End Function

Public Sub BAEMB_EXPORT(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    If Not IsMissing(varDebug) Then
        aegitClassTest varDebug:="varDebug", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlData:=THE_XML_FOLDER, varBackEndDb1:=THE_BACK_END_DB1
    Else
        aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlData:=THE_XML_FOLDER, varBackEndDb1:=THE_BACK_END_DB1
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BAEMB_EXPORT"
    Resume Next

End Sub

'=============================================================================================================================
' Tasks:
' %010 -
' %009 -
' %008 -
' %007 - Performance metrics and testing framework
' %006 - Load latest aegit and test export with relative paths
' %005 - Consider Msi installer package with SHA signature
' %004 - Create muilti resolution project icon and branding - the baemb pronunciation and potential disruption effect
' %003 - http://clipartist.net/svg/bomb-scalable-vector-graphics-svg/ - add project icon
' %002 - Add debug tracking info
'=============================================================================================================================
'
'
'20150723 - v0503 - Added demoBaembChartJS, Ref: http://stackoverflow.com/questions/19847582/chart-js-canvas-resize
'20150720 - v0501 - FIXED - %001 - Turn off chart animation Ref: http://www.chartjs.org/docs/ - Global chart configuration
'20150717 - v0501 - Export new release to review and test fixes for #1 and #2 errors on github
'20150717 - v0500 - Initial setup for code export