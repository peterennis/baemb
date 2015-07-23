Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Web Browser Control - Handling Events in Visual Basic Applications
' http://msdn.microsoft.com/en-us/library/aa752045(VS.85).aspx

Private mObj As Object
Private mMethod As String

Public Sub Init(obj As Object, mthd As String)
    'takes the WebBrowserHost instance's object and the method "EventReceived"
    Set mObj = obj
    mMethod = mthd
End Sub

Public Sub Receiver()
Attribute Receiver.VB_UserMemId = 0
    'Attribute Value.VB_UserMemId = 0
    'this is a default method, must be specified by applying the attribute via
    'text editor and imported into VBA
    If Not mObj Is Nothing Then CallByName mObj, mMethod, VbMethod
End Sub