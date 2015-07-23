Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'name of the form that is embedded during the Init operation
Const SOURCEOBJECT_NAME = "BembSubform"
Const DEFAULT_SNIPPET_START_QUALIFIER = "/// StartSnippet: "
Const DEFAULT_SNIPPETS_FILENAME = "bembSnippets.js"

Public Event Initialized()

Private WithEvents mHostForm As Form_BembSubform  'the subform that hosts the browser control
Attribute mHostForm.VB_VarHelpID = -1
Private wbControl As Object                       'the web browser control object (the actual control, not the .Object property of the control)
Private mEventReceiver As BembEventHandler        'this is for the core event bridge (see the js bemb framework docs), the only direct event correlation between JS and VBA
Private mInitialURL As String                     'the URL that was assigned at initialization.  We need to track it over a few calls, so... not used after complete intialization is done
                                                  'on a related note, there's no currently handling for page navigation/reloads... maybe in a future version
Private mAllowScroll As Boolean
Private mSnippetStartQualifier As String
Private mSnippetEndQualifier As String
Private mSnippetsPath As String

Private Type JS_EVENT_INFO
  ObjToCall As Object
  MethodToCall As String
End Type

'eventID passed to/from JS is the same as the array index
Private jsEvents() As JS_EVENT_INFO
'

Public Function GetSnippet(SnippetName As String) As String
    'returns ZLS if not found

    Dim ret As String
    Dim FileNum As Integer
    Dim v As Variant          'array of lines in the file
    Dim l As Long             'counter to loop the array
    Dim s As String           'content of the current line
    Dim InSnippet As Boolean  'let's us know if we've found (and are inside) the snippet during the loop

    Debug.Print "BembObject GetSnippet"
    Debug.Print , SnippetName & " Me.SnippetsPath = " & Me.SnippetsPath

    FileNum = FreeFile
    Open Me.SnippetsPath For Input As #FileNum
        v = Split(Input$(LOF(FileNum), #FileNum), vbCrLf)
    Close #FileNum

    For l = 0 To UBound(v)

        s = CStr(v(l))

        'if we're in a snippet, check for the closure
        If InSnippet Then

            If Left(s, Len(Me.SnippetStartQualifier)) = Me.SnippetStartQualifier Then
                GetSnippet = ret
                Exit Function
            End If

            ret = ret & s & vbCrLf

        Else 'not InSnippet

            'if we haven't found the snippet yet, check for it
            If Left(s, Len(Me.SnippetStartQualifier)) = Me.SnippetStartQualifier Then
                s = Trim(Replace(s, Me.SnippetStartQualifier, ""))
                If s = SnippetName Then
                    InSnippet = True
                End If
            End If

        End If

    Next l

    GetSnippet = ret

End Function

Public Property Get SnippetsPath() As String
    If mSnippetsPath = "" Then mSnippetsPath = CurrentProject.Path & "\" & DEFAULT_SNIPPETS_FILENAME
    SnippetsPath = mSnippetsPath
    'Debug.Print "BembObject SnippetsPath"
    'Debug.Print "SnippetsPath = " & SnippetsPath
End Property

Public Property Let SnippetsPath(s As String)
    mSnippetsPath = s
End Property

Public Property Get SnippetStartQualifier() As String
    If mSnippetStartQualifier = "" Then mSnippetStartQualifier = DEFAULT_SNIPPET_START_QUALIFIER
    SnippetStartQualifier = mSnippetStartQualifier
    'Debug.Print "BembObject SnippetStartQualifier"
    'Debug.Print "SnippetStartQualifier = " & SnippetStartQualifier
End Property

Public Property Let SnippetStartQualifier(s As String)
    mSnippetStartQualifier = s
End Property

Public Property Get Document() As Object
    Set Document = wbControl.Object.Document
End Property

Public Property Get AllowScroll() As Boolean
    AllowScroll = mAllowScroll
End Property

Public Property Let AllowScroll(b As Boolean)
    mAllowScroll = b
End Property

Private Sub FinalizeInitialization()
    Set mEventReceiver = New BembEventHandler
    mEventReceiver.Init Me, "EventReceived"
    wbControl.Object.Document.All("bemb-event-element").OnClick = mEventReceiver

    Me.Document.body.scroll = IIf(Me.AllowScroll, "yes", "no")
End Sub

Public Function LoadScriptFileToString(scriptFilePath As String) As String
'just a helper function... if user/dev wants to store script snippets in files,
'they can use this to load it into a string (hell of a lot easier than trying to
'type javascript code to a string in the VBE...)

    On Error GoTo Err_Proc

    Dim ret As String
    Dim i As Integer

    If Len(Dir(scriptFilePath)) = 0 Then GoTo Exit_Proc

    i = FreeFile
    Open scriptFilePath For Input As #i
        ret = Input(LOF(i), i)
    Close #i

Exit_Proc:
    LoadScriptFileToString = ret
    Exit Function

Err_Proc:
    Err.Source = "BembObject.LoadScriptFileToString"
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error"
    End Select
    Resume Exit_Proc
    Resume
End Function

Public Function Init(SubformContainer As Access.SubForm, URL As String) As Boolean

    On Error GoTo Err_Proc

    Dim ret As Boolean

    SubformContainer.SourceObject = SOURCEOBJECT_NAME
    Set mHostForm = SubformContainer.Form
    Set wbControl = mHostForm.Controls(mHostForm.BrowserControlName)
    mInitialURL = URL

    ret = True

Exit_Proc:
    Init = ret
    Exit Function

Err_Proc:
    Err.Source = "BembObject.Init"
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error"
    End Select
    Resume Exit_Proc
    Resume
End Function

Public Function Eval(CodeToRun As String, Optional postFunction As String = "") As String
    Eval = mHostForm.Eval(CodeToRun, postFunction)
End Function

Public Property Get jQueryNativeVersion() As String
    jQueryNativeVersion = mHostForm.Eval("jQuery.fn.jquery;")
End Property

Public Sub EventReceived()
    'public only for calling by event handler, not for user consumption
    Dim EventID As Long
    EventID = wbControl.Object.Document.script.bembRaisedEventID
    CallByName jsEvents(EventID).ObjToCall, jsEvents(EventID).MethodToCall, VbMethod
End Sub

Public Property Get LastData() As String
    LastData = mHostForm.bembData
End Property

Public Sub AddEventHandler( _
        ByVal ElementSelector As String, _
        jsEventName As String, _
        ReceivingObject As Object, _
        ReceivingMethodName As String, _
        Optional jsPostEventFunction As String = "", _
        Optional jsOmitElementSelectorQuoteWrap As Boolean = False _
        )
    'most times, we want to wrap the element selector in quotes for jQuery:
    '   mjvb_$('#canvas')
    'occassionally though, we can't:
    '   mjvb_$(window).on('resize');
    'Thus, the OmitElementSelectorQuoteWrap optional arg
    'This also makes it possible to use js variables or js object reference as selectors
    '
    'We've set ElementSelector to ByVal so we can wrap/ignore as requested

    Dim EventID As Long

    If Not jsOmitElementSelectorQuoteWrap Then ElementSelector = "'" & ElementSelector & "'"

    EventID = AddToJSEvents()

    With jsEvents(EventID)
        Set .ObjToCall = ReceivingObject
        .MethodToCall = ReceivingMethodName
    End With

    If jsPostEventFunction = "" Then
        Me.ExecScript "bemb.addEventListener(" & ElementSelector & ", '" & jsEventName & "', " & EventID & ");"
    Else
        Me.ExecScript "bemb.addEventListener(" & ElementSelector & ", '" & jsEventName & "', " & EventID & ", function() { " & jsPostEventFunction & "});"
    End If

End Sub

Public Sub ExecScript(script As String)
    mHostForm.Exec script
End Sub

Private Function JsEventsInitialized() As Boolean
    On Error Resume Next
    Dim x As String
    x = jsEvents(0).MethodToCall
    JsEventsInitialized = Not CBool(Err.Number)
End Function

Private Function AddToJSEvents() As Long
    'returns the zero based index that was added
    Dim Idx As Long
    Dim info As JS_EVENT_INFO

    If JsEventsInitialized Then
        Idx = UBound(jsEvents) + 1
        ReDim Preserve jsEvents(0 To Idx)
    Else
        Idx = 0
        ReDim jsEvents(0)
    End If

    AddToJSEvents = Idx

End Function

Private Sub mHostForm_InitializationComplete()
    FinalizeInitialization
    RaiseEvent Initialized
End Sub

Private Sub mHostForm_InitialNavReady()
    wbControl.Navigate2 mInitialURL
End Sub