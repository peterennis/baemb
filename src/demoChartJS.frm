Version =20
VersionRequired =20
PublishOption =1
Checksum =-1561099242
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7380
    DatasheetFontHeight =11
    ItemSuffix =7
    Left =195
    Top =255
    Right =7575
    Bottom =8340
    DatasheetGridlinesColor =14806254
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0xf20e8ebdd288e440
    End
    GUID = Begin
        0xd9ebe934fb4e8640a01cf32118f82d5c
    End
    NameMap = Begin
        0x0acc0e5500000000e85f08ffb6e68643a3c51a0c13e9002301000000682b92a8 ,
        0x5489e4400000000000000000640065006d006f00430068006100720074004a00 ,
        0x53005300750062003100000000000000e6d317b41b8a1246b847e23efd487cae ,
        0x01000000682b92a85489e4400000000000000000640065006d006f0043006800 ,
        0x6100720074004a00530053007500620032000000000000000000000000000000 ,
        0x00000000000000000c000000050000000000000000000000000000000000
    End
    Caption ="ChartJS Demo"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =8100
            Name ="Detail"
            GUID = Begin
                0x546c569153ad414ab7d24617e8cd6543
            End
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =60
                    Top =60
                    Width =7260
                    Height =3360
                    BorderColor =10921638
                    Name ="subChart"
                    GUID = Begin
                        0x750f42c3234b6b48b4726878ff7d4d39
                    End
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =3420
                End
                Begin Subform
                    OverlapFlags =215
                    Left =60
                    Top =4260
                    Width =3600
                    Height =3420
                    TabIndex =1
                    BorderColor =10921638
                    Name ="subDataset1"
                    SourceObject ="Form.demoChartJSSub1"
                    GUID = Begin
                        0x65aa38c3430dc249b7faf26ab4ac8ef9
                    End
                    GridlineColor =10921638
                    VerticalAnchor =1

                    LayoutCachedLeft =60
                    LayoutCachedTop =4260
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =7680
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =4020
                            Width =930
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label4"
                            Caption ="Dataset 1"
                            GUID = Begin
                                0xafddd5bf13ab614fad831028a46b1493
                            End
                            GridlineColor =10921638
                            VerticalAnchor =1
                            LayoutCachedLeft =60
                            LayoutCachedTop =4020
                            LayoutCachedWidth =990
                            LayoutCachedHeight =4335
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =3720
                    Top =4260
                    Width =3540
                    Height =3420
                    TabIndex =2
                    BorderColor =10921638
                    Name ="subDataset2"
                    SourceObject ="Form.demoChartJSSub2"
                    GUID = Begin
                        0x28e7c7c8f0bbbe4fa4bd68e739db5cd1
                    End
                    GridlineColor =10921638
                    VerticalAnchor =1

                    LayoutCachedLeft =3720
                    LayoutCachedTop =4260
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =7680
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3720
                            Top =4020
                            Width =930
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label6"
                            Caption ="Dataset 2"
                            GUID = Begin
                                0x8cbe68861a13f84e88cc364810bd68a9
                            End
                            GridlineColor =10921638
                            VerticalAnchor =1
                            LayoutCachedLeft =3720
                            LayoutCachedTop =4020
                            LayoutCachedWidth =4650
                            LayoutCachedHeight =4335
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private WithEvents bemb As BembObject
Attribute bemb.VB_VarHelpID = -1




Private Sub bemb_Initialized()
  
  bemb.SnippetsPath = CurrentProject.Path & "\bemb\snippets.js"
  
  addWindowResizeHandler
  addChartClickEventHandler
  
  RefreshChart
  
End Sub


Private Sub addChartClickEventHandler()

' Method one, build script here in VBA
'  Dim preReturnScript As String
'
'  preReturnScript = ""
'  preReturnScript = preReturnScript & "  var points = window.myLine.getPointsAtEvent(event); "
'  preReturnScript = preReturnScript & "  bembData = points[0].value + ', ' + points[1].value; "
'
'  bemb.AddEventHandler "#canvas", "click", Me, "bembEvent_ChartClick", preReturnScript


' Method two, load script from snippets
  Dim script As String
  
  script = bemb.GetSnippet("ChartClickHandlerSnippet")
  
  bemb.AddEventHandler "#canvas", "click", Me, "bembEvent_ChartClick", script

End Sub

Public Sub bembEvent_ChartClick()
  DoCmd.OpenForm "demoChartJSPopup", , , , , acDialog, "The values of the clicked point were " & bemb.LastData
End Sub


Private Sub addWindowResizeHandler()

  'In this one, the element selector for jQuery is the window object, so we
  'can't have that wrapped in quotes, thus we'll set the OmitElementSelectorQuoteWrap to true


' Method one, build script here in VBA
'  Dim s As String
'  s = ""
'  s = s & "Canvas.width = window.innerWidth;"
'  s = s & "Canvas.height = window.innerHeight;"
'
'  bemb.AddEventHandler "window", "resize", Me, "bembEvent_WindowResized", s, True


' Method two, load script from snippets
  Dim script As String
  
  script = bemb.GetSnippet("WindowResizedHandlerSnippet")
  
  bemb.AddEventHandler "window", "resize", Me, "bembEvent_WindowResized", script, True
  
End Sub


Public Sub bembEvent_WindowResized()
  RefreshChart
End Sub


Public Sub RefreshChart()

' METHOD ONE: Build script in VBA
'
'  Dim script As String
'  Dim rs As DAO.Recordset
'
'  script = "refreshChart({"
'  script = script & " labels: ['January','February','March','April','May','June','July'],"
'  script = script & " datasets: [{ label: 'My First dataset', fillColor: 'rgba(220,220,220,0.2)',"
'  script = script & "strokeColor:'rgba(220,220,220,1)',pointColor:'rgba(220,220,220,1)',pointStrokeColor:'#fff',"
'  script = script & "pointHighlightFill:'#fff',pointHighlightStroke:'rgba(220,220,220,1)',"
'  script = script & "data : ["
'
'  Set rs = CurrentDb.OpenRecordset("demoChartJSQuery1")
'
'  While Not rs.EOF
'    script = script & rs("TotalPerMonth") & ","
'    rs.MoveNext
'  Wend
'  rs.Close
'  'trim the last comma
'  script = Left(script, Len(script) - 1)
'
'  script = script & "]},{"
'  script = script & "label: 'My Second dataset',fillColor:'rgba(151,187,205,0.2)',strokeColor:'rgba(151,187,205,1)',"
'  script = script & "pointColor:'rgba(151,187,205,1)',pointStrokeColor:'#fff',pointHighlightFill:'#fff',pointHighlightStroke:'rgba(151,187,205,1)',"
'  script = script & "data: ["
'
'  Set rs = CurrentDb.OpenRecordset("demoChartJSQuery2")
'
'  While Not rs.EOF
'    script = script & rs("TotalPerMonth") & ","
'    rs.MoveNext
'  Wend
'  rs.Close
'
'  'trim the last comma
'  script = Left(script, Len(script) - 1)
'
'  script = script & "]}]"
'
'  script = script & "});"
'
'  bemb.ExecScript script

' METHOD TWO: Load Script from Snippet
'   |VALUE| is used to denote parameters, which we'll replace with generated data...

  Dim script As String
  Dim rs As DAO.Recordset
  Dim s As String
  
  script = bemb.GetSnippet("RefreshChartDataBuild")
  
  'load dataset1
  Set rs = CurrentDb.OpenRecordset("demoChartJSQuery1")

  While Not rs.EOF
    s = s & rs("TotalPerMonth") & ","
    rs.MoveNext
  Wend
  rs.Close
  
  'trim the last comma
  s = Left(s, Len(s) - 1)
  script = Replace(script, "|DATASET1|", s)
  
  
  'load dataset2
  s = ""
  Set rs = CurrentDb.OpenRecordset("demoChartJSQuery2")

  While Not rs.EOF
    s = s & rs("TotalPerMonth") & ","
    rs.MoveNext
  Wend
  rs.Close
  
  'trim the last comma
  s = Left(s, Len(s) - 1)
  script = Replace(script, "|DATASET2|", s)

  bemb.ExecScript script

End Sub

Private Sub btnRefreshChart_Click()
  RefreshChart
End Sub

Private Sub Form_Load()
  Set bemb = New BembObject
  
  bemb.AllowScroll = False
  
  bemb.Init Me.subChart, CurrentProject.Path & "\bemb\demoChartJS\charting.html"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not bemb Is Nothing Then Set bemb = Nothing
End Sub