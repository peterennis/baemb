Version =20
VersionRequired =20
PublishOption =1
Checksum =1699062140
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    DatasheetFontHeight =11
    ItemSuffix =2
    Left =4500
    Top =1620
    Right =15165
    Bottom =11235
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xd1b687925489e440
    End
    GUID = Begin
        0xe85f08ffb6e68643a3c51a0c13e90023
    End
    NameMap = Begin
        0x0acc0e550000000069df445e2d673a4a84bcead4d8ca044900000000de857178 ,
        0x5489e4400000000000000000640065006d006f00430068006100720074004a00 ,
        0x53005400610062006c006500310000000000000020d4f119b85f304893b4f296 ,
        0xeedf3b520700000069df445e2d673a4a84bcead4d8ca04494500760065006e00 ,
        0x74004400610074006500000000000000a5b38cac38781447bab8b0b02fc20836 ,
        0x0700000069df445e2d673a4a84bcead4d8ca04494500760065006e0074005600 ,
        0x61006c0075006500000000000000000000000000000000000000000000000c00 ,
        0x0000050000000000000000000000000000000000
    End
    RecordSource ="demoChartJSTable1"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7560
            Name ="Detail"
            GUID = Begin
                0x35d7b1a7a243cc4d824e966ebd03733f
            End
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5640
                    Top =1500
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="EventDate"
                    ControlSource ="EventDate"
                    GUID = Begin
                        0x28ea6aaba8255d4fb77ce9b5c401bd88
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =1500
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =1815
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3840
                            Top =1500
                            Width =1035
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label0"
                            Caption ="EventDate"
                            GUID = Begin
                                0xc9c53655e44525438e7ce9366432831b
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =3840
                            LayoutCachedTop =1500
                            LayoutCachedWidth =4875
                            LayoutCachedHeight =1815
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5640
                    Top =1980
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="EventValue"
                    ControlSource ="EventValue"
                    GUID = Begin
                        0xa3c557383aad6049bc8aa0d04e01f76b
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =2295
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3840
                            Top =1980
                            Width =1140
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label1"
                            Caption ="EventValue"
                            GUID = Begin
                                0x2c299f1db2c08b4fb21ccde25653e534
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =3840
                            LayoutCachedTop =1980
                            LayoutCachedWidth =4980
                            LayoutCachedHeight =2295
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

Private Sub Form_AfterUpdate()
  Me.Parent.RefreshChart
End Sub