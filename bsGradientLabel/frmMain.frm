VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\..\..\-_MAIN~1\Projects\BSGRAD~1\Source\bsGradientLabel.vbp"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\AColour picker\ClrPckr.vbp"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "bsGradientLabel demo"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3840
      TabIndex        =   25
      Top             =   2040
      Width           =   4935
      Begin ClrPckr.ColorPicker cpText 
         Height          =   285
         Left            =   720
         TabIndex        =   36
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Select a fount"
         Height          =   375
         Left            =   2160
         TabIndex        =   33
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   30
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   2160
         List            =   "frmMain.frx":044F
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmMain.frx":0468
         Left            =   120
         List            =   "frmMain.frx":0475
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Colour"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Offset"
         Height          =   195
         Left            =   4200
         TabIndex        =   31
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Text Alignment"
         Height          =   195
         Left            =   2160
         TabIndex        =   29
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label Type"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Appearance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6360
      TabIndex        =   22
      Top             =   480
      Width           =   2415
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMain.frx":04A2
         Left            =   240
         List            =   "frmMain.frx":04BE
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Border Style"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Gradient"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3840
      TabIndex        =   17
      Top             =   480
      Width           =   2415
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1936
         TabIndex        =   35
         Top             =   1035
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "Text1"
         BuddyDispid     =   196622
         OrigLeft        =   2040
         OrigTop         =   960
         OrigRight       =   2280
         OrigBottom      =   1335
         Max             =   359
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1035
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":050C
         Left            =   240
         List            =   "frmMain.frx":0519
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Angle"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   360
      End
   End
   Begin bsGLabel.bsGradientLabel bsGradientLabel1 
      Height          =   6135
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   9975
      Caption         =   "bsGradientLabel Demo"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      LabelType       =   1
      GradientAngle   =   90
   End
   Begin bsGLabel.bsGradientLabel glTest 
      Height          =   2535
      Left            =   1200
      TabIndex        =   15
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4471
      Caption         =   "bsGradientLabel1"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   32768
      Colour1         =   8421631
      Colour2         =   16761024
      Colour4         =   12648384
      WordWrap        =   -1  'True
   End
   Begin bsGLabel.bsGradientLabel glVersion 
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   450
      Caption         =   "Hello World"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   -2147483630
      Colour1         =   -2147483632
      Colour2         =   -2147483633
      TextShadowColour=   33023
      TextShadowYOffset=   1
      TextShadowXOffset=   1
   End
   Begin VB.CheckBox chkMultiline 
      Caption         =   "Multiline label"
      Height          =   195
      Left            =   2760
      TabIndex        =   11
      Top             =   5880
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   1200
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   69891
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      TabIndex        =   10
      Top             =   3600
      Width           =   4935
      Begin ClrPckr.ColorPicker cpBG1 
         Height          =   285
         Left            =   1080
         TabIndex        =   38
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin ClrPckr.ColorPicker cpBG2 
         Height          =   285
         Left            =   1680
         TabIndex        =   39
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin ClrPckr.ColorPicker cpBG3 
         Height          =   285
         Left            =   1680
         TabIndex        =   40
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin ClrPckr.ColorPicker cpBG4 
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin ClrPckr.ColorPicker cpFlatBorder 
         Height          =   285
         Left            =   3000
         TabIndex        =   42
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin ClrPckr.ColorPicker cpBorder1 
         Height          =   285
         Left            =   3600
         TabIndex        =   43
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin ClrPckr.ColorPicker cpBorder2 
         Height          =   285
         Left            =   4200
         TabIndex        =   44
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin ClrPckr.ColorPicker cpBorder3 
         Height          =   285
         Left            =   3600
         TabIndex        =   45
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin ClrPckr.ColorPicker cpBorder4 
         Height          =   285
         Left            =   4200
         TabIndex        =   46
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Border"
         Height          =   195
         Left            =   2400
         TabIndex        =   13
         Top             =   405
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Background"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   405
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text Shadow"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1200
      TabIndex        =   4
      Top             =   3120
      Width           =   2535
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   315
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Text Shadow"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
      Begin ClrPckr.ColorPicker cpTextShadow 
         Height          =   285
         Left            =   960
         TabIndex        =   37
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BackColor       =   -2147483633
         ShowDefault     =   0   'False
         ShowCustomColors=   0   'False
         MoreColorsCaption=   "More..."
         ShowSysColorButton=   0   'False
         ShowToolTips    =   0   'False
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Colour"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Y Offset"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   765
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "X Offset"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMain.frx":0532
      Top             =   4920
      Width           =   6015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About this control"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Caption"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   4920
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Check1_Click()
   glTest.TextShadow = (Check1.Value = 1)
End Sub

Private Sub chkMultiline_Click()
   glTest.WordWrap = (chkMultiline.Value = 1)
End Sub

Private Sub Combo1_Click()
   glTest.GradientType = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
   glTest.BorderStyle = Combo2.ListIndex
End Sub

Private Sub Combo3_Click()
   glTest.CaptionAlignment = Combo3.ListIndex
End Sub

Private Sub Combo4_Click()
   glTest.LabelType = Combo4.ListIndex
End Sub

Private Sub Command1_Click()
   glTest.ShowAbout
End Sub

Private Sub Command2_Click()
   End
End Sub

Private Sub Command3_Click()

   Dim temp As New StdFont
   
   On Error GoTo forget_it
   dlgFont.ShowFont
   With temp
      .Name = dlgFont.FontName
      .Bold = dlgFont.FontBold
      .Italic = dlgFont.FontItalic
      .Underline = dlgFont.FontUnderline
      .Size = dlgFont.FontSize
   End With
   
   Set glTest.Fount = temp
forget_it:
   
End Sub

Private Sub Command7_Click()

End Sub

Private Sub cpBG1_Click()
   glTest.Colour1 = cpBG1.Color
End Sub

Private Sub cpBG2_Click()
   glTest.Colour2 = cpBG2.Color
End Sub

Private Sub cpBG3_Click()
   glTest.Colour3 = cpBG3.Color
End Sub

Private Sub cpBG4_Click()
   glTest.Colour4 = cpBG4.Color
End Sub

Private Sub cpBorder1_Click()
   glTest.HighlightColour = cpBorder1.Color
End Sub

Private Sub cpBorder2_Click()
   glTest.HighlightDKColour = cpBorder2.Color
End Sub

Private Sub cpBorder3_Click()
   glTest.ShadowColour = cpBorder3.Color
End Sub

Private Sub cpBorder4_Click()
   glTest.ShadowDKColour = cpBorder4.Color
End Sub

Private Sub cpFlatBorder_Click()
   glTest.FlatBorderColour = cpFlatBorder.Color
End Sub

Private Sub cpText_Click()
   glTest.CaptionColour = cpText.Color
End Sub

Private Sub cpTextShadow_Click()
   glTest.TextShadowColour = cpTextShadow.Color
End Sub

Private Sub Form_Load()
   glVersion.Caption = "Version " & glVersion.Version
   Combo1.ListIndex = glTest.GradientType
   Combo2.ListIndex = glTest.BorderStyle
   Combo3.ListIndex = glTest.CaptionAlignment
   Combo4.ListIndex = glTest.LabelType
   Check1.Value = Abs(glTest.TextShadow)
   Text1.Text = glTest.GradientAngle
   Text2.Text = glTest.Caption
   Text3.Text = glTest.Offset
   Text4.Text = glTest.TextShadowXOffset
   Text5.Text = glTest.TextShadowyOffset
   cpText.Color = glTest.CaptionColour
   cpTextShadow.Color = glTest.TextShadowColour
   cpBG1.Color = glTest.Colour1
   cpBG2.Color = glTest.Colour2
   cpBG3.Color = glTest.Colour3
   cpBG4.Color = glTest.Colour4
   cpBorder1.Color = glTest.HighlightColour
   cpBorder2.Color = glTest.HighlightDKColour
   cpBorder3.Color = glTest.ShadowColour
   cpBorder4.Color = glTest.ShadowDKColour
   cpFlatBorder.Color = glTest.FlatBorderColour
   chkMultiline.Value = Abs(glTest.WordWrap = True)
   
   With glTest.Fount
      dlgFont.FontName = .Name
      dlgFont.FontBold = .Bold
      dlgFont.FontItalic = .Italic
      dlgFont.FontUnderline = .Underline
      dlgFont.FontSize = .Size
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub glTest_Click()
   MsgBox "Yo!"
End Sub

Private Sub Text1_Change()
   glTest.GradientAngle = Val(Text1.Text)
End Sub

Private Sub Text2_Change()
   glTest.Caption = Text2.Text
End Sub

Private Sub Text3_Change()
   glTest.Offset = Val(Text3.Text)
End Sub

Private Sub Text4_Change()
   glTest.TextShadowXOffset = Val(Text4.Text)
End Sub

Private Sub Text5_Change()
   glTest.TextShadowyOffset = Val(Text5.Text)
End Sub
