VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8070
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Okay"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "drew@badsoft.co.uk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Width           =   1710
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.badsoft.co.uk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   3480
      Width           =   1590
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   480
      Picture         =   "frmAbout.frx":0000
      Top             =   3075
      Width           =   2190
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   7395
      Picture         =   "frmAbout.frx":6762
      Top             =   4515
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   360
      Picture         =   "frmAbout.frx":6D34
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblComments 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "comments"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   4935
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":7306
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   3960
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "©2000 - 2001 BadSoft, all rights reserved."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4350
      TabIndex        =   3
      Top             =   1080
      Width           =   3090
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6915
      TabIndex        =   1
      Top             =   840
      Width           =   525
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4455
      Left            =   360
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Load()
    lblTitle = App.ProductName
    lblVersion = "Version " & App.Major & "." & App.Minor & App.Revision
    lblCopyright = App.LegalCopyright
    lblComments.Caption = App.Comments
End Sub

Private Sub cmdOkay_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    'Set the window position to topmost
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
