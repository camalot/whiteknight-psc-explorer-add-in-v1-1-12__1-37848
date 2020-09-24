VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "About PSC Explorer"
   ClientHeight    =   3840
   ClientLeft      =   2355
   ClientTop       =   1905
   ClientWidth     =   5670
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "frmAbout.frx":058A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   4320
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblDesc 
      Height          =   735
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblWarn 
      Caption         =   $"frmAbout.frx":0A19
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Copyright © 2002 Camalot Designs"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmAbout.frx":0B3A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "PSC Explorer"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   6
      X2              =   376.933
      Y1              =   163
      Y2              =   163
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   7
      X2              =   376.933
      Y1              =   164
      Y2              =   164
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  lblDesc.Caption = App.FileDescription
End Sub

Private Sub Label1_Click()
  LaunchURL ("http://camalotdesigns.com")
End Sub
