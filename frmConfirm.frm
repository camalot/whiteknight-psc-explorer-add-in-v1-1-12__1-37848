VERSION 5.00
Begin VB.Form frmConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Before You Download..."
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDisagree 
      Caption         =   "No Thanks"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton mcdAgree 
      Caption         =   "I Agree"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdMinnow 
      Caption         =   "Get Minnow's Project Scanner"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FOR YOUR OWN SAFETY, PLEASE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label lblTerms 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   6255
   End
   Begin VB.Label lblWarn 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' member variable for TheURL property
Private m_TheURL As String

' The TheURL property

Property Get TheURL() As String
  TheURL = m_TheURL
End Property

Property Let TheURL(ByVal newValue As String)
  m_TheURL = newValue
End Property




Private Sub cmdDisagree_Click()
  Unload Me
End Sub

Private Sub cmdMinnow_Click()
  LaunchURL ("http://www.pscode.com/xq/ASP/txtCodeId.22222/lngWId.1/qx/vb/scripts/ShowCode.htm")
End Sub

Private Sub Form_Load()
  lblWarn = "1)Re-scan downloaded files using your personal virus checker before using it." & _
      vbCrLf & "2)NEVER, EVER run compiled files (.exe's, .ocx's, .dll's etc.)--only run source code. " & _
      vbCrLf & "3)Scan the source code with Minnow's Project Scanner"
  lblTerms = "Terms of Agreement:   " & vbCrLf & _
      "By using this code, you agree to the following terms..." & vbCrLf & _
      "1) You may use this code in your own programs (and may compile it into a program and distribute it in compiled format for langauges that allow it) freely and with no charge." & vbCrLf & _
      "2) You MAY NOT redistribute this code (for example to a web site) without written permission from the original author. Failure to do so is a violation of copyright laws." & vbCrLf & _
      "3) You may link to this code from another website, but ONLY if it is not wrapped in a frame." & vbCrLf & _
      "4) You will abide by any additional copyright restrictions which the author may have placed in the code or code's description."
End Sub

Private Sub mcdAgree_Click()
  LaunchURL m_TheURL
  Unload Me
End Sub
