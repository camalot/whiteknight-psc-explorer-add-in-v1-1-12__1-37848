VERSION 5.00
Begin VB.Form frmCode 
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3360
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' member variable for Text property
Private m_Text As String
' member variable for FilePath property
Private m_FilePath As String


Private Sub Form_Initialize()
  m_Text = ""
  m_FilePath = ""
End Sub


' The FilePath property

Property Get FilePath() As String
  FilePath = m_FilePath
End Property

Property Let FilePath(ByVal newValue As String)
  m_FilePath = newValue
  Dim FSO As FileSystemObject
  Dim fsoTextStream As TextStream
  Set FSO = New FileSystemObject

  If m_FilePath <> "" Then
    If FSO.FileExists(m_FilePath) Then
      Set fsoTextStream = FSO.OpenTextFile(m_FilePath)
      txtCode = fsoTextStream.ReadAll
      'This removes chr(13) so we can relpace all chr(10) with vbcrlf's
      txtCode = Replace$(txtCode, Chr(13), "")
      txtCode = Replace$(Replace$(Replace$(txtCode, Chr(10), vbCrLf), _
          "<xmp>", ""), "</xmp>", "")
      fsoTextStream.Close
    End If
  End If
  Set FSO = Nothing
  Set fsoTextStream = Nothing
End Property


' The Text property

Property Get Text() As String
  Text = m_Text
End Property

Property Let Text(ByVal newValue As String)
  m_Text = newValue
  txtCode = m_Text
  txtCode = Replace$(txtCode, Chr(13), "")
  txtCode = Replace$(txtCode, Chr(10), vbCrLf)
  'this will replace the spaces at the begining.
  txtCode = Replace$(txtCode, "  ", "")
End Property


Private Sub Form_Resize()
  On Error Resume Next
  txtCode.Move 2, 2, ScaleWidth - 4, ScaleHeight - 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim FSO As FileSystemObject
  Set FSO = New FileSystemObject
  If m_FilePath <> "" Then
    If FSO.FileExists(m_FilePath) Then FSO.DeleteFile m_FilePath, True
  End If
  Set FSO = Nothing
End Sub
