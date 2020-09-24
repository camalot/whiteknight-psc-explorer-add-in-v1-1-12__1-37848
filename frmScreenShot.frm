VERSION 5.00
Begin VB.Form frmScreenShot 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmScreenShot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgSShot 
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' member variable for ImagePath property
Private m_ImagePath As String

' The ImagePath property

Property Get ImagePath() As String
  ImagePath = m_ImagePath
End Property

Property Let ImagePath(ByVal newValue As String)
  m_ImagePath = newValue
End Property




Private Sub Form_Unload(Cancel As Integer)
  'Once the user closes this window, we will delete the image
  Dim FSO As FileSystemObject
  Set FSO = New FileSystemObject
  FSO.DeleteFile m_ImagePath, True
  Set FSO = Nothing
End Sub
