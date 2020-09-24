VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      Height          =   3735
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox txtTimeOut 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Text            =   "30"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Display In"
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   5415
         Begin VB.OptionButton optSort 
            Caption         =   "Most Popular"
            Height          =   255
            Index           =   3
            Left            =   2880
            TabIndex        =   25
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Oldest Submissions "
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Newest Submissions"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   23
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Alphabetical Order"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Code Difficulty Level"
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   5415
         Begin VB.CheckBox chkLevelAdvanced 
            Caption         =   "Advanced"
            Height          =   255
            Left            =   2880
            TabIndex        =   20
            Top             =   480
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkLevelIntermediate 
            Caption         =   "Intermediate"
            Height          =   255
            Left            =   2880
            TabIndex        =   19
            Top             =   240
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkLevelBeginner 
            Caption         =   "Beginner"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkLevelUnranked 
            Caption         =   "Unranked"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Value           =   1  'Checked
            Width           =   1575
         End
      End
      Begin VB.TextBox txtMaxResults 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "50"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Code Type"
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   5415
         Begin VB.CheckBox chk3rdParty 
            Caption         =   "3rd Party Review"
            Height          =   255
            Left            =   2880
            TabIndex        =   16
            Top             =   480
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkArticles 
            Caption         =   "Articles / Tutorials"
            Height          =   255
            Left            =   2880
            TabIndex        =   15
            Top             =   240
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.CheckBox chkCopyPaste 
            Caption         =   "'Copy-and-paste' source code "
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.CheckBox chkZipFiles 
            Caption         =   "Zip Files"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   1  'Checked
            Width           =   2055
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Time Out:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Max Results Per Page:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame fraProxy 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CheckBox chkProxy 
         Appearance      =   0  'Flat
         Caption         =   "Use Proxy Server"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox txtProxyServer 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtProxyPort 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "8080"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Proxy Server:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkProxy_Click()
  Dim blnEnabled As Boolean
  blnEnabled = chkProxy.Value
  Label1.Enabled = blnEnabled
  Label2.Enabled = blnEnabled
  txtProxyServer.Enabled = blnEnabled
  txtProxyPort.Enabled = blnEnabled
  bUseProxy = blnEnabled
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  'Set max per page
  If IsNumeric(txtMaxResults) Then
    If Val(txtMaxResults) > 50 Or Val(txtMaxResults) < 1 Then
      intMaxPerPage = 50
    Else
      intMaxPerPage = Val(txtMaxResults)
    End If
  Else
    intMaxPerPage = 50
  End If
  'Set the Timeout
  If IsNumeric(txtTimeOut) Then
    If Val(txtTimeOut) > 60 Or Val(txtTimeOut) < 1 Then
      intTimeOut = 30
    Else
      intTimeOut = Val(txtTimeOut)
    End If
  Else
    intTimeOut = 50
  End If

  'Set the DiffLevel
  strDiffLevel = ""
  If chkLevelUnranked.Value Then
    strDiffLevel = "1"
  End If
  If chkLevelBeginner.Value Then
    If strDiffLevel = "" Then
      strDiffLevel = "2"
    Else
      strDiffLevel = strDiffLevel & "%2C+2"
    End If
  End If
  If chkLevelIntermediate.Value Then
    If strDiffLevel = "" Then
      strDiffLevel = "3"
    Else
      strDiffLevel = strDiffLevel & "%2C+3"
    End If
  End If
  If chkLevelAdvanced.Value Then
    If strDiffLevel = "" Then
      strDiffLevel = "4"
    Else
      strDiffLevel = strDiffLevel & "%2C+4"
    End If
  End If

  'Zip File Code
  If chkZipFiles.Value Then
    strZipFiles = "on"
  Else
    strZipFiles = ""
  End If

  'Copy Paste Code
  If chkCopyPaste.Value Then
    strCodeText = "on"
  Else
    strCodeText = ""
  End If

  'Articles / Tutorials
  If chkArticles.Value Then
    strArticles = "on"
  Else
    strArticles = ""
  End If

  '3rd Party
  If chk3rdParty.Value Then
    str3rdParty = "on"
  Else
    str3rdParty = ""
  End If

  If optSort(0).Value Then
    'Alpha
    strSort = "Alphabetical"
  ElseIf optSort(1).Value Then
    'Oldest
    strSort = "DateAscending"
  ElseIf optSort(2).Value Then
    'Newest
    strSort = "DateDescending"
  ElseIf optSort(3).Value Then
    'Most Popular
    strSort = "CountDescending"
  End If

  'Set Proxy
  If bUseProxy Then
    strProxyServer = txtProxyServer
    intProxyPort = CInt(txtProxyPort)
  End If
  Unload Me
End Sub

