Attribute VB_Name = "modMain"
Option Explicit

'Used to download the screenshot
Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Private Const ERROR_SUCCESS As Long = 0

'Used to launch the web browser
Private Declare Function ShellExecute Lib "shell32" _
    Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

'URL To PSC :)
Public Const URL As String = "www.pscode.com"
'URL to the screenshots
Public Const IMG_URL As String = "http://www.exhedra.com/upload/ScreenShots/"

'For the newest submissions
Public Const PAGE As String = "/vb/linktous/ScrollingCode.asp?lngWId="
'For the Submission "Home" page
Public Const CODE_PAGE As String = "/vb/scripts/ShowCode.asp?txtCodeId=XXCODEIDXX&lngWId=XXLNGWIDXX"
'The search page
Public Const SEARCH_PAGE As String = "/vb/scripts/BrowseCategoryOrSearchResults.asp?chkCode3rdPartyReview=XX3RDPARTYXX&optSort=XXSORTXX&cmSearch=Search&blnResetAllVariables=TRUE&txtMaxNumberOfEntriesPerPage=XXMAXPERPAGEXX&txtCriteria=XXSEARCHXX&chkCodeTypeZip=XXZIPXX&chkCodeTypeText=XXTEXTCODEXX&chkCodeTypeArticle=XXARTICLEXX&chkCodeDifficulty=XXDIFFXX&cmdGoToPage=XXPAGEXX&lngWId="
'To Download Zip File
Public Const ZIP_PAGE As String = "/vb/scripts/ShowZip.asp?lngWId=XXWORLDIDXX&lngCodeId=XXCODEIDXX&strZipAccessCode=XXZIPACCESSCODEXX"
'Get the Code as text
Public Const CODETEXT_PAGE As String = "/vb/scripts/ShowCodeAsText.asp?txtCodeId=XXCODEIDXX&lngWId=XXWORLDIDXX"
'URL for Code of the month
Public Const CODEMONTH_PAGE As String = "/vb/contest/ContestAndLeaderBoard.asp?lngWId=XXWORLDIDXX"

'Stores Data for Difficulty Level (For Search)
Public strDiffLevel As String
'Stores Data for Results Per Page(For Search) Default = 50
Public intMaxPerPage As Integer
'Stores Data for Search For Zip Files (For Search) Default = "on"
Public strZipFiles As String
'Stores Data for Search For Text Code (For Search) Default = "on"
Public strCodeText As String
'Stores Data for Search For Articles (For Search) Default = "on"
Public strArticles As String
'Stores Data for Search For 3rd Party Reviews(For Search) Default = "on"
Public str3rdParty As String
'Stores Data for Sorting (For Search) Default = "DateDescending"
Public strSort As String

'Use a Proxy Server?
Public bUseProxy As Boolean
'Proxy Address
Public strProxyServer As String
'Proxy Prot
Public intProxyPort As Integer

'Stores the Time Out Delay (Seconds) Default = 30
Public intTimeOut As Integer


'These are used in the ParseDataAdvanced to see what type of submission it is
Private Const PSC_CODE As String = "/vb/images/vbicon.gif"
Private Const PSC_ZIP As String = "/vb/scripts/images/CodeZip_small.gif"
Private Const PSC_ARTICLE As String = "/vb/scripts/images/ArticleCopyAndPaste_small.gif"
Private Const PSC_ARTICLEZIP As String = "/vb/scripts/images/ArticleZip_small.gif"
Private Const PSC_CCODE As String = "/vb/images/CIcon_small.gif"
Private Const PSC_REVIEW As String = "/vb/scripts/images/3rdPartyReview.gif"


Public Type PSCData
  PSC_TITLE As String  'The Title of the submission
  PSC_ID As String  'The Code ID
  PSC_IMG As String  'File name for the screenshot
  PSC_WORLD As String  'What World (1-13) (11 & 12 are nothing?)
  PSC_AUTHOR As String  'Authors Name
  PSC_DATE As String  'Date Posted (MM/DD/YY for newest, MM/DD/YYYY HH:MM:SS AP for search)
  PSC_CODECOMPAT As String  'Code Compatibility (N/A for Newest)
  PSC_LEVEL As String  'Code Level (N/A for Newest)
  PSC_VIEWS As String  'Total Views (N/A for Newest)
  PSC_RATE As String  'Ratting (N/A for Newest)
  PSC_DESCRIPTION As String  'Description (N/A for Newest)
  PSC_TYPE As String  'Type - Code, Zip, Article (N/A for Newest)
  PSC_ZIPACCESSCODE As String  'ZipAccessCode, for quick download (N/A for Newest)
  PSC_EXECELLENTVOTES As String 'Number of "Execellent" Votes for Top Code
  PSC_TCRATING As String 'Top Code Rating
End Type

Public booStopSearch As Boolean
Public lngTotalResults As Long


Public Function ParseData(Data As String, PSCInfo() As PSCData)
  Dim lngCount As Long, lngTotal As Long
  Dim lStart As Long, lEnd As Long
  Dim tlStart As Long, tlEnd As Long
  Dim str As String, str2 As String
  Dim x As Long

  'This function is for the "Newest Submissions"

  'The First Submission
  str = "<a href=""/vb/scripts/ShowCode.asp?"
  'Set the start point
  lStart = InStr(1, Data, str, vbTextCompare)
  'Loop until there are no more
  Do While lStart > 1
    'Did the user cancel the search
    If booStopSearch Then Exit Function
    'Set the end point
    lEnd = InStr(lStart + 1, Data, "</a>", vbTextCompare) + 4
    'Set the Array
    ReDim Preserve PSCInfo(lngCount) As PSCData

    'Check For A Screenshot
    If LCase$(Mid(Data, lStart - 4, 4)) = "</a>" Then
      'This one has a screenshot so store "IMAGE" in there so we can replace it later
      PSCInfo(lngCount).PSC_IMG = "IMAGE"
    Else
      'Nope, no screenshot here
      PSCInfo(lngCount).PSC_IMG = ""
    End If

    'Get The Item ID
    'Set the tempStart Position
    tlStart = InStr(1, Mid$(Data, lStart, lEnd - lStart), "txtCodeId=", vbTextCompare) + 10
    'Set the TempEnd Position
    tlEnd = InStr(tlStart, Mid$(Data, lStart, lEnd - lStart), "&", vbTextCompare)
    'store the ID
    str2 = Mid$(Data, lStart, lEnd - lStart)
    'Trim Some more off ;)
    PSCInfo(lngCount).PSC_ID = Mid$(str2, tlStart, tlEnd - tlStart)
    'Here we are just going to put an emptystring in the Zip Acces Code and Code In the Type
    PSCInfo(lngCount).PSC_TYPE = "Code"
    PSCInfo(lngCount).PSC_ZIPACCESSCODE = ""

    'Get The World ID
    'Set the tempStart Position
    tlStart = InStr(1, Mid$(Data, lStart, lEnd - lStart), "lngWId=", vbTextCompare) + 7
    'Set the TempEnd Position
    tlEnd = InStr(tlStart, Mid$(Data, lStart, lEnd - lStart), """target=", vbTextCompare)
    'store the World ID
    str2 = Mid$(Data, lStart, lEnd - lStart)
    'Trim Some more off ;)
    PSCInfo(lngCount).PSC_WORLD = Mid$(str2, tlStart, tlEnd - tlStart)

    'Get Title
    'Set the tempStart Position
    tlStart = InStr(1, Mid$(Data, lStart, lEnd - lStart), ">", vbTextCompare) + 1
    'Set the TempEnd Position
    tlEnd = InStr(tlStart, Mid$(Data, lStart, lEnd - lStart), "<", vbTextCompare)
    'store the World ID
    str2 = Mid$(Data, lStart, lEnd - lStart)
    'Trim Some more off ;)
    PSCInfo(lngCount).PSC_TITLE = Mid$(str2, tlStart, tlEnd - tlStart)


    'Next Item
    lStart = InStr(lStart + 1, Data, str, vbTextCompare)
    lngCount = lngCount + 1
  Loop
  lngTotal = lngCount


  'Set Start String for the next Section of Data
  str = "<BR>By"
  'Set the Start Position
  lStart = InStr(1, Data, str, vbTextCompare) + Len(str)
  'Reset The Counter
  lngCount = 0
  'Loop untill there are no more
  Do While lStart > Len(str)

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Get Author
    'Set the End Position
    lEnd = InStr(lStart + 1, Data, "&nbsp;on&nbsp;", vbTextCompare)
    PSCInfo(lngCount).PSC_AUTHOR = Mid$(Data, lStart, lEnd - lStart)

    'Get Date
    'Set the TempStart
    tlStart = lEnd + 14
    'Set the TempEnd
    tlEnd = InStr(tlStart, Data, "</b>", vbTextCompare)
    'Store the Date
    PSCInfo(lngCount).PSC_DATE = Mid$(Data, tlStart, tlEnd - tlStart) & "/" & Right$(Date, 2)

    'Next Item
    lStart = InStr(lStart, Data, str, vbTextCompare) + Len(str)
    lngCount = lngCount + 1
  Loop

  'Find the Screenshot Filename for Submissions that have ScreenShots
  str = "<b><a href=""http://www.exhedra.com/upload/ScreenShots/"
  'Set Start Position
  lStart = InStr(1, Data, str, vbTextCompare)
  'Reset Counter
  lngCount = 0
  'Loop until there are no more
  Do While lStart > 1

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Set End Position
    lEnd = InStr(lStart + 1, Data, "target=", vbTextCompare)
    'Did we say that this item had a screen shot earlier?
    If PSCInfo(lngCount).PSC_IMG = "IMAGE" Then
      'Yes, It has a screenshot
      'Save the File Name for use later
      PSCInfo(lngCount).PSC_IMG = Trim$(Replace$(Mid(Data, lStart + Len(str), lEnd - (lStart + Len(str))), """", ""))
      'Set the Next start position
      lStart = InStr(lStart + 1, Data, str, vbTextCompare)
    Else
      'Do Nothing
    End If
    lngCount = lngCount + 1

  Loop
End Function

'This Function is similar to the ParseData Function but it is used for the search and returns more information
Public Function ParseDataAdvanced(Data As String, PSCInfo() As PSCData) As Boolean
  Dim lngCount As Long, lngTotal As Long
  Dim lStart As Long, lEnd As Long
  Dim tlStart As Long, tlEnd As Long
  Dim str As String, str2 As String
  Dim x As Long
  Dim RateCount As Integer
  Dim strRate As String


  On Error Resume Next
  'Did the user cancel the search
  If booStopSearch Then Exit Function


  'Did the search results from PSC Find NO Matches?
  If InStr(1, Data, "No records found matching your search.", vbTextCompare) Then
    'Yep, no matches were found, we are done here
    ParseDataAdvanced = False
    'Lets let the user know
    MsgBox "No records found matching your search.", vbApplicationModal + vbDefaultButton1 + vbInformation, "No Matches"
    Exit Function
  End If

  str = "</a><a href=/vb/scripts/showcode.asp"
  lStart = InStr(1, Data, str, vbTextCompare)
  Do While lStart > 1

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    lEnd = InStr(lStart + 1, Data, "</a>", vbTextCompare) + 4
    ReDim Preserve PSCInfo(lngCount) As PSCData

    'Get The Item ID
    tlStart = InStr(1, Mid$(Data, lStart, lEnd - lStart), "txtCodeId=", vbTextCompare) + 10
    tlEnd = InStr(tlStart, Mid$(Data, lStart, lEnd - lStart), "&", vbTextCompare)
    str2 = Mid$(Data, lStart, lEnd - lStart)
    PSCInfo(lngCount).PSC_ID = Mid$(str2, tlStart, tlEnd - tlStart)

    'Get The World ID
    tlStart = InStr(1, Mid$(Data, lStart, lEnd - lStart), "lngWId=", vbTextCompare) + 7
    tlEnd = InStr(tlStart, Mid$(Data, lStart, lEnd - lStart), ">", vbTextCompare)
    str2 = Mid$(Data, lStart, lEnd - lStart)
    PSCInfo(lngCount).PSC_WORLD = Mid$(str2, tlStart, tlEnd - tlStart)

    'Get Title
    tlStart = InStr(5, str2, ">", vbTextCompare) + 1
    tlEnd = InStr(tlStart, str2, "<", vbTextCompare)
    str2 = Mid$(str2, tlStart, tlEnd - tlStart)
    PSCInfo(lngCount).PSC_TITLE = str2

    lStart = InStr(lStart + 1, Data, str, vbTextCompare)
    lngCount = lngCount + 1
    ParseDataAdvanced = True
  Loop

  str = "<!--level-->"
  lStart = InStr(1, Data, str, vbTextCompare) + Len(str)
  lngCount = 0
  Do While lStart > Len(str)

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Get Level
    lEnd = InStr(lStart + 1, Data, "&nbsp;/<BR>", vbTextCompare)
    PSCInfo(lngCount).PSC_LEVEL = Mid$(Data, lStart, lEnd - lStart)

    'Get Author
    tlStart = lEnd + 11
    tlEnd = InStr(tlStart, Data, "</TD>", vbTextCompare)
    str2 = Mid$(Data, tlStart, tlEnd - tlStart)
    If InStr(1, str2, "<a href", vbTextCompare) Then
      'Remove the url, if there is one
      str2 = Mid$(str2, InStr(1, str2, ">", vbTextCompare) + 1)
      str2 = Mid$(str2, 1, InStr(1, str2, "<", vbTextCompare) - 1)
    End If
    PSCInfo(lngCount).PSC_AUTHOR = str2

    'Next Item
    lStart = InStr(lStart, Data, str, vbTextCompare) + Len(str)
    lngCount = lngCount + 1
  Loop

  '<!--views/date submitted-->
  str = "<!--views/date submitted-->"
  lStart = InStr(1, Data, str, vbTextCompare) + Len(str)
  lngCount = 0
  Do While lStart > Len(str)

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Get Views
    lEnd = InStr(lStart + 1, Data, "since<br>", vbTextCompare)
    PSCInfo(lngCount).PSC_VIEWS = StripHTML(Mid$(Data, lStart, lEnd - lStart))

    'Get Date
    tlStart = lEnd + 9
    tlEnd = InStr(tlStart, Data, "</TD>", vbTextCompare)
    str2 = Mid$(Data, tlStart, tlEnd - tlStart)
    PSCInfo(lngCount).PSC_DATE = str2

    'Next Item
    lStart = InStr(lStart, Data, str, vbTextCompare) + Len(str)
    lngCount = lngCount + 1
  Loop

  '<!description><FONT Size=2 >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  str = "<!description><FONT Size=2 >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
  lStart = InStr(1, Data, str, vbTextCompare) + Len(str)
  lngCount = 0
  Do While lStart > Len(str)

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Get Description
    lEnd = InStr(lStart + 1, Data, "<HR></TD></TR>", vbTextCompare)
    PSCInfo(lngCount).PSC_DESCRIPTION = StripHTML(Mid$(Data, lStart, lEnd - lStart))
    If InStr(1, PSCInfo(lngCount).PSC_DESCRIPTION, "(ScreenShot)", vbTextCompare) > 0 Then
      'Remove (ScreenShot) from description
      PSCInfo(lngCount).PSC_DESCRIPTION = Replace$(PSCInfo(lngCount).PSC_DESCRIPTION, "(ScreenShot)", "", , , vbTextCompare)
      PSCInfo(lngCount).PSC_IMG = "IMAGE"
    End If

    'Next Item
    lStart = InStr(lStart, Data, str, vbTextCompare) + Len(str)
    lngCount = lngCount + 1
  Loop

  '<!--user rating--><TD align=center><FONT Size=1 ><center>
  str = "<!--user rating-->"
  lStart = InStr(1, Replace$(Replace$(Data, vbTab, ""), vbCrLf, ""), str, vbTextCompare) + Len(str)
  lngCount = 0


  Do While lStart > Len(str)

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Get User Ratting
    lEnd = InStr(lStart + 1, Replace$(Replace$(Data, vbTab, ""), vbCrLf, ""), "<BR></center>", vbTextCompare)

    'Now that we have all the rating data, lets get rid of all the tabs and CrLf's
    str2 = Mid$(Replace(Replace$(Data, vbTab, ""), vbCrLf, ""), lStart, lEnd - lStart)
    str2 = Mid$(str2, InStr(1, str2, "<center>", vbTextCompare) + 8)
    PSCInfo(lngCount).PSC_RATE = str2

    'Check to see if it has been rated
    If PSCInfo(lngCount).PSC_RATE <> "Unrated" Then
      'Yes, it has been rated
      RateCount = 0
      'This loops through all the "full rating images" and counts them up
      tlStart = InStr(1, PSCInfo(lngCount).PSC_RATE, "<img src=""/vb/scripts/voting/images/RatingSmall", vbTextCompare)
      Do While tlStart > 0
        RateCount = RateCount + 1
        tlStart = InStr(tlStart + 1, PSCInfo(lngCount).PSC_RATE, "<img src=""/vb/scripts/voting/images/RatingSmall", vbTextCompare)
      Loop
      'This looks to see if there is a "half rating image"
      If InStr(1, PSCInfo(lngCount).PSC_RATE, "<img src=""/vb/scripts/voting/images/RatingHalf", vbTextCompare) Then
        strRate = RateCount & ".50"
      Else
        strRate = RateCount & ".00"
      End If
      'We now have the rating
      PSCInfo(lngCount).PSC_RATE = strRate
    End If
    'Next Item
    lStart = InStr(lStart, Replace(Replace(Data, vbTab, ""), vbCrLf, ""), str, vbTextCompare) + Len(str)
    lngCount = lngCount + 1
  Loop

  'This is similar to the screenshort part of "ParseData" function
  'Get ScreenShots
  '<a href="/upload/ScreenShots/
  str = "<a href=""/upload/ScreenShots/"
  lStart = InStr(1, Replace$(Data, vbTab, ""), str, vbTextCompare) + Len(str)
  lngCount = 0
  Do While lStart > Len(str)

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Get ScreenShot
    lEnd = InStr(lStart + 1, Replace$(Data, vbTab, ""), """ target=""_new""", vbTextCompare)
    If PSCInfo(lngCount).PSC_IMG = "IMAGE" Then
      PSCInfo(lngCount).PSC_IMG = Trim$(Mid$(Replace$(Data, vbTab, ""), lStart, lEnd - lStart))
      lStart = InStr(lStart, Replace$(Data, vbTab, ""), str, vbTextCompare) + Len(str)
    Else
      'Do Nothing
    End If
    'Next Item

    lngCount = lngCount + 1
  Loop


  'This looks to see if the submission is a Zip, Text, or article.
  'Get Submission Type
  str = "<img align=""left""  border=0 src="""
  lStart = InStr(1, Data, str, vbTextCompare) + Len(str)
  lngCount = 0
  Do While lStart > Len(str)

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Get Type
    lEnd = InStr(lStart + 1, Data, """ alt=", vbTextCompare)
    If Trim$(Mid$(Data, lStart, lEnd - lStart)) = PSC_CODE Or Trim$(Mid$(Data, lStart, lEnd - lStart)) = PSC_CCODE Then
      PSCInfo(lngCount).PSC_TYPE = "Code"
    ElseIf Trim$(Mid$(Data, lStart, lEnd - lStart)) = PSC_ZIP Then
      PSCInfo(lngCount).PSC_TYPE = "Zip"
      'Here we are going to get the ZipAccessCode.
      tlEnd = lStart - Len(str)
      tlStart = InStrRev(Data, "strZipAccessCode=", tlEnd, vbTextCompare) + 17
      str2 = Mid$(Data, tlStart, (tlEnd - tlStart) - 2)
      PSCInfo(lngCount).PSC_ZIPACCESSCODE = str2
    ElseIf Trim$(Mid$(Data, lStart, lEnd - lStart)) = PSC_REVIEW Then
      PSCInfo(lngCount).PSC_TYPE = "Review"
    ElseIf Trim$(Mid$(Data, lStart, lEnd - lStart)) = PSC_ARTICLE Then
      PSCInfo(lngCount).PSC_TYPE = "Article"
    ElseIf Trim$(Mid$(Data, lStart, lEnd - lStart)) = PSC_ARTICLEZIP Then
      PSCInfo(lngCount).PSC_TYPE = "Article"
    End If
    'Next Item
    lStart = InStr(lStart, Data, str, vbTextCompare) + Len(str)
    lngCount = lngCount + 1
  Loop


  'This finds and gets the Code Compatitability
  '<!--code compat-->
  str = "<!--code compat-->"
  lStart = InStr(1, Data, str, vbTextCompare) + Len(str)
  lngCount = 0
  Do While lStart > Len(str)

    'Did the user cancel the search
    If booStopSearch Then Exit Function

    'Get Code Compat
    lEnd = InStr(lStart + 1, Data, "</TD><!--level/author-->", vbTextCompare)
    str2 = Trim$(Mid$(Data, lStart, lEnd - lStart))
    PSCInfo(lngCount).PSC_CODECOMPAT = StripHTML(str2)
    lStart = InStr(lStart, Data, str, vbTextCompare) + Len(str)
    lngCount = lngCount + 1
  Loop

  'Find the Total Nuber of Entries in Search
  'lngTotalResults
  'Entries
  str = "Entries "
  lStart = InStr(1, Data, str, vbTextCompare) + Len(str)
  If lStart > Len(str) Then
    'Get Total Found
    lEnd = InStr(lStart + 1, Data, "</font>", vbTextCompare)

    str2 = Trim$(Mid$(Data, lStart, lEnd - lStart))
    tlStart = InStr(1, str2, "of ", vbTextCompare) + 3
    tlEnd = InStr(tlStart, str2, " found", vbTextCompare)
    str2 = Mid$(str2, tlStart, tlEnd - tlStart)
    lngTotalResults = CLng(str2)
  End If



  If lngCount = 0 Then ParseDataAdvanced = False
End Function

'This is to download the Screenshot and save it to the local computer
Public Function DownloadToFile(ByVal sURL As String, ByVal sLocalFile As String) As Boolean
  Dim lngRetVal As Long
  DownloadToFile = URLDownloadToFile(0&, sURL, sLocalFile, 0&, 0&) = ERROR_SUCCESS
End Function

Public Function StripHTML(ByVal strData As String) As String
  Dim lStart As Long, lEnd As Long
  Dim strTemp As String

  'This Function Just Removes ALL HTML tags <BLAH>
  lStart = InStr(1, strData, "<", vbTextCompare)
  Do While lStart > 0
    lEnd = InStr(lStart, strData, ">", vbTextCompare) + 1
    strTemp = Mid(strData, lStart, lEnd - lStart)
    strData = Replace(strData, strTemp, "")

    lStart = InStr(lStart, strData, "<", vbTextCompare)
  Loop
  strData = Replace(strData, "&nbsp;", " ", , , vbTextCompare)
  strData = Replace(strData, "&amp;", "&", , , vbTextCompare)
  StripHTML = strData
End Function


'I'll Give you 1 guess what this does :)
Public Sub LaunchURL(strURL As String)
  ShellExecute 0, "Open", strURL, vbNullString, App.Path, SW_SHOWNORMAL

End Sub
