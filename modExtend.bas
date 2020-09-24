Attribute VB_Name = "modExtend"
Option Explicit



'This module is mainly for getting the data from the PSC submission page. Like
'The description mainly :)

'Also for getting the copy & Paste code.


'Start of Description
'<td width="419" colspan="3" bgcolor="white">    <font size="2" color="#000000">  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
'<td width=""419"" colspan=""3"" bgcolor=""white"">    <font size=""2"" color=""#000000"">
'End Of Description
'<br>    </font>  </td></tr></table>

Public Function ParseDescription(strData As String) As String
  Dim lStart As Long
  Dim lEnd As Long
  Dim str1 As String, str2 As String
  str1 = "<td width=""419"" colspan=""3"" bgcolor=""white"">    <font size=""2"" color=""#000000"">"
  str2 = "<br>    </font>  </td></tr></table>"
  lStart = InStr(1, strData, str1, vbTextCompare) + Len(str1)
  'See if we found the start position
  If lStart > Len(str1) Then
    'set the end position
    lEnd = InStr(lStart + 1, strData, str2, vbTextCompare)

    'see if we found the end position
    If lEnd > 0 Then
      'return the description
      ParseDescription = StripHTML(Mid$(strData, lStart, lEnd - lStart))
    End If
  End If
End Function


Public Function ParseTopCode(strData As String, PSCInfo() As PSCData) As Boolean
  Dim lStart As Long, lEnd As Long
  Dim str1 As String, str2 As String
  Dim tlStart As Long, tlEnd As String
  Dim lngCount As Long

  str1 = "<a href=/vb/default.asp?lngCId="
  lStart = InStr(1, strData, str1, vbTextCompare) + Len(str1)
  'MsgBox strData
  Do While lStart > Len(str1)

    'Set the end position
    lEnd = InStr(lStart + 1, strData, "</td>", vbTextCompare)

    str1 = Replace$(Replace$(Replace$(Mid$(strData, lStart, lEnd - lStart), _
      vbTab, ""), Chr(10), ""), Chr(13), "")
    'MsgBox "Str1=" & str1
    'The text we will get will contain the CodeID, Title, Author
    'Get the CodeID
    'Set the end position of the CodeID


    tlEnd = InStr(1, str1, "&lngWId", vbTextCompare)
    'we now have the Code ID
    str2 = Mid$(str1, 1, tlEnd - 1)
    'MsgBox "ID=" & str2
    ReDim Preserve PSCInfo(lngCount) As PSCData

    PSCInfo(lngCount).PSC_ID = str2

    'Get the Title
    tlStart = InStr(1, str1, ">", vbTextCompare) + 1
    If tlStart > 1 Then tlEnd = InStr(tlStart, str1, "</a>", vbTextCompare)
    If tlEnd > 0 Then str2 = Trim$(Mid$(str1, tlStart, tlEnd - tlStart))
    If str2 <> "" Then PSCInfo(lngCount).PSC_TITLE = str2
    'MsgBox "Title=" & str2


    'Get the Author
    tlStart = tlEnd
    tlEnd = Len(str1)
    str2 = Mid$(str1, tlStart)
    str2 = StripHTML(str2)
    PSCInfo(lngCount).PSC_AUTHOR = str2
    'MsgBox "Author=" & str2
    
    'Get the number of Execellent Votes
    
    tlStart = InStr(lStart + 1, strData, "<td valign=top>", vbTextCompare) + 15
    tlEnd = InStr(tlStart + 1, strData, "</td>", vbTextCompare)
    
    str2 = Mid$(strData, tlStart, tlEnd - tlStart)
    str2 = StripHTML(Replace$(Replace$(Replace$(str2, vbTab, ""), Chr(10), ""), _
      Chr(13), ""))
    PSCInfo(lngCount).PSC_EXECELLENTVOTES = Trim$(str2)
    
    'Get Rating Data
    tlStart = InStr(tlStart + 1, strData, "<td valign=top>", vbTextCompare) + 15
    tlEnd = InStr(tlStart + 1, strData, "</td>", vbTextCompare)
    str2 = Mid$(strData, tlStart, tlEnd - tlStart)
    str2 = StripHTML(Replace$(Replace$(Replace$(str2, vbTab, ""), Chr(10), ""), _
      Chr(13), ""))
    PSCInfo(lngCount).PSC_TCRATING = Trim$(Replace$(str2, "from", " from", , , vbTextCompare))
    
    'Get the Date Posted
    tlStart = InStr(tlStart + 1, strData, "<td valign=top>", vbTextCompare) + 15
    tlEnd = InStr(tlStart + 1, strData, "</td>", vbTextCompare)
    str2 = Mid$(strData, tlStart, tlEnd - tlStart)
    str2 = StripHTML(Replace$(Replace$(Replace$(str2, vbTab, ""), Chr(10), ""), _
      Chr(13), ""))
    PSCInfo(lngCount).PSC_DATE = Trim$(str2)
    
    
    PSCInfo(lngCount).PSC_TYPE = "TopCode"
    
    'Next start position
    str1 = "<a href=/vb/default.asp?lngCId="
    lStart = InStr(lStart + 1, strData, str1, vbTextCompare) + Len(str1)
    lngCount = lngCount + 1


  Loop
  
  
  
  If lngCount > 0 Then ParseTopCode = True
End Function
