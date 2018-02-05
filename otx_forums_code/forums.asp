<!--#include file="header.inc"-->
<script>current_screen = "forums";</script>
<%

' Declare variables
Dim objConn
Dim objRS
Dim strSQL

' Open the connection
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Driver=mySQL;Server=localhost;Port=2214;Option=0;Socket=;Stmt=;Database=forums_database;Uid=forums_dbuser;Pwd=forums_dbuser;"

' Retrieve a list of all forums
Set objRS = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM OMNYTEX_ForumsStats"
objRS.Open strSQL, objConn, 3, 3

TAAF_Topics = objRS("StatsTotalTopics")
TAAF_Posts  = objRS("StatsTotalPosts")

' Retrieve value of LastVisit date/time cookie
LastVisit   = Trim(UCase(Request.Cookies("LastVisit")))
If LastVisit = "" Then
 LastVisit = "Never"
 LastVisitDT = #01/01/1000#
Else
 LastVisitDT = LastVisit
End If
' Get the current date/time
CurrentDT   = Now


' Header
Response.Write "<table width=100% border=0 cellpadding=0 cellspacing=0>" & vbCR
Response.Write "<tr>" & vbCR
Response.Write "<td class=cssForumsNormal>&nbsp;</td>" & vbCR
Response.Write "<td align=right class=cssForumsTitle>Omnytex Message Forums</td>" & vbCR
Response.Write "</tr>" & vbCR
Response.Write "<tr><td colspan=2 align=right><hr width=100% height=1 color=#000000></td></tr>" & vbCR
Response.Write "<tr><td class=cssForumsNormal>Your last visit: " & LastVisit & "</td>" & vbCR
Response.Write "<td align=right class=cssForumsNormal>" & vbCR
Response.Write "<strike>New Topic</strike>&nbsp;|&nbsp;" & vbCR
Response.Write "<strike>Reply</strike>&nbsp;|&nbsp;" & vbCR
Response.Write "<i>Forums List</i></a>&nbsp;|&nbsp;" & vbCR
Response.Write "<strike>Topics List</strike>" & vbCR
Response.Write "</td>" & vbCR
Response.Write "</tr>" & vbCR
Response.Write "</table>" & vbCR
Response.Write "<br>" & vbCR

' Retrieve a list of all forums
objRS.Close
strSQL = "SELECT * FROM OMNYTEX_ForumsForums, OMNYTEX_ForumsCategories WHERE OMNYTEX_ForumsForums.CategoryKey = OMNYTEX_ForumsCategories.CategoriesKey ORDER BY OMNYTEX_ForumsCategories.CategoryName"
objRS.Open strSQL, objConn, 3, 3

' Create our output
lastCategory = ""
If objRS.EOF <> True Then
 Response.Write "<table width=100% cellpadding=4 cellspacing=2>" & vbCR
 Response.Write "<tr bgcolor=#003366>" & vbCR
 Response.Write "<td align=center class=cssForumsHeader></td>" & vbCR
 Response.Write "<td align=left class=cssForumsHeader>Forum</td>" & vbCR
 Response.Write "<td align=center class=cssForumsHeader>Topics</td>" & vbCR
 Response.Write "<td align=center class=cssForumsHeader>Posts</td>" & vbCR
 Response.Write "<td align=center class=cssForumsHeader>Last Post</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Do Until objRS.EOF = True
  ' Do this when a new category is encountered
  If objRS("CategoryName") <> lastCategory Then
   Response.Write "<tr bgcolor=#006699><td colspan=5 class=cssForumsCategories>" & objRS("CategoryName") & "</td></tr>" & vbCR
   lastCategory = objRS("CategoryName")
  End If
  Response.Write "<tr>" & vbCR
  Response.Write "<td align=center bgcolor=#ffffff class=cssForumsSmall>"
  If DateDiff("s", LastVisitDT, objRS("ForumLastPost")) > 0 Then
   Response.Write "<img src=images/NewActivity.gif>"
  Else
   Response.Write "<img src=images/spacer.gif width=16 height=16>"
  End If
  Response.Write "</td>" & vbCR
  Response.Write "<td align=left   bgcolor=#f7f7f7 class=cssForums><a href='topicslist.asp?forumKey=" & objRS("ForumsKey") & "&forumName=" & Server.URLEncode(objRS("ForumName")) & "'>" & objRS("ForumName") & "</a><br><span class=cssForumsSmall>"  & objRS("ForumDescription") & "</span></td>" & vbCR
  Response.Write "<td align=center bgcolor=#dedfdf class=cssForumsSmall>" & objRS("ForumTopics")   & "</td>" & vbCR
  Response.Write "<td align=center bgcolor=#f7f7f7 class=cssForumsSmall>" & objRS("ForumPosts")    & "</td>" & vbCR
  Response.Write "<td align=center bgcolor=#dedfdf class=cssForumsSmall>" & objRS("ForumLastPost") & "</td>" & vbCR
  Response.Write "</tr>" & vbCR
  objRS.MoveNext
 Loop
 Response.Write "<tr bgcolor=#003366>" & vbCR
 Response.Write "<td align=center class=cssForumsHeaderSmall colspan=6>" & vbCR
 Response.Write "Totals across ALL forums - Topics: " & TAAF_Topics & ", Posts: " & TAAF_Posts
 Response.Write "</td></tr>" & vbCR

 Response.Write "<tr>" & vbCR
 Response.Write "<td align=center class=cssForumsHeaderSmall><img src=images/NewActivity.gif></td>" & vbCR
 Response.Write "<td align=left colspan=4 class=cssForumsNormal style=color:#000000>Means there has been new activity since your last visit</td>" & vbCR
 Response.Write "</tr>" & vbCR

 Response.Write "</table>" & vbCR
Else
 Response.Write "No Forums" & vbCR
End If

' Clean up
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing

' Write out a cookie with the current date and time with
Response.Cookies("LastVisit") = Now
Response.Cookies("LastVisit").Expires = Now + 1000

%>

<!--#include file="footer.inc"-->

