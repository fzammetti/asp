<!--#include file="header.inc"-->
<script>current_screen = "";</script>

<%

' Declare variables
Dim objConn
Dim objRS
Dim strSQL

' Get input variables
inCommand    = Trim(Request("Command"))
forumKey     = Trim(Request("forumKey"))
forumName    = Trim(Request("forumName"))
topicKey     = Trim(Request("topicKey"))

' Header
Response.Write "<table width=100% border=0 cellpadding=0 cellspacing=0>" & vbCR
Response.Write "<tr>" & vbCR
Response.Write "<td class=cssForumsNormal>&nbsp;</td>" & vbCR
Response.Write "<td align=right class=cssForumsTitle>Omnytex Message Forums</td>" & vbCR
Response.Write "</tr>" & vbCR
Response.Write "<tr><td colspan=2 align=right><hr width=100% height=1 color=#000000></td></tr>" & vbCR
Response.Write "<tr><td class=cssForumsNormal>&nbsp;</td>" & vbCR
Response.Write "<td align=right class=cssForumsNormal>" & vbCR
Response.Write "<a href='newtopic.asp?Command=showForm&forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>New Topic</a>&nbsp;|&nbsp;" & vbCR
Response.Write "<strike>Reply</strike>&nbsp;|&nbsp;" & vbCR
Response.Write "<a href='forums.asp'>Forums List</a>&nbsp;|&nbsp;" & vbCR
Response.Write "<i>Topics List</i></a>" & vbCR
Response.Write "</td>" & vbCR
Response.Write "</tr>" & vbCR
Response.Write "</table>" & vbCR
Response.Write "<br>" & vbCR

' Open the connection
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Driver=mySQL;Server=localhost;Port=2214;Option=0;Socket=;Stmt=;Database=forums_database;Uid=forums_dbuser;Pwd=forums_dbuser;"

' Retrieve a list of all forums
Set objRS = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM OMNYTEX_ForumsTopics WHERE ForumKey=" & forumKey & " ORDER BY TopicPosted DESC"
objRS.Open strSQL, objConn, 3, 3

' Create our output
lastCategory = ""
If objRS.EOF <> True Then
 TFTF_Topics = 0
 TFTF_Replies = 0
 Response.Write "<table width=100% cellpadding=4 cellspacing=2>" & vbCR
 Response.Write "<tr bgcolor=#003366>" & vbCR
 Response.Write "<td align=center class=cssForumsHeader>&nbsp;</font></td>" & vbCR
 Response.Write "<td align=left class=cssForumsHeader>Topic</td>" & vbCR
 Response.Write "<td align=center class=cssForumsHeader>Topic Starter</td>" & vbCR
 Response.Write "<td align=center class=cssForumsHeader>Replies</td>" & vbCR
 Response.Write "<td align=center class=cssForumsHeader>Last Post</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Do Until objRS.EOF = True
  TFTF_Topics = TFTF_Topics + 1
  TFTF_Replies = TFTF_Replies + objRS("TopicReplies")
  Response.Write "<tr>" & vbCR
  Response.Write "<td align=center bgcolor=#ffffff><img src=images/mficon_" & objRS("TopicIcon") & ".gif></td>" & vbCR
  Response.Write "<td align=left bgcolor=#f7f7f7 class=cssForumsSmall><a href='topic.asp?topicKey=" & objRS("TopicsKey") & "&forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>" & objRS("TopicTitle") & "</a></td>" & vbCR
  Response.Write "<td align=center bgcolor=#dedfdf class=cssForumsSmall>" & objRS("TopicStarter")   & "</td>" & vbCR
  Response.Write "<td align=center bgcolor=#f7f7f7 class=cssForumsSmall>" & objRS("TopicReplies")    & "</td>" & vbCR
  Response.Write "<td align=center bgcolor=#dedfdf class=cssForumsSmall>" & objRS("TopicPosted") & "</td>" & vbCR
  Response.Write "</tr>" & vbCR
  objRS.MoveNext
 Loop
 Response.Write "<tr bgcolor=#003366>" & vbCR
 Response.Write "<td align=center class=cssForumsHeaderSmall colspan=6>" & vbCR
 Response.Write "Totals for this forum - Topics: " & TFTF_Topics & ", Replies: " & TFTF_Replies & " (" & (TFTF_Topics + TFTF_Replies) & " Total Posts)"
 Response.Write "</td></tr>" & vbCR
 Response.Write "</table>" & vbCR
Else
 Response.Write "<table width=100% cellpadding=4 cellspacing=2>" & vbCR
 Response.Write "<tr bgcolor=#003366><td class=cssForumsHeader>&nbsp;</td></tr>" & vbCR
 Response.Write "<tr><td class=cssForumsLargeNoBold>There are currently no topics in this forum.  Why not be the first and start a topic?</tr>" & vbCR
 Response.Write "<tr bgcolor=#003366><td align=left class=cssForumsHeader colspan=6>&nbsp;</td></tr>" & vbCR
 Response.Write "</table>" & vbCR

End If

' Clean up
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing

%>

<!--#include file="footer.inc"-->

