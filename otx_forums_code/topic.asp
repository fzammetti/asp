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
Response.Write "<a href='reply.asp?Command=showForm&topicKey=" & topicKey & "&forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>Reply</a>&nbsp;|&nbsp;" & vbCR
Response.Write "<a href='forums.asp'>Forums List</a>&nbsp;|&nbsp;" & vbCR
Response.Write "<a href='topicslist.asp?forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>Topics List</a>" & vbCR
Response.Write "</td>" & vbCR
Response.Write "</tr>" & vbCR
Response.Write "</table>" & vbCR
Response.Write "<br>" & vbCR

' Open the connection
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Driver=mySQL;Server=localhost;Port=2214;Option=0;Socket=;Stmt=;Database=forums_database;Uid=forums_dbuser;Pwd=forums_dbuser;"

' Retrieve a list of all forums
Set objRS = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM OMNYTEX_ForumsTopics WHERE TopicsKey=" & topicKey
objRS.Open strSQL, objConn, 3, 3

' Create our output (topic)
If objRS.EOF <> True Then
 Response.Write "<table width=100% cellpadding=4 cellspacing=2>" & vbCR
 Response.Write "<tr bgcolor=#003366>" & vbCR
 Response.Write "<td width=25% align=left class=cssForumsHeader>Author</td>" & vbCR
 Response.Write "<td width=75% align=left class=cssForumsHeader>Topic: " & objRS("TopicTitle") & "</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr bgcolor=#dedfdf>" & vbCR
 Response.Write "<td valign=top class=cssForumsLargeBlack>" & objRS("TopicStarter")
 If objRS("TopicEMailAddress") <> "" And objRS("TopicEMailPublic") = "Y" Then
  Response.Write "<br><span class=cssFooter>(<a href='mailto:" & objRS("TopicEMailAddress") & "'>eMail Author</a>)</span>" & vbCR
 End If
 Response.Write "</td>" & vbCR
 Response.Write "<td class=cssForumsSmall><img src=images/mficon_" & objRS("TopicIcon") & ".gif>&nbsp;Posted " & objRS("TopicPosted") & "<br>" & vbCR
 Response.Write "<hr width=100% ><span class=cssForumsMediumNoBold>" & objRS("TopicText") & "</span><br><br></td>" & vbCR
 Response.Write "</tr>" & vbCR
Else
 Response.Write "No Data" & vbCR
End If

' Save the number of replies to this topic
numReplies = objRS("TopicReplies")

' Get rid of the recordset
objRS.Close

' Open recordset to get replies and create output if here were replies
If numReplies <> 0 Then
 strSQL = "SELECT * FROM OMNYTEX_ForumsReplies WHERE TopicKey=" & topicKey & " ORDER BY ReplyPosted"
 objRS.Open strSQL, objConn, 3, 3
 lastColor = 0
 Do Until objRS.EOF = True
  If lastColor = 0 Then
   lastColor = 1
   Response.Write "<tr bgcolor=#f7f7f7>" & vbCR
  Else
   lastColor = 0
   Response.Write "<tr bgcolor=#dedfdf>" & vbCR
  End If
  Response.Write "<td valign=top class=cssForumsLargeBlack>" & objRS("ReplyPoster")
  If objRS("ReplyEMailAddress") <> "" And objRS("ReplyEMailPublic") = "Y" Then
   Response.Write "<br><span class=cssFooter>(<a href='mailto:" & objRS("ReplyEMailAddress") & "'>eMail Author</a>)</span>" & vbCR
  End If
  Response.Write "<td class=cssForumsSmall><img src=images/mficon_" & objRS("ReplyIcon") & ".gif>&nbsp;Posted " & objRS("ReplyPosted") & "<br>" & vbCR
  Response.Write "<hr width=100% ><span class=cssForumsMediumNoBold>" & objRS("ReplyText") & "</span><br><br></td>" & vbCR
  Response.Write "</tr>" & vbCR
  objRS.MoveNext
 Loop
 objRS.Close
End If

Set objRS = Nothing
objConn.Close
Set objConn = Nothing

' Reply button
Response.Write "<tr bgcolor=#003366><td align=right colspan=2>&nbsp;</td></tr>" & vbCR
Response.Write "</table>" & vbCR

%>

<!--#include file="footer.inc"-->

