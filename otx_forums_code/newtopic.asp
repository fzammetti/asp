<!--#include file="header.inc"-->
<script>current_screen = "";</script>

<script>
//  Make sure they entered something in all fields
function validateForm() {
 if (newtopicform.topicStarter.value == "") {
  alert("You must enter something in the Topic Starter field");
  newtopicform.topicStarter.focus();
  return false;
 }
 if (newtopicform.topicTitle.value == "") {
  alert("You must enter something in the Topic Title field");
   newtopicform.topicTitle.focus();
  return false;
 }
 if (newtopicform.topicText.value == "") {
  alert("You must enter something in the Text field");
   newtopicform.topicText.focus();
  return false;
 }
 if (newtopicform.topicActivityNotify.value == true && newtopicform.topicEMailAddress.value == "") {
  alert("You must enter an eMail address if you want to be notified of activity in this topic");
   newtopicform.topicEMailAddress.focus();
  return false;
 }
 return true;
}
</script>

<%

' Declare variables
Dim objConn
Dim objRS
Dim strSQL

' Get input variables
inCommand           = Trim(Request("Command"))
forumKey            = Trim(Request("forumKey"))
forumName           = Trim(Request("forumName"))
topicKey            = Trim(Request("topicKey"))
topicStarter        = Trim(Request("topicStarter"))
topicTitle          = Trim(Request("topicTitle"))
topicText           = Trim(Request("topicText"))
topicIcon           = Trim(Request("topicIcon"))
topicEMailAddress   = Trim(Request("topicEMailAddress"))
topicEMailPublic    = Trim(Request("topicEMailPublic"))
topicActivityNotify = Trim(Request("topicActivityNotify"))

' Header
Response.Write "<table width=100% border=0 cellpadding=0 cellspacing=0>" & vbCR
Response.Write "<tr>" & vbCR
Response.Write "<td class=cssForumsNormal>&nbsp;</td>" & vbCR
Response.Write "<td align=right class=cssForumsTitle>Omnytex Message Forums</td>" & vbCR
Response.Write "</tr>" & vbCR
Response.Write "<tr><td colspan=2 align=right><hr width=100% height=1 color=#000000></td></tr>" & vbCR
Response.Write "<tr><td class=cssForumsNormal>&nbsp;</td>" & vbCR
Response.Write "<td align=right class=cssForumsNormal>" & vbCR
Response.Write "<i>New Topic</i>&nbsp;|&nbsp;" & vbCR
Response.Write "<strike>Reply</strike>&nbsp;|&nbsp;" & vbCR
Response.Write "<a href='forums.asp'>Forums List</a>&nbsp;|&nbsp;" & vbCR
Response.Write "<a href='topicslist.asp?forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>Topics List</a>" & vbCR
Response.Write "</td>" & vbCR
Response.Write "</tr>" & vbCR
Response.Write "</table>" & vbCR
Response.Write "<br>" & vbCR

If inCommand = "showForm" Then ' Show entry form
 Response.Write "<form name=newtopicform onSubmit='return validateForm();' method=post action='newtopic.asp?Command=saveTopic&forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>" & vbCR
 Response.Write "<table width=100% cellpadding=4 cellspacing=2>" & vbCR
 Response.Write "<tr bgcolor=#003366><td colspan=2 align=left class=cssForumsHeader>Start New Topic...</td></tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>Your Name:</td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom><input type=text name=topicStarter size=30 maxlength=29></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>Topic Title:</td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom><input type=text name=topicTitle size=50 maxlength=200></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack valign=top>Message Icon:</td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom>" & vbCR
 Response.Write "<table width=300 border=0 cellpadding=0 cellspacing=0 class=cssForumsSmall>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon checked value=default><img src=images/mficon_default.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc1><img src=images/mficon_doc1.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc2><img src=images/mficon_doc2.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc3><img src=images/mficon_doc3.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc4><img src=images/mficon_doc4.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc5><img src=images/mficon_doc5.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc6><img src=images/mficon_doc6.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc7><img src=images/mficon_doc7.gif>&nbsp;</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc8><img src=images/mficon_doc8.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc9><img src=images/mficon_doc9.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc10><img src=images/mficon_doc10.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc11><img src=images/mficon_doc11.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc12><img src=images/mficon_doc12.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=doc13><img src=images/mficon_doc13.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=eye><img src=images/mficon_eye.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=face1><img src=images/mficon_face1.gif></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=face2><img src=images/mficon_face2.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=face3><img src=images/mficon_face3.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=face4><img src=images/mficon_face4.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=face5><img src=images/mficon_face5.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=face6><img src=images/mficon_face6.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=exclam1><img src=images/mficon_exclam1.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=info1><img src=images/mficon_info1.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=info2><img src=images/mficon_info2.gif></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=question1><img src=images/mficon_question1.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=question2><img src=images/mficon_question2.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=question3><img src=images/mficon_question3.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=man><img src=images/mficon_man.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=man2><img src=images/mficon_man2.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=woman><img src=images/mficon_woman.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=misc1><img src=images/mficon_misc1.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=topicIcon value=misc2><img src=images/mficon_misc2.gif></td>" & vbCR
 Response.Write "<td>&nbsp;</td>" & vbCR
 Response.Write "<td>&nbsp;</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "</table>" & vbCR
 Response.Write "</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td valign=top bgcolor=#dedfdf class=cssForumsLargeBlack>Text:</td>" & vbCR
 Response.Write "<td bgcolor=#f7f7f7 valign=bottom><textarea name=topicText cols=40 rows=10></textarea></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>eMail Address:<br></td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom><input type=text name=topicEMailAddress size=40 maxlength=39></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>eMail Address Security:<br></td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom class=cssForumsNormal><input type=checkbox name=topicEMailPublic>Allow my eMail address to be seen by all users</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>Activity Notification:<br></td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom class=cssForumsNormal><input type=checkbox name=topicActivityNotify>Send me an eMail whenever someone replies</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr bgcolor=#003366><td align=right colspan=2><input type=submit value='Submit Topic'></td></tr>" & vbCR
 Response.Write "</table></form>" & vbCR
End If

If inCommand = "saveTopic" Then ' Write the topic
 ' Replace CR+LF combinations with <BR>
 outText = ""
 for i = 1 to len(topicText)
  if asc(mid(topicText, i, 1)) = 13 then
   outText = outText & "<br>"
  elseif asc(mid(topicText, i, 1)) <> 10 then
   outText = outText & mid(topicText, i, 1)
  end if
 next
 ' Open the connection
 Set objConn = Server.CreateObject("ADODB.Connection")
 objConn.Open "Driver=mySQL;Server=localhost;Port=2214;Option=0;Socket=;Stmt=;Database=forums_database;Uid=forums_dbuser;Pwd=forums_dbuser;"
 ' Retrieve a list of all topics
 Set objRS = Server.CreateObject("ADODB.Recordset")
 strSQL = "SELECT * FROM OMNYTEX_ForumsTopics WHERE ForumKey=" & forumKey
 objRS.Open strSQL, objConn, 3, 3
 objRS.AddNew
 objRS("ForumKey") = forumKey
 currentDT = Now
 objRS("TopicPosted") = currentDT
 objRS("TopicTitle") = topicTitle
 objRS("TopicStarter") = topicStarter
 objRS("TopicReplies") = 0
 objRS("TopicText") = outText
 objRS("TopicIcon") = topicIcon
 objRS("TopicEMailAddress") = topicEMailAddress
 If topicEMailPublic = "" Then
  topicEMailPublic = "N"
 Else
  topicEMailPublic = "Y"
 End If
 objRS("TopicEMailPublic") = topicEMailPublic
 If topicActivityNotify = "" Then
  topicActivityNotify = "N"
 Else
  topicActivityNotify = "Y"
 End If
 objRS("TopicActivityNotify") = topicActivityNotify
 objRS.Update
 objRS.Close
 ' Update forum stats
 strSQL = "SELECT * FROM OMNYTEX_ForumsForums WHERE ForumsKey=" & forumKey
 objRS.Open strSQL, objConn, 3, 3
 objRS("ForumTopics") = objRS("ForumTopics") + 1
 objRS("ForumPosts") = objRS("ForumPosts") + 1
 objRS("ForumLastPost") = currentDT
 objRS.Update
 objRS.Close
 ' Update over-all stats
 strSQL = "SELECT * FROM OMNYTEX_ForumsStats"
 objRS.Open strSQL, objConn, 3, 3
 objRS("StatsNewest") = currentDT
 objRS("StatsTotalTopics") = objRS("StatsTotalTopics") + 1
 objRS("StatsTotalPosts") = objRS("StatsTotalPosts") + 1
 objRS.Update
 ' Clean up
 objRS.Close
 Set objRS = Nothing
 objConn.Close
 Set objConn = Nothing
 Response.Write "<hr width=100% ><br><table class=cssForumsLargeNoBold width=100% border=0 cellpadding=0 cellspacing=0><tr><td>Thank you for posting!  Click <a href='topicslist.asp?forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>here</a> to return to the topics list for this forum.</td></tr></table>" & vbCR
End If

%>

<!--#include file="footer.inc"-->

