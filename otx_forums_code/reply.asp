<!--#include file="header.inc"-->
<script>current_screen = "";</script>

<script>
//  Make sure they entered something in all fields
function validateForm() {
 if (replyform.replyPoster.value == "") {
  alert("You must enter something in the Your Name field");
  replyform.replyPoster.focus();
  return false;
 }
 if (replyform.replyText.value == "") {
  alert("You must enter something in the Your Reply field");
   replyform.replyText.focus();
  return false;
 }
 if (replyform.replyActivityNotify.value == true && replyform.replyEMailAddress.value == "") {
  alert("You must enter an eMail address if you want to be notified of activity in this topic");
   replyform.replyEMailAddress.focus();
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
replyPoster         = Trim(Request("replyPoster"))
replyText           = Trim(Request("replyText"))
replyIcon           = Trim(Request("replyIcon"))
replyEMailAddress   = Trim(Request("replyEMailAddress"))
replyEMailPublic    = Trim(Request("replyEMailPublic"))
replyActivityNotify = Trim(Request("replyActivityNotify"))

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
Response.Write "<i>Reply</i>&nbsp;|&nbsp;" & vbCR
Response.Write "<a href='forums.asp'>Forums List</a>&nbsp;|&nbsp;" & vbCR
Response.Write "<a href='topicslist.asp?forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>Topics List</a>" & vbCR
Response.Write "</td>" & vbCR
Response.Write "</tr>" & vbCR
Response.Write "</table>" & vbCR
Response.Write "<br>" & vbCR

If inCommand = "showForm" Then ' Show entry form
 Response.Write "<form name=replyform onSubmit='return validateForm();' method=post action='reply.asp?Command=saveReply&topicKey=" & topicKey & "&forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>" & vbCR
 Response.Write "<table width=100% cellpadding=4 cellspacing=2>" & vbCR
 Response.Write "<tr bgcolor=#003366><td colspan=2 class=cssForumsHeader>Reply To Topic...</td></tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>Your Name:</td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7><input type=text name=replyPoster size=30 maxlength=29></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack valign=top>Message Icon:</td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7>" & vbCR
 Response.Write "<table width=300 border=0 cellpadding=0 cellspacing=0 class=cssForumsSmall>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon checked value=default><img src=images/mficon_default.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc1><img src=images/mficon_doc1.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc2><img src=images/mficon_doc2.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc3><img src=images/mficon_doc3.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc4><img src=images/mficon_doc4.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc5><img src=images/mficon_doc5.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc6><img src=images/mficon_doc6.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc7><img src=images/mficon_doc7.gif>&nbsp;</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc8><img src=images/mficon_doc8.gif>&nbsp;</td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc9><img src=images/mficon_doc9.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc10><img src=images/mficon_doc10.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc11><img src=images/mficon_doc11.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc12><img src=images/mficon_doc12.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=doc13><img src=images/mficon_doc13.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=eye><img src=images/mficon_eye.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=face1><img src=images/mficon_face1.gif></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=face2><img src=images/mficon_face2.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=face3><img src=images/mficon_face3.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=face4><img src=images/mficon_face4.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=face5><img src=images/mficon_face5.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=face6><img src=images/mficon_face6.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=exclam1><img src=images/mficon_exclam1.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=info1><img src=images/mficon_info1.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=info2><img src=images/mficon_info2.gif></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=question1><img src=images/mficon_question1.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=question2><img src=images/mficon_question2.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=question3><img src=images/mficon_question3.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=man><img src=images/mficon_man.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=man2><img src=images/mficon_man2.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=woman><img src=images/mficon_woman.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=misc1><img src=images/mficon_misc1.gif></td>" & vbCR
 Response.Write "<td><input type=radio name=replyIcon value=misc2><img src=images/mficon_misc2.gif></td>" & vbCR
 Response.Write "<td>&nbsp;</td>" & vbCR
 Response.Write "<td>&nbsp;</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "</table>" & vbCR
 Response.Write "</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td valign=top bgcolor=#dedfdf class=cssForumsLargeBlack>Your Reply:</td>" & vbCR
 Response.Write "<td bgcolor=#f7f7f7><textarea name=replyText cols=40 rows=10></textarea></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>eMail Address:<br></td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom><input type=text name=replyEMailAddress size=40 maxlength=39></td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>eMail Address Security:<br></td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom class=cssForumsNormal><input type=checkbox name=replyEMailPublic>Allow my eMail address to be seen by all users</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr>" & vbCR
 Response.Write "<td width=25% bgcolor=#dedfdf class=cssForumsLargeBlack>Activity Notification:<br></td>" & vbCR
 Response.Write "<td width=75% bgcolor=#f7f7f7 valign=bottom class=cssForumsNormal><input type=checkbox name=replyActivityNotify>Send me an eMail whenever someone replies</td>" & vbCR
 Response.Write "</tr>" & vbCR
 Response.Write "<tr bgcolor=#003366><td align=right colspan=2><input type=submit value='Submit Reply'></td></tr>" & vbCR
 Response.Write "</table>" & vbCR
 Response.Write "</form>" & vbCR
End If

If inCommand = "saveReply" Then

 ' Replace CR+LF combinations with <BR>
 outText = ""
 for i = 1 to len(replyText)
  if asc(mid(replyText, i, 1)) = 13 then
   outText = outText & "<br>"
  elseif asc(mid(replyText, i, 1)) <> 10 then
   outText = outText & mid(replyText, i, 1)
  end if
 next

 ' Open the connection
 Set objConn = Server.CreateObject("ADODB.Connection")
 objConn.Open "Driver=mySQL;Server=localhost;Port=2214;Option=0;Socket=;Stmt=;Database=forums_database;Uid=forums_dbuser;Pwd=forums_dbuser;"

 ' Retrieve a list of all forum replies
 Set objRS = Server.CreateObject("ADODB.Recordset")
 strSQL = "SELECT * FROM OMNYTEX_ForumsReplies WHERE TopicKey=" & topicKey
 objRS.Open strSQL, objConn, 3, 3
 objRS.AddNew
 objRS("TopicKey") = topicKey
 currentDT = Now
 objRS("ReplyPosted") = currentDT
 objRS("ReplyPoster") = replyPoster
 objRS("ReplyText") = outText
 objRS("ReplyIcon") = replyIcon
 objRS("ReplyEMailAddress") = replyEMailAddress
 If replyEMailPublic = "" Then
  replyEMailPublic = "N"
 Else
  replyEMailPublic = "Y"
 End If
 objRS("ReplyEMailPublic") = replyEMailPublic
 If replyActivityNotify = "" Then
  replyActivityNotify = "N"
 Else
  replyActivityNotify = "Y"
 End If
 objRS("ReplyActivityNotify") = replyActivityNotify
 objRS.Update
 objRS.Close

 ' Update forum stats
 strSQL = "SELECT * FROM OMNYTEX_ForumsForums WHERE ForumsKey=" & forumKey
 objRS.Open strSQL, objConn, 3, 3
 mail_forum_name = objRS("ForumName")
 forumPosts = CInt(objRS("ForumPosts"))
 forumPosts = forumPosts + 1
 objRS("ForumPosts") = forumPosts
 objRS("ForumLastPost") = currentDT
 objRS.Update
 objRS.Close

 ' Update Topic stats
 strSQL = "SELECT * FROM OMNYTEX_ForumsTopics WHERE TopicsKey=" & topicKey
 objRS.Open strSQL, objConn, 3, 3
 mail_topic_title = objRS("TopicTitle")
 mail_topic_key   = objRS("TopicsKey")
 tpnNotify        = objRS("TopicActivityNotify")
 tpnName          = objRS("TopicStarter")
 tpnAddress       = Trim(UCase(objRS("TopicEMailAddress")))
 topicReplies     = objRS("TopicReplies")
 objRS.Close
 topicReplies = topicReplies + 1
 strSQL = "UPDATE OMNYTEX_ForumsTopics SET TopicReplies=" & topicReplies & ", TopicLastReply=now() WHERE TopicsKey=" & topicKey
 objConn.Execute strSQL, dummyvar

 ' Update over-all stats
 strSQL = "SELECT * FROM OMNYTEX_ForumsStats"
 objRS.Open strSQL, objConn, 3, 3
 objRS("StatsNewest") = currentDT
 statsTotalPosts = CInt(objRS("StatsTotalPosts"))
 statsTotalPosts = statsTotalPosts + 1
 objRS("StatsTotalPosts") = statsTotalPosts
 objRS.Update
 objRS.Close

 ' Set up mailer component
 Dim sentAlready(1000)
 sentCount = 0
 Set Mailer = Server.CreateObject("CDONTS.NewMail")
 Mailer.From = "omnytex@omnytex.com"
 Mailer.Subject     = "Notification of thread activity"

 ' If the topic poster wants to be notified, and it's not the person posting the reply, send to them now
 If tpnNotify = "Y" And tpnAddress <> Trim(UCase(replyEMailAddress)) Then
  outSubj             = vbCR & " Hello, " & tpnName & "!" & vbCR & vbCR
  outSubj = outSubj   & " This is an automated notification, as per your request, to alert you of a new reply to the topic" & vbCR & vbCR
  outSubj = outSubj   & " " & Chr(34) & mail_topic_title & Chr(34) & " in the Omnytex Technologies message forum " & UCase(mail_forum_name) & vbCR & vbCR
  outSubj = outSubj   & " You can use the URL http://www.omnytex.com/forums.asp to access the forums.  Have a great day!"
  Mailer.Body     = outSubj
  Mailer.To = tpnAddress
  Mailer.Send
  sentCount = sentCount + 1
  sentAlready(sentCount) = ReplierEMailAddress
 End If

 ' Grab all replies to this topic
 strSQL = "SELECT * FROM OMNYTEX_ForumsReplies WHERE TopicKey=" & topicKey
 objRS.Open strSQL, objConn, 3, 3

 ' If there were replies, scan through them...
 If objRS.EOF <> True Then
  Do Until objRS.EOF = True
   ReplierEMailAddress = Trim(UCase(objRS("ReplyEMailAddress")))

   ' See if we already sent to this address, skip it if so
   skipMe = 0
   If sentCount <> 0 Then
    For i = 1 to sentCount
     If ReplierEMailAddress = sentAlready(i) Then
      skipMe = 1
     End If
    Next
   End If

   ' If the reply poster is the topic starter, skip because we would have sent to them already
   If ReplierEMailAddress = tpnAddress Then
    skipMe = 1
   End If

   ' If this address is the reply poster's address, skip it
   If ReplierEMailAddress = Trim(UCase(replyEMailAddress)) Then
    skipMe = 1
   End If

   ' If this person didn't want to be notified, skip them
   If objRS("ReplyActivityNotify") <> "Y" Then
    skipMe = 1
   End IF

   ' If this address doesn't match tpnAddress and we're not skipping it, go ahead and send
   If skipMe = 0 Then
    outSubj             = vbCR & " Hello, " & objRS("ReplyPoster") & "!" & vbCR & vbCR
    outSubj = outSubj   & " This is an automated notification, as per your request, to alert you of a new reply to the topic" & vbCR
    outSubj = outSubj   & " " & Chr(34) & mail_topic_title & Chr(34) & " in the Omnytex Technologies message forum " & UCase(mail_forum_name) & vbCR & vbCR
    outSubj = outSubj   & " You can use the URL http://www.omnytex.com/forums.asp to access the forums.  Have a great day!"
    Mailer.From = "omnytex@omnytex.com"
    Mailer.Subject     = "Notification of thread activity"
    Mailer.Body     = outSubj
    Mailer.To = objRS("ReplyEMailAddress")
    Mailer.Send
    sentCount = sentCount + 1
    sentAlready(sentCount) = ReplierEMailAddress
   End If
   objRS.MoveNext
  Loop
 End If

 ' Clean up
 Set Mailer = Nothing

 objRS.Close
 Set objRS = Nothing
 objConn.Close
 Set objConn = Nothing

 ' Report...
 Response.Write "<hr width=100% ><br><table class=cssForumsLargeNoBold width=100% border=0 cellpadding=0 cellspacing=0><tr><td>Thank you for posting!  Click <a href='topic.asp?topicKey=" & topicKey & "&forumKey=" & forumKey & "&forumName=" & Server.URLEncode(forumName) & "'>here</a> to return to the topic you just replied to.</td></tr></table>" & vbCR

End If

%>

<!--#include file="footer.inc"-->






