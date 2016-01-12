---
layout: default
title: Content-Class Types
name: ContentClassTypes
---

**Content-Class Types**

# <font size="2">Original article from:</font>

# [<font size="1">http://msdn.microsoft.com/library/en-us/wss/wss/_exchsv2k_mapping_pr_message_class_to_dav_contentclass.asp</font>](http://msdn.microsoft.com/library/en-us/wss/wss/_exchsv2k_mapping_pr_message_class_to_dav_contentclass.asp)

# <font size="5">Mapping PR_MESSAGE_CLASS to DAV:contentclass</font>

* * *

When the item is not a folder, the value of its PR_MESSAGE_CLASS property is examined for a match in the following tables, traversing the first column from top to bottom. The first match found is returned as the value for the item's DAV:contentclass property. The asterisk (*) wildcard character denotes any set of characters. The dollar-sign ($) character denotes the end of the string. All other strings must be matched exactly. The tables are broken down into three categories: values with the prefix [<u><font color="#0000ff">IPC</font></u>](http://msdn.microsoft.com/library/en-us/wss/wss/_exchsv2k_mapping_pr_message_class_to_dav_contentclass.asp?frame=true#IPC), [<u><font color="#0000ff">IPM</font></u>](http://msdn.microsoft.com/library/en-us/wss/wss/_exchsv2k_mapping_pr_message_class_to_dav_contentclass.asp?frame=true#IPM), and [<u><font color="#0000ff">REPORT</font></u>](http://msdn.microsoft.com/library/en-us/wss/wss/_exchsv2k_mapping_pr_message_class_to_dav_contentclass.asp?frame=true#REPORT). All values not resolved in one of the three tables map to the value "urn:content-classes:document".

## <a name="IPC"></a>IPC

<table class="clsStd" cellspacing="0" cellpadding="0" summary="" border="1">

<tbody>

<tr>

<th>PR_MESSAGE_CLASS</th>

</tr>

<tr>

<td>IPC</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPC.*</td>

<td>urn:content-classes:document</td>

</tr>

</tbody>

</table>

## <a name="IPM"></a>IPM

<table class="clsStd" cellspacing="0" cellpadding="0" summary="" border="1">

<tbody>

<tr>

<th>PR_MESSAGE_CLASS</th>

<th>Content Class</th>

</tr>

<tr>

<td>IPM</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Activity</td>

<td>urn:content-classes:activity</td>

</tr>

<tr>

<td>IPM.Appointment</td>

<td>urn:content-classes:appointment</td>

</tr>

<tr>

<td>IPM.Conflict.Resolution.Message</td>

<td>http://content-classes.microsoft.com/exchange/conflict</td>

</tr>

<tr>

<td>IPM.Contact</td>

<td>urn:content-classes:person</td>

</tr>

<tr>

<td>IPM.ContentClassDef</td>

<td>urn:content-classes:contentclassdef</td>

</tr>

<tr>

<td>IPM.DistList</td>

<td>urn:content-classes:group</td>

</tr>

<tr>

<td>IPM.Document</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.*doc</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.Excel.Sheet.5</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.Excel.Sheet.8</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.Microsoft Internet Mail Message</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Document.MSProject.Project.4_1</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.MSProject.Project.8</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.PowerPoint.Show.4</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.PowerPoint.Show.7</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.PowerPoint.Show.8</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.textfile</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.Word.Document.6</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Document.Word.Document.8</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.Microsoft.KeyExchange</td>

<td>http://content-classes.microsoft.com/exchange/keyexchange</td>

</tr>

<tr>

<td>IPM.Microsoft.ScheduleData.FreeBusy</td>

<td>urn:content-classes:freebusy</td>

</tr>

<tr>

<td>IPM.Note</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.Exchange.Security.Enrollment</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.IMC.Notification</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.P772</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.RFC822.MIME</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.Rules.OofTemplate.Microsoft</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.Rules.ReplyTemplate.Microsoft</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.Secure</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.Secure.Service.Reply</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.Secure.Sign</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.SMIME</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.SMIME.MultipartSigned</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Note.StorageQuotaWarning</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Object</td>

<td>urn:content-classes:object</td>

</tr>

<tr>

<td>IPM.Organization</td>

<td>urn:content-classes:organization</td>

</tr>

<tr>

<td>IPM.Outlook.Recall</td>

<td>urn:content-classes:recallmessage</td>

</tr>

<tr>

<td>IPM.Post</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.PropertyDef</td>

<td>urn:content-classes:propertydef</td>

</tr>

<tr>

<td>IPM.Recall.Report</td>

<td>urn:content-classes:recallreport</td>

</tr>

<tr>

<td>IPM.Recall.Report.Failure</td>

<td>urn:content-classes:recallreport</td>

</tr>

<tr>

<td>IPM.Recall.Report.Success</td>

<td>urn:content-classes:recallreport</td>

</tr>

<tr>

<td>IPM.Report</td>

<td>urn:content-classes:reportmessage</td>

</tr>

<tr>

<td>IPM.Schedule.Meeting.Canceled</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.Schedule.Meeting.Request</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.Schedule.Meeting.Resp.Neg</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.Schedule.Meeting.Resp.Pos</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.Schedule.Meeting.Resp.Tent</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.StickyNote</td>

<td>urn:content-classes:note</td>

</tr>

<tr>

<td>IPM.Task</td>

<td>urn:content-classes:task</td>

</tr>

<tr>

<td>IPM.TaskRequest</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.TaskRequest.Accept</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.TaskRequest.Decline</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.TaskRequest.Update</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.Appointment.*</td>

<td>urn:content-classes:appointment</td>

</tr>

<tr>

<td>IPM.Schedule.Meeting.*</td>

<td>urn:content-classes:calendarmessage</td>

</tr>

<tr>

<td>IPM.Contact.*</td>

<td>urn:content-classes:person</td>

</tr>

<tr>

<td>IPM.Note.*</td>

<td>urn:content-classes:message</td>

</tr>

<tr>

<td>IPM.Document.*</td>

<td>urn:content-classes:document</td>

</tr>

<tr>

<td>IPM.*</td>

<td>urn:content-classes:message</td>

</tr>

</tbody>

</table>

## <a name="REPORT"></a>REPORT

<table class="clsStd" cellspacing="0" cellpadding="0" summary="" border="1">

<tbody>

<tr>

<th>PR_MESSAGE_CLASS</th>

<th>Content Class</th>

</tr>

<tr>

<td>Report</td>

<td>urn:content-classes:reportmessage</td>

</tr>

<tr>

<td>Report.IPM.Note.DR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.Note.IPNNDR</td>

<td>urn:content-classes:mdn</td>

</tr>

<tr>

<td>Report.IPM.Note.IPNNRN</td>

<td>urn:content-classes:mdn</td>

</tr>

<tr>

<td>Report.IPM.Note.IPNRN</td>

<td>urn:content-classes:mdn</td>

</tr>

<tr>

<td>Report.IPM.Note.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.Recall.Report.Failure.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.Schedule.Meeting.Canceled.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.Schedule.Meeting.Request.DR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.Schedule.Meeting.Request.IPNNRN</td>

<td>urn:content-classes:mdn</td>

</tr>

<tr>

<td>Report.IPM.Schedule.Meeting.Request.IPNRN</td>

<td>urn:content-classes:mdn</td>

</tr>

<tr>

<td>Report.IPM.Schedule.Meeting.Request.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.Schedule.Meeting.Resp.Neg.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.Schedule.Meeting.Resp.Pos.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.Schedule.Meeting.Resp.Tent.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.TaskRequest.Accept.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.TaskRequest.Decline.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.TaskRequest.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>Report.IPM.TaskRequest.Update.NDR</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>REPORT.*.DR$</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>REPORT.*.IPNNDR$</td>

<td>urn:content-classes:mdn</td>

</tr>

<tr>

<td>REPORT.*.IPNNRN$</td>

<td>urn:content-classes:mdn</td>

</tr>

<tr>

<td>REPORT.*.IPNRN$</td>

<td>urn:content-classes:mdn</td>

</tr>

<tr>

<td>REPORT.*.NDR$</td>

<td>urn:content-classes:dsn</td>

</tr>

<tr>

<td>REPORT.*</td>

<td>urn:content-classes:message</td>

</tr>

</tbody>

</table>

## All Other Values

All other values in PR_MESSAGE_CLASS map to the DAV:contentclass value "urn:content-classes:document".
