<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/feedback.class.asp" -->

<%
dim action : action = e.get("action")

select case lcase(action)
	case "rep"
		call feedback.rep()
	case "delete"
		call feedback.delete()

	case "mail_test"
		e.die(feedback.testMail())

end select

redirect "feedback.listview.asp"
%>

<!--#include file="cms.footer.asp"-->
