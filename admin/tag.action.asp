<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/tag.class.asp" -->
<%
dim action : action = e.get("action")
select case lcase(action)
	case "create"
		call tag.create()
	case "update"
		call tag.update()
	case "delete"
		call tag.delete()
end select

redirect "tag.listview.asp"
%>

<!--#include file="cms.footer.asp" -->
