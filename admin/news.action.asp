<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/sort.class.asp" -->
<!--#include file="../include/news.class.asp" -->
<!--#include file="../include/upload.class.asp" -->
<%
dim action : action = e.get("action")
select case lcase(action)
	case "create"
		call news.create()
	case "update"
		call news.update()
	case "delete"
		call news.delete()
	case "rec"
		call news.rec()
	case "display"
		call news.display()
	case "ajax_weight"
		call news.ajaxWeight()
	case "ajax_title"
		call news.ajaxTitle()
	case "delatt"
		call news.delAtt()
	case "addatt"
		call news.addAtt()
end select

set news = nothing
redirect "news.listview.asp"
%>
<!--#include file="cms.footer.asp" -->
