<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/sort.class.asp" -->
<%
dim action : action = e.get("action")
select case lcase(action)
	case "create"
		call sort.create()
	case "update"
		call sort.update()
	case "delete"
		call sort.delete()
	case "ajax_weight"
		call sort.ajaxWeight()
	case "ajax_title"
		call sort.ajaxTitle()
end select
set sort = nothing
redirect("sort.listview.asp")
%>
<!--#include file="cms.footer.asp" -->
