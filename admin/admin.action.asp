<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/admin.class.asp" -->
<%
dim action : action = e.get("action")
dim admin : set admin = new t_admin

select case lcase(action)
	case "save"
		call admin.saveData()

	case "del_data"
		call admin.del()

	case "password"
		call admin.changePassword()

end select

redirect "admin.listview.asp"

%>

<!--#include file="cms.footer.asp" -->