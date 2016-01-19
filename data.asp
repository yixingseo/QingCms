<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%response.codepage = 65001%>
<%response.charset = "utf-8"%>
<%'on error resume next%>
<!--#include file="content/site.config.asp" -->
<!--#include file="include/ide.class.asp" -->
<!--#include file="include/db.class.asp" -->
<%db.path = CONFIG_DATA%>
<!--#include file="include/site.class.asp" -->
<!--#include file="include/log.class.asp" -->
<!--#include file="include/sort.class.asp" -->
<!--#include file="include/news.class.asp" -->
<!--#include file="include/tpl.class.asp" -->
<!--#include file="include/feedback.class.asp" -->

<%
call site.safe()
dim action : action = e.get("action")

select case lcase(action)

	'添加评论'
	case "comment"
		call feedback.insert()
		call e.alertUrl("留言/评论成功，我们将尽快与您取得联系！",site.dic("root"))
		set feedback = nothing

	case "order"
		call feedback.insert()
		call e.alertUrl("订购成功，我们将尽快与您取得联系！",site.dic("root"))
		set feedback = nothing

	case "news"
		id = e.getint("id")
		if not news.reader(id) then
			e.die "{error:未找到}"
		else
			dim json : json = ""
			json = "{""title"":"""& news.title &""",""url"":"""& news.url &"""}"
			e.die json
		end if
end select

set db = nothing : set e =  nothing
%>
