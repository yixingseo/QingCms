<%
'数据库'
'2015-11-27'

dim db : set db = new ClassDb
class ClassDb

	public conn,connString,pager

	private sub class_initialize()
		intselectcount = 0
		set conn = server.CreateObject("adodb.connection")
	end sub

	private sub class_terminate()
		if conn.state = 1 then conn.close()
		set conn = nothing
	end sub

	property let path(s)
		on error resume next
		if not e.fileExists(s) then e.msg("数据库不存在")
		connString = "provider=microsoft.jet.oledb.4.0;data source=" & server.mappath(s)
		conn.open(connString)
		if err then e.msg(err.Description)
	end property

	'查询数据(sql) 'obj,recordset
	function query(strSql)
		set query = e.objRs
			query.open strSql,conn,1,1
		if err then e.msg "查询失败：" & strSql
	end function

	'执行操作(sql) '受影响行数 int
	function exec(strSql)
		on error resume next
		conn.execute strSql,exec
		intselectcount = intselectcount + 1
		if err then e.die "查询失败：" & strSql
	end function

	'获取分页(sql,分页大小)'obj,redcordset
	function listpage(strSql,intpagesize)
		on error resume next
		set listpage = e.objRs
			listpage.open strSql,conn,1,1
			if listpage.eof then exit function
			listpage.pagesize = intpagesize
			listpage.absolutepage = e.iif(e.getid("page") > 0 ,e.getid("page"),1)
		call setPager(listpage.pagecount,listpage.absolutepage)
		if err then e.die "查询失败：" & strSql
	end function

	'设置分页
	function setPager(intpagecount,intpagecurrent)
		dim pageurl : pageurl = e.get_url("-page")
			pageurl = pageurl & e.iif(instr(pageurl,"?") > 0 ,"&","?") & "page="
		dim s
			s = "<ul class='pagination{0}'>"

		if intpagecurrent > 1 then
			s = s & "<li><a href='"& pageurl & intpagecurrent - 1 &"'>&laquo;</a></li>"
		else
			s = s & "<li class='disabled'><a href='#' aria-label='previous'><span aria-hidden='true'>&laquo;</span></a></li>"
		end if
		dim intpagei
		for intpagei = 1 to intpagecount
			if intpagei = intpagecurrent then
				s = s & "<li class='active'><a href='#'>"& intpagei &" <span class='sr-only'>(current)</span></a></li>"
			else
				s = s & "<li><a href='"& pageurl & intpagei &"'>" & intpagei & "</a></li>"
			end if
		next
		if intpagecurrent < intpagecount then
			s = s & "<li><a href='"& pageurl & page_current + 1 &"'>&raquo;</a></li>"
		else
			s = s & "<li class='disabled'><a href='#' aria-label='next'><span aria-hidden='true'>&raquo;</span></a></li>"
		end if

		pager = s & "</ul>"
	end function

	'获取分页'
	function getpager(strsize)
		dim s
		select case lcase(strsize)
			case "lg"
				s = replace(pager,"{0}"," pagination-lg")
			case "sm"
				s = replace(pager,"{0}"," pagination-sm")
			case else
				s = replace(pager,"{0}"," " & strsize)
		end select
		getpager = s
	end function

end class
%>
