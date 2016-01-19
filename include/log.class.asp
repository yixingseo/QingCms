<%
'日志'
'2015-11-27'
dim log : set log = new t_log

class t_log
	
	public id,user,content,insert_date,insert_ip,title

	sub class_initialize
		id = 0
		title = ""
		insert_date = now()
		insert_ip = e.getip()
		user = e.iif(session("admin") <> "",session("admin"),"GUEST")
	end sub

	'读取日志'
	sub reader()
		set rs = db.query("select top 1 * from [t_logs] where id =" & id)
		if rs.eof then exit sub
		title = rs("title")
		content = e.encode(rs("content"))
		insert_date = rs("insert_date")
		insert_ip = rs("insert_ip")
		user = rs("user")
		rs.close
		set rs = nothing
	end sub

	'保存日志'
	sub saveData()		
		set rs = e.objRs
		rs.open "select top 1 * from [t_logs] where id = 0",db.conn,1,3
		rs.addnew
			rs("user") = user
			rs("title") = title
			rs("content") = content
			rs("insert_date") = insert_date
			rs("insert_ip") = insert_ip
		rs.update
		rs.close
		set rs = nothing
	end sub

	'删除日志'
	sub del()
		if not site.isSuper() then e.alert("权限不足")
		id = e.get("id")
		if not isnumeric(id) then exit sub
		if db.exec("delete from [t_logs] where id in("& id&")") < 1 then e.die "操作失败"
	end sub

	'清空日志'
	sub clear()
		if not site.isSuper() then e.alert("权限不足")
		db.exec("delete from [t_logs] where id > 0")
		add "清空日志",  "清空了所有日志"
	end sub

	sub ajax_content()		
		id = e.getint("id")
		if id = 0 then exit sub
		call reader()
		e.die content
	end sub

	'记录操作'
	sub add(sTitle,sLog)
		title = e.safe(sTitle)
		content = e.safe(sLog)
		result = db.exec("insert into [t_logs] ([user],[title],[content],[insert_ip],[insert_date]) values('"& user &"','"& title &"','"& content &"','"& insert_ip &"','"& insert_date&"')")				
	end sub
	
end class

%>