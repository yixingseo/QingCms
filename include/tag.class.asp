<%
'标签'
dim tag : set tag = new ClassTag

class ClassTag
	public id,title,description,content,table
	dim rs

	sub class_initialize()
		id = 0
		table = "[t_tag]"
	end sub

	'读取'
	function read(intID)
		read = false
		if isNull(intID) or not isNumeric(intID) then exit function
		set rs = db.query("select top 1 * from "& table &" where id =" & intID)
		if rs.eof then exit function
			id = rs("id")
			title = rs("title")
			description = rs("description")
			content = rs("content")
		rs.close
		set rs = nothing
	end function

	sub postData()
		id = e.postint("id")
		title = e.post("title")
		if title = "" then e.alert("请输入标签名称")
		if not e.test("username",title)	 then e.alert("标签名称不合法")
		description = e.post("description")
		content = e.post("content")
	end sub

	'新增'
	function create()
		call postData()
		set rs = e.objRs
			rs.open "select top 1 * from "& table,db.conn,1,3
			rs.addnew()
			rs("title") = title
			rs("description") = description
			rs("content") = content
			rs.update()
			rs.close
		set rs = nothing
	end function

	'修改'
	function update()
		call postData()
		set rs = e.objRs
			rs.open "select top 1 * from "& table &" where id =" & id,db.conn,1,3
			rs("title") = title
			rs("description") = description
			rs("content") = content
			rs.update()
			rs.close
		set rs = nothing
	end function

	'删除'
	function delete()
		id = e.getInt("id")
		if id = 0 then exit function
		delete = db.exec("delete from "& table &" where id = " & id)
	end function

	'名字检测'
	function checkTitle()
		if db.query("select count(*) from "& table &" where title = '"& title &"' and id <> " & id)(0) > 0 then
			checkTitle = false
		else
			checkTitle = true
		end if
	end function

end class
%>
