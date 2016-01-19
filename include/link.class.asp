<%
'友情链接'
'2015-11-27'
dim link : set link = new ClassLink
class ClassLink
	public id,title,url,target,pic,weight
	public table

	private sub class_initialize()
		id = 0
		target = "_blank"
		url = "http://"
		table = "[t_link]"
		weight = CONFIG_WEIGHT
	end sub

	dim rs

	function read(intID)
		read = false
		if isNull(intID) or not isNumeric(intID) then exit function
			set rs = db.query("select top 1 * from "& table &" where id = " & intID)
			if rs.eof then exit function
				id = rs("id")
				title = e.encode(rs("title"))
				url = rs("url")
				target = rs("target")
				pic = rs("pic")
				weight = rs("weight")
			rs.close
			set rs = nothing
		read = true
	end function

	sub postData()
		id = e.postint("id")
		title = e.post("title")
		if title = "" then e.alert("请输入链接名称")
		url = e.post("url")
		if url = "" then e.alert("请输入网址")
		if not e.test("url",url) then e.alert("网址不正确")
		target = e.post("target")
		pic = e.post("pic")
		weight = e.postint("weight")
	end sub

	function create()
		call postData()
		set rs = e.objRs
			rs.open "select top 1 * from "& table,db.conn,1,3
			rs.addnew
			rs("title") = title
			rs("pic") = pic
			rs("url") = url
			rs("target") = target
			rs("weight") = weight
		rs.update
		rs.close
		set rs = nothing
	end function

	function update()
		call postData()
		set rs = e.objRs
			rs.open "select top 1 * from "& table &" where id = " & id,db.conn,1,3
			rs("title") = title
			rs("pic") = pic
			rs("url") = url
			rs("target") = target
			rs("weight") = weight
		rs.update
		rs.close
		set rs = nothing
	end function

	function delete()
		id = e.get("id")
		if not isnumeric(id) then exit function
		if instr(id,",") > 0 then
			set rs = db.query("select pic from "& table &" where id in("& id &")")
			while not rs.eof
				e.fileDelete(rs(0))
			rs.movenext
			wend
			rs.close
			set rs = nothing
			if db.exec("delete from "& table &" where id in("& id &")") < 1 then e.die "操作失败"
		else
			id = cint(id)
			if read(id) then
				if db.exec("delete from "& table &" where id = "& id &"") < 1 then e.die "操作失败"
				if pic <> "" then e.fileDelete(pic)
			end if
		end if
	end function

end class
%>
