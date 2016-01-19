<%
'分类'
'2015-11-27'
dim sort : set sort = new ClassSort
class ClassSort
	public id,title,pid,url,urlname,sort_template,content_template,position,content,pic,guid,weight
	public table
	dim rs

	private sub class_initialize()
		id = 0
		pid = 0
		weight = CONFIG_WEIGHT
		guid = e.guid()
		table = "[t_sort]"
	end sub

	'读取分类(id) 'bool
	function read(intID)
		read = false
		if isNull(intID) or not isNumeric(intID) then exit function
		set rs = db.query("select top 1 * from "& table &" where id = " & intID)
		if rs.eof then exit function
			id = rs("id")
			title = rs("title")
			pid = rs("pid")
			url = rs("url")
			weight = rs("weight")
			urlname = rs("urlname")
			sort_template = rs("sort_template")
			content_template = rs("content_template")
			position = rs("position")
			content = rs("content")
			pic = rs("pic")
			rs.close
			set rs = nothing
		read = true
	end function

	'读取分类(url)'
	function readerByUrl(sUrl)
		readerByUrl = false
		set rs = db.query("select top 1 * from "& table &" where url = '"& sUrl &"'")
		if rs.eof then exit function
			id = rs("id")
			title = rs("title")
			pid = rs("pid")
			url = rs("url")
			weight = rs("weight")
			urlname = rs("urlname")
			sort_template = rs("sort_template")
			content_template = rs("content_template")
			position = rs("position")
			content = rs("content")
			pic = rs("pic")
			rs.close
			set rs = nothing
		readerByUrl = true
	end function

	'获取post'
	sub postData()
		id = e.postInt("id")
		title = e.post("title")
		if title = "" then e.alert("请输入分类名称")
		pid = e.postInt("pid")
		weight = e.postInt("weight")
		urlname = e.post("urlname")
		sort_template = e.post("sort_template")
		content_template = e.post("content_template")
		content = e.post("content")
		if instr(content,"<img") = 0 then
			if e.htmlFilter(content) = "" then content = ""
		end if
		pic = e.post("pic")
	end sub

	'新建'
	function create()
		call postData()
		set rs = e.objRs
				rs.open "select top 1 * from " & table,db.conn,1,3
				rs.addnew()
				id = rs("id")
				call createPosition
				call createUrl
				rs("guid") = guid
				rs("title") = title
				rs("pid") = pid
				rs("weight") = weight
				rs("sort_template") = sort_template
				rs("content_template") = content_template
				rs("url") = url
				rs("urlname") = urlname
				rs("position") = position
				rs("content") = content
				rs("pic") = pic
				rs.update
				rs.close
		set rs = nothing
	end function

	'保存'
	function update()
			call postData()
			if id > 0 and pid > 0 then
				dim strParentPosition : strParentPosition = db.query("select [position] from "& table &" where id = " & pid)(0)
				if instr(strParentPosition,"," & id & ",") > 0 then e.alert("上级分类无效")
			end if
			set rs = e.objRs
			rs.open "select top 1 * from "& table &" where id =" & id,db.conn,1,3
					call createPosition
					call createUrl
					rs("guid") = guid
					rs("title") = title
					rs("pid") = pid
					rs("weight") = weight
					rs("sort_template") = sort_template
					rs("content_template") = content_template
					rs("url") = url
					rs("urlname") = urlname
					rs("position") = position
					rs("content") = content
					rs("pic") = pic
					rs.update
					rs.close
			set rs = nothing
			if e.post("changeChild") = "True" then
				db.exec("update "& table &" set [sort_template] = '"& sort_template &"',[content_template] = '"& content_template &"' where [pid] = " & id)
			end if
	end function

	'删除'
	function delete()
		id = e.getInt("id")
		if id > 0 then
			if hasChild then e.alert("请先删除该分类下的小类")
			delete = db.exec("delete from "& table &" where id in("& id &")")
		end if
	end function

	'读取分类列表'array
	function getRows()
		set rs = db.query("select * from "& table &" order by weight desc,id desc")
			getRows = rs.getrows()
		rs.close
		set rs = nothing
	end function

	'获取分类下拉菜单'string
	function getSelect()
		dim arrList : arrList = getRows()
		getSelect = getSelectLoop(arrList,0,0)
	end function

	'循环读取分类信息'
	private function getSelectLoop(arrList,intPid,intDeep)
		for i = 0 to ubound(arrList,2)
			if cint(intPid) = cint(arrList(2,i)) then
				getSelectLoop = getSelectLoop & "<option value="""& arrList(0,i) &""">" & getSelectDeep(intDeep) & arrList(1,i) & "</option>" & getSelectLoop(arrList,arrList(0,i),intDeep + 1)
			end if
		next
	end function

	'显示分类层级'
	private function getSelectDeep(intDeep)
		if intDeep = 0 then
			getSelectDeep = "┠ "
		else
			dim strTemp : strTemp = "┈"
			for i = 0 to intDeep
				strTemp = strTemp  & strTemp
			next
			getSelectDeep = "└" & strTemp
		end if
	end function

	'判断是否有小类'int
	function hasChild()
		hasChild = false
		intChild = db.query("select count(*) from "& table &" where pid = " & id)
		if intChild(0) > 0 then hasChild = true
	end function

	'所有小类id'numeric
	public function getChild()
		if position = "" then exit function
		set rs = db.query("select id from "& table &" where [position] like '"& position &"%' order by weight desc,id desc")
		if rs.eof then
			getChild = id
		else
			dim ids
			while not rs.eof
				ids = ids & rs(0) & ","
			rs.movenext
			wend
			getChild = ids
		end if
		rs.close
		set rs = nothing
	end function


	'修改分类名称'
	function ajaxTitle
		id = e.getInt("id")
		title = e.get("title")
		if id < 1 or title = "" then exit function
		ajaxTitle = db.exec("update "& table &" set [title] = '"& e.get("title") &"' where id = " & id)
	end function

	'修改分类排序'
	function ajaxWeight
		id = e.getInt("id")
		weight = e.getInt("weight")
		ajaxWeight = db.exec("update "& table &" set [weight] = " & weight & " where id = " & id)
	end function

	'设置路径'
	private sub createPosition()
		if pid > 0 then
			dim objRs : set objRs = db.query("select [position] from "& table &" where id = " & pid)
			if not objRs.eof then
				position = objRs(0) & id & ","
			end if
			objRs.close
			set objRs = nothing
		else
			position = "0," & id & ","
		end if
	end sub

	'设置url'
	private sub createUrl()
		if urlname <> "" then
			if instr(urlname,"/") > 0 then url = urlname : exit sub
			if not e.test("username",urlname) then e.alert("URL名称不合法，仅能包含数字和字母，且不能纯数字")
		end if
		url = CONFIG_HTML_SORT
		url = replace(url,"{id}",e.iif(urlname <> "",urlname,id))
		url = replace(url,"{ext}",CONFIG_HTML_EXT)
		if instr(url,"{parent}") > 0 then
			if  pid > 0 then
				dim oRs : set oRs = db.query("select [url] from "& table &" where id =" & pid)
				url = replace(url,"{parent}",oRs(0))
				set oRs = nothing
			else
				url = replace(url,"{parent}","?")
			end if
		end if
	end sub

end class
%>
