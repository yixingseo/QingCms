<%
dim news : set news = new ClassNews
class ClassNews

	public id,pid,title,seotitle,keywords,description,content,info,url,pic,att,show,insert_user,insert_date,urlname,guid,hits,weight
	public table
	dim rs

	sub class_initialize()
		id = 0
		weight = CONFIG_WEIGHT
		insert_date = now()
		insert_user = session("admin")
		hits = CONFIG_NEWS_HITS
		show = CONIFG_NEWS_SHOW
		guid = e.guid
		table = "[t_news]"
	end sub

	'读取'
	function read(intID)
		read = false
		if isNull(intID) or not isNumeric(intID) then exit function
		intID = cInt(intID)
		set rs = db.query("select top 1 * from "& table &" where id = " & intID)
		if rs.eof then exit function
			id = rs("id")
			pid = rs("pid")
			title = rs("title")
			seotitle = rs("seotitle")
			keywords =rs("keywords")
			description = rs("description")
			content = rs("content")
			info = rs("info")
			url = rs("url")
			pic = rs("pic")
			att = rs("att")
			show = rs("show")
			insert_user = rs("insert_user")
			insert_date = rs("insert_date")
			urlname = rs("urlname")
			guid = rs("guid")
			hits = rs("hits")
			weight = rs("weight")
			rs.close
			set rs = nothing
			read = true
	end function

	'读取'
	function readerByurl(sUrl)
		readerByurl = false
		set rs = db.query("select top 1 * from "& table &" where url = '"& sUrl &"'")
		if rs.eof then exit function
			id = rs("id")
			pid = rs("pid")
			title = rs("title")
			seotitle = rs("seotitle")
			keywords = rs("keywords")
			description = rs("description")
			content = rs("content")
			info = rs("info")
			url = rs("url")
			pic = rs("pic")
			att = rs("att")
			show = rs("show")
			insert_user = rs("insert_user")
			insert_date = rs("insert_date")
			urlname = rs("urlname")
			guid = rs("guid")
			hits = rs("hits")
			weight = rs("weight")
			rs.close
			set rs = nothing
		readerByurl = true
	end function

	'获取post'
	sub postData()
		id = e.postInt("id")
		pid = e.postInt("pid")
		title = e.post("title")
		if title = "" then e.alert("请输入标题")
		seotitle = e.post("seotitle")
		keywords = e.post("keywords")
		description = e.post("description")
		content = e.post("content")
		urlname = e.post("urlname")
		pic = e.post("pic")
		hits = e.postint("hits")
		weight = e.postint("weight")
		info = e.post("info")
	end sub

	'新建'
	function create()
		call postData()
		set rs = e.objRs
			rs.open "select top 1 * from " & table,db.conn,1,3
			rs.addnew()
			id = rs("id")
			call createUrl()
			rs("guid") = guid
			rs("insert_user") = insert_user
			rs("pid") = pid
			rs("title") = title
			rs("seotitle") = seotitle
			rs("keywords") = keywords
			rs("description") = description
			rs("content") = content
			rs("info") = info
			rs("url") = url
			rs("urlname") = urlname
			rs("pic") = pic
			rs("hits") = hits
			rs("weight") = weight
			rs("show") = show
			rs.update()
			rs.close
			set rs = nothing
	end function

	'修改'
	function update()
		call postData()			
		set rs = e.objRs
			rs.open "select top 1 * from "& table &" where id = " & id,db.conn,1,3
			call createUrl()
			rs("pid") = pid
			rs("title") = title
			rs("seotitle") = seotitle
			rs("keywords") = keywords
			rs("description") = description
			rs("content") = content
			rs("info") = info
			rs("url") = url
			rs("urlname") = urlname
			rs("pic") = pic
			rs("weight") = weight
			rs.update()
			rs.close
			set rs = nothing
	end function

	'删除'
	function delete()
		id = e.get("id")
		if not isnumeric(id) then exit function
		set rs = db.query("select pic from "& table &" where id in("& id &")")
		while not rs.eof
			if rs(0) <> "" then e.fileDelete(rs(0))
		rs.movenext
		wend
		rs.close
		set rs = nothing
		delete = db.exec("delete from "& table &" where id in("& id &")")
	end function

	'显示'
	function display()
		id = e.get("id")
		if not isnumeric(id) then exit function
		display = db.exec("update "& table &" set show = 0 where id in("& id &")")
	end function

	'隐藏/回收站'
	function rec()
		id = e.get("id")
		if not isnumeric(id) then exit function
		rec = db.exec("update "& table &" set show = -1 where id in("& id &")")
	end function

	'创建url'
	private sub createUrl()
		if urlname <> "" then
			if instr(urlname,"/") > 0 then url = urlname : exit sub
			if not e.test("username",urlname) then e.alert("URL名称不合法，仅能包含数字和字母，且不能纯数字")
		end if
		url = CONFIG_HTML_CONTENT
		url = replace(url,"{ext}",CONFIG_HTML_EXT)
		url = replace(url,"{id}",e.iif(urlname <> "",urlname,id))
		if pid > 0 then
			dim oRs : set oRs = db.query("select top 1 url from [t_sort] where id = " & pid)
			url = replace(url,"{parent}",oRs(0))
			set oRs = nothing
		else
			url = replace(url,"{parent}","")
		end if
	end sub

	'修改名称'
	function ajaxTitle()
		id = e.getInt("id")
		title = e.get("title")
		if id < 1 or title = "" then exit function
		ajaxTitle = db.exec("update "& table &" set title = '"& title &"' where id = " & id)
	end function

	'修改排序'
	function ajaxWeight()
		id = e.getint("id")
		weight = e.getint("weight")
		if id = 0 then exit function
		ajaxWeight = db.exec("update "& table &" set weight = "& weight &" where id = " & id)
	end function

	'删除特性'
	function delAtt()
		id = e.getInt("id")
		strAtt = e.get("att")
		if id < 1 or strAtt = "" then exit function
		if not read(id) then exit function
		if instr(att,strAtt) < 1 then exit function
		att = replace(att,strAtt & ",","")
		att = replace(att,",,",",")
		delAtt = db.exec("update "& table &" set att = '"& att &"' where id = " & id)
	end function

	'添加特性'
	function addAtt()
		id = e.getint("id")
		strAtt = e.get("att")
		if id = 0 or strAtt = "" then exit function
		if not read(id) then exit function
		if instr(att,strAtt & ",") > 0 then exit function
		if att = "" then att = strAtt & "," else att =  att & strAtt & ","
		addAtt = db.exec("update "& table &" set att = '"& att &"' where id = " & id)
	end function

	'点击'
	function hit()
		hit = db.exec("update "& table &" set hits = hits + 1 where id = " & id)
	end function

	property get insertDate
		insertdate = insert_date
	end property

	property let insertDate(d)
		if isdate(d) then insert_date = d
	end property

end class
%>
