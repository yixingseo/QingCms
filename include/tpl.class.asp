<%
class class_tpl

	dim rs,reg,isHome
	public html,source
	public thisUrl,nodeUrl,pageCurrent,searchKeywords
	public cachePath,cacheFile,isCache,cacheFilePath

	sub class_initialize()
		set reg = e.objReg
		cachePath = site.dic("root") & e.iif(CONFIG_CACHE_PATH <> "",CONFIG_CACHE_PATH,"cache/")
		isCache = e.iif(site.dic("cache")="true",true,false)
	end sub

	sub class_terminate()
		set reg = nothing
		set news = nothing
		set srot = nothing
	end sub

	'路由'
	sub route()
		thisUrl = lcase(request.Querystring)
		if thisUrl = "" then thisUrl = "index.html"
		cacheFile = replace(thisUrl,"/",CONFIG_HTML_TAG)
		if instr(cacheFile,CONFIG_HTML_EXT)	= 0 then cacheFile = cacheFile & CONFIG_HTML_EXT
		thisUrl = "?" & thisUrl
		call readerCache()
		if thisUrl = "?index.html" then call showIndex() : exit sub
		if instr(thisUrl,CONFIG_PAGELIST_TAG) > 0 then
			dim arrUrl : arrUrl = split(thisUrl,CONFIG_PAGELIST_TAG)
			thisUrl = e.regReplace(thisUrl,CONFIG_PAGELIST_TAG & "[0-9]+","")
			pageCurrent = replace(arrUrl(1),CONFIG_HTML_EXT,"")
			pageCurrent = replace(pageCurrent,"/","")
			if not isNumeric(pageCurrent) then pageCurrent = 1
			pageCurrent = cint(pageCurrent)
		end if
		if e.get("search") <> "" or e.get("keywords") <> "" then call showSearch() : exit sub
		if sort.readerByUrl(thisUrl) then call showSort() : exit sub
		if news.readerByurl(thisUrl) then call showNews() : exit sub
		if thisUrl = "?sitemap.xml" then call showSitemapXml() : exit sub
		if thisUrl = "?sitemap" then call showSitemap() : exit sub
		call showError()
	end sub

	'显示404错误页面'
	sub showError()
		html = getTemplate("404.htm")
		call display()
	end sub

	'显示sitemap'
	sub showSitemap()
		dim urlPath : urlPath = site.dic("domain")
		if right(urlpath,"1") = "/" then urlPath = left(urlPath,len(urlPath)-1)
		urlPath = urlPath & site.dic("root")
		set rs = db.query("select [url] from [t_sort] order by weight desc,id desc")
		while not rs.eof
			e.echo urlPath & rs(0) & vbcrlf
		rs.movenext
		wend
		rs.close
		set rs = nothing
		set rs = db.query("select [url] from [t_news] order by weight desc,id desc")
		while not rs.eof
			e.echo urlPath & rs(0) & vbcrlf
		rs.movenext
		wend
		rs.close
		set rs = nothing
	end sub

	'显示sitemap.xml'
	sub showSitemapXml()
		dim s,datenow
		dim urlPath : urlPath = site.dic("domain")
		if right(urlpath,"1") = "/" then urlPath = left(urlPath,len(urlPath)-1)
		urlPath = urlPath & site.dic("root")
		datenow = e.datetime(now(),"y-m-d")
		s = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf
		s = s & "<urlset>" & vbcrlf
		s = s & "<url>" & vbcrlf
		s = s & "<loc>" & urlPath & "</loc>" & vbcrlf
		s = s & "<lastmod>"& datenow &"</lastmod>" & vbcrlf
		s = s & "<changefreq>daily</changefreq>" & vbcrlf
		s = s & "<priority>1.0</priority>" & vbcrlf
		s = s & "</url>" & vbcrlf
		set rs = db.query("select url from t_sort")
		while not rs.eof
			s = s & "<url>" & vbcrlf
			s = s & "<loc>"& urlPath & rs("url") &"</loc>" & vbcrlf
			s = s & "<lastmod>"& datenow &"</lastmod>" & vbcrlf
			s = s & "<changefreq>daily</changefreq>" & vbcrlf
			s = s & "<priority>1.0</priority>" & vbcrlf
			s = s & "</url>" & vbcrlf
		rs.movenext
		wend
		rs.close
		set rs = nothing

		set rs = db.query("select url from t_news")
		while not rs.eof
			s = s & "<url>" & vbcrlf
			s = s & "<loc>"& urlPath & rs("url") &"</loc>" & vbcrlf
			s = s & "<lastmod>"& datenow &"</lastmod>" & vbcrlf
			s = s & "<changefreq>daily</changefreq>" & vbcrlf
			s = s & "<priority>1.0</priority>" & vbcrlf
			s = s & "</url>" & vbcrlf
		rs.movenext
		wend
		rs.close
		set rs = nothing

		s = s & "</urlset>" & vbcrlf
		e.die s
	end sub

	'读取缓存文件'
	sub readerCache()
		if site.dic("cache") <> "true" then exit sub
		cacheFilePath = cachePath & cacheFile
		if not e.fileExists(cacheFilePath) then exit sub
		html = e.fileRead(cacheFilePath)
		e.die html & "<!--"& cacheFile &"-->"
	end sub

	'写入缓存文件'
	sub writeCache()
		if not isCache then exit sub
		cacheFilePath = cachePath & cacheFile
		if not e.fileExists(cachePath) then call e.fileCreate(cachePath,1)
		call e.fileSave(html,cacheFilePath,1)
	end sub

	'显示模板'
	sub display()
		call replaceAsp()
		call writeCache()
		response.clear()
		dim endTimer : endTimer = timer()
		e.die html & "<!--载入耗时："& formatNumber((endTimer-starTimer)*1000,3) &"毫秒-->"
	end sub

	'显示首页'
	sub showIndex()
		html = getTemplate(CONFIG_TEMPLATES_INDEX)
		site.dic("meta") = site.getMeta(site.Dic("keywords"),site.dic("description"))
		isHome = true
		call qqOnline()
		call replaceDic("site",site.dic)
		call replaceLoops()
		call display()
	end sub

	'显示分类页面'
	sub showSort()
		html = getTemplate(sort.sort_template)
		'处理meta'
		dim strMeta : strMeta = ""
		call qqOnline()
		call replaceDic("site",site.dic)
		dim oDic : set oDic = e.objDic
			oDic.add "id",sort.id
			oDic.add "title",sort.title
			oDic.add "url",sort.url
			oDic.add "content",sort.content
			oDic.add "pic",sort.pic
		call replaceDic("sort",oDic)
		call replaceLoops()
		call display()
	end sub

	'显示内容页面'
	sub showNews()
		if not sort.read(news.pid) then e.die("未找到分类：ID/" & news.id)
		html = getTemplate(sort.content_template)
		site.dic("meta") = site.getMeta(news.keywords,news.description)
		call qqOnline()
		call replaceDic("site",site.dic)
		dim oDic : set oDic = e.objDic
			oDic.add "id",news.id
			oDic.add "pid",news.title
			oDic.add "seotitle",news.seoTitle
			oDic.add "title",news.title
			oDic.add "keywords",news.keywords
			oDic.add "description",news.description
			oDic.add "content",news.content
			oDic.add "url",news.url
			oDic.add "pic",news.pic
			oDic.add "att",news.att
			oDic.add "user",news.insert_user
			oDic.add "date",news.insert_date
			oDic.add "guid",news.guid
			oDic.add "hits",news.hits
			oDic.add "info",news.info
		'上一页'
		dim s : s = ""
			set rs = db.query("select top 1 [id],[title],[url] from [t_news] where [id] > "& news.id &" and [show] > -1 order by [weight] desc,[id] asc")
			if not rs.eof then
				s = "<a href="""& rs("url") &"""><span aria-hidden=""true"">&larr;</span>"& e.cutstring(rs("title"),20) &"</a>"
			else
				s = "<a href='#'>没有了</a>"
			end if
			rs.close
			set rs = nothing
			oDic.add "prev",s
		'下一页'
			set rs = db.query("select top 1 [id],[title],[url] from [t_news] where id < "& news.id &" and [show] > -1 order by [weight] desc,[id] desc")
			if not rs.eof then
				s = "<a href="""& rs("url") &""">"& e.cutstring(rs("title"),20) &"<span aria-hidden=""true"">&rarr;</span></a>"
			else
				s = "<a href='#'>没有了</a>"
			end if
			rs.close
			set rs = nothing
			oDic.add "next",s
		call replaceDic("news",oDic)
		set oDic = nothing

		'内容的分类'
		set oDic = e.objDic
			oDic.add "id",sort.id
			oDic.add "title",sort.title
			oDic.add "url",sort.url
			oDic.add "content",sort.content
			oDic.add "pic",sort.pic
		call replaceDic("sort",oDic)
		call replaceLoops()
		set oDic = nothing
		news.hit()
		set news = nothing
		call display()
	end sub

	'在线QQ代码'
	sub qqOnline()
		dim tempStr : tempStr = ""
		if site.dic("online") <> "" then
			tempStr = tempStr & "<div class='qqonline' style='top: 150px;'><ul>"
			dim qqArr : qqArr = split(site.dic("online"),",")
			for i = 0 to ubound(qqArr)
				if isNumeric(qqArr(i)) then
				tempStr = tempStr & "<li><a target=""_blank"" href=""http://wpa.qq.com/msgrd?v=3&amp;uin="& qqArr(i) &"&amp;site=qq&amp;menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=2:"& qqArr(i) &":51""></a></li>"
				end if
			next
			tempStr = tempStr & "</ul></div>"
			tempStr = tempStr & "<script>window.onscroll=function(){var oDiv=$('.qqonline');var target=150;var scrollTop=document.documentElement.scrollTop||document.body.scrollTop;var iTarget=scrollTop+target;oDiv.stop(true).animate({top:iTarget+'px'},500)};</script>"
			site.dic.add "qq",tempStr
		end if
	end sub

	'显示搜索页面'
	sub showSearch()
		html = getTemplate("search.htm")
		call replaceDic("site",site.dic)
		call replaceLoops()
		call display()
	end sub

	'处理搜索页面'
	sub replaceSearch(ma)
		searchKeywords = e.get("search")
		if searchKeywords = "" then searchKeywords = e.get("keywords")
		searchKeywords = e.safe(searchKeywords)
		html = replace(html,"{#search.keywords}",searchKeywords)
		dim pagesize
		if ma.submatches(1) <> "" then
			pagesize = getAtt(ma.submatches(1),"size")
		end if
		if isNumeric(pagesize) then
			pagesize = cint(pagesize)
		else
			pagesize = 20
		end if
		if not isNumeric(page_current) then
			page_current = 1
		else
			page_current = cint(page_current)
		end if
		if page_current < 1 then page_current = 1
		dim sql
		sql = "select [id],[title],[pic],[url],[insert_date] as [date] from [t_news] where [show] > -1 and [title] like '%"& searchKeywords &"%' order by [weight] desc,[id] desc"
		dim rs,strLoophtml
		i = 0
		set rs = e.objRs
			rs.open sql,db.conn,1,1
		if not rs.eof then
			rs.pagesize = pagesize
			rs.absolutepage = page_current
			html = replace(html,"{#pagelist}",getPager(rs.pagecount,rs.absolutepage))
			while not rs.eof and i < rs.pagesize
				strLoophtml = strLoophtml & replaceFields(rs,ma.submatches(0),ma.submatches(2))
			i = i + 1
			rs.movenext
			wend
			html = replace(html,ma,strLoophtml)
		else
			html = replace(html,ma,"SORRY,没有搜索到与 <em>"& q_keywords &"</em> 相关的内容!")
		end if
	end sub

	'替换循环标签'
	sub replaceLoops()
		dim ma,mas
		set mas = e.regExecute(html,"\{#([a-zA-Z]+)+\s([^\}]*)\}([\s\S]*?)\{/\#\1\}")
		for each ma in mas
			if ma.submatches(0) = "page" then
				call replacePagelist(ma)
			elseif ma.submatches(0) = "search" then
				call replaceSearch(ma)
			else
				call replaceLoopsHtml(ma)
			end if
			call replaceLoops()
		next
		set mas = nothing
	end sub

	'替换循环标签里面的内容'
	sub replaceLoopsHtml(ma)
		dim strSql : strSql = getLoopSql(ma)
		if strSql = "" then html = replace(html,ma.value,"") : exit sub
		dim strTag,strAtt,strFields
			strTag = ma.subMatches(0)
			strAtt = ma.subMatches(1)
			strFields = ma.subMatches(2)
		set rs = db.query(strSql)
		if rs.eof then
			set mas = e.regExecute(ma,"<\!--\s?"& strTag &"\.empty\s?-->([\s\S]*?)<!--\s?\/"& strTag &"\.empty\s?-->")
			if mas.count > 0 then
				html = replace(html,ma,mas(0).value) : exit sub
			else
				html = replace(html,ma,"<!--empty-->") : exit sub
			end if
		end if
		dim i,strLoophtml
		i = 0
		while not rs.eof
			strLoophtml = strLoophtml & replaceFields(rs,strTag,strFields)
			rs.movenext
		wend
		rs.close
		set rs = nothing
		html = replace(html,ma.value,strLoophtml)
	end sub

	'替换循环标签字段'
	function replaceFields(objRs,strTag,strFields)
		dim objMa,objMas,strRowhtml
		strRowhtml = strFields
		set objMas = e.regExecute(strFields,"\["& strTag &"\.([a-z]+)(.*?)\]")
		for each objMa in objMas
			dim strName,strValue
				strName = lcase(objMa.submatches(0))
				strValue = ""
			if strName = "i" then
				strValue = rs.AbsolutePosition - 1
			elseif strName = "num" then
				strValue = rs.AbsolutePosition
			else
				on error resume next '防止不存在的字段报错'
				strValue = e.iif(isNull(objRs(strName).value),"",objRs(strName).value)
				if objMa.submatches(1) <> "" then
					strValue = formatTag(objMa.submatches(1),strValue)
				end if
				if err then err.clear
			end if
			strRowhtml = replace(strRowhtml,objMa,strValue)
		next
		replaceFields = strRowhtml
	end function

	'获取格式化的值'
	function formatTag(strAtts,strVal)

		'截取字符串'
		dim intLen : intLen = getAtt(strAtts,"len")
		if not isEmpty(intLen) then
			intLen = e.strToInt(intLen)
			if intLen > 0 then formatTag = e.cutstring(strVal,intLen)
			exit function
		end if

		'日期格式化'
		dateTemplate = getAtt(strAtts,"datetime")
		if not isEmpty(dateTemplate) then
			formatTag = e.datetime(strVal,dateTemplate)
		end if
	end function

	'获取循环标签的Sql值'
	function getLoopSql(strAtts)
		if getAtt(strAtt,"sql") <> "" then getLoopSql = getAtt(strAtts,"sql") : exit function
		dim strTable,strFields,strOrderby,strWhere
		dim intPid,intRid,strKeywords,strNewsatt,intNum
		strTable = getAtt(strAtts,"data")
		if strTable = "" then strTable = "news"
		select case lcase(strTable)
			case "news"
				strTable = "[t_news]"
				strFields = "[id],[title],[info],[url],[seotitle],[pic],[hits],[insert_date] as [date]"
				strWhere = "[show] > -1 "
				intPid = getAtt(strAtts,"pid")
				if isEmpty(intPid) then
					intRid = getAtt(strAtts,"rid")								
				end if
				strKeywords = getAtt(strAtts,"keywords")
				if not isEmpty(strKeywords) then strWhere = strWhere & " and keywords like '%"& strKeywords &"%'"				
				strNewsAtt = getAtt(strAtts,"att")
				if not isEmpty(strNewsatt) then strWhere = strWhere & " and instr(att,'"& strNewsatt &"') > 0"
				
			case "sort"
				strTable = "[t_sort]"
				strFields = "[id],[title],[url]"
				strWhere = "[id] > 0 "
				intPid = getAtt(strAtts,"pid")
				if isEmpty(intPid) then
					intRid = getAtt(strAtts,"rid")
				end if
				
			case "link"
				strTable = "[t_link]"
				strFields = "[id],[title],[url],[target],[pic]"
				strWhere = "[id] > 0 "
		end select
			
		'pid'
		if not isEmpty(intPid) then
			if inStr(intPid,",") > 0 then
				strWhere = strWhere & " and pid in("& intPid &")"
			else
				intPid = Cint(intPid)
				strWhere = strWhere & " and pid = " & intPid
			end if
		end if

		'rid'
		if not isEmpty(intRid) and isNumeric(intRid) then
			intRid = Cint(intRid)
			if intRid > 0 then
				dim oSort : set oSort = new ClassSort
				if oSort.read(intRid) then
					intRid = oSort.getChild()
				end if
				if err then e.die err.description
				set oSort = nothing
				strWhere = strWhere & " and pid in("& intRid &") "
			end if
		end if

		'top'
		intNum = getAtt(strAtts,"num")
		if isEmpty(intNum) then 
			intNum = 10
		else
			intNum = e.strToInt(intNum)
		end if

		'order by'
		strOrderby = getAtt(strOrderby,"orderby")
		if isEmpty(strOrderby) then
			strOrderby = " [weight] desc,[id] desc"
		end if

		getLoopSql = "select top "& intNum & " " & strFields &" from "& strTable &" where "& strWhere &" order by " & strOrderby
	end function

	'获取属性值'
	function getAtt(strAtts,strAttName)
		dim mas
		set mas = e.regExecute(strAtts,lcase(strAttName) & "=""(.*?)""")
		if mas.count > 0 then getAtt = mas(0).submatches(0)
		set mas = nothing
	end function

	'替换标签(标签头，字典)
	sub replaceDic(tag,objDic)
		dim ma,mas,tmp,strValue
		set mas = e.regExecute(html,"\{#"& tag &"\.([-a-z_\.]+)(.*?)\}")
		for each ma in mas
			if not objDic.exists(ma.submatches(0)) then
				strValue = ""
			else
				strValue = objDic.item(ma.submatches(0))
				if isnull(strValue) then strValue = ""
				if ma.subMatches(1) <> "" then
					strValue = formatTag(ma.submatches(1),strValue)
				end if
			end if
			'e.echoline ma & "-" & strValue & objDic("title")
			html = replace(html,ma,strValue)
		next
		set mas = nothing
		set objDic = nothing
	end sub

	'替换tags'
	sub replaceTags()
		set rs = db.query("select * from [t_tag]")
		dim oDic : set oDic = e.objDic
		while not rs.eof
			oDic.add rs("title").value,e.decode(rs("content").value)
		rs.movenext
		wend
		rs.close
		set rs = nothing
		call replaceDic("tag",oDic)
	end sub

	'替换分页'
	sub replacePagelist(ma)
		dim pagesize
		if ma.submatches(1) <> "" then
			pagesize = getAtt(ma.submatches(1),"size")
		end if
		if isNumeric(pagesize) then
			pagesize = cint(pagesize)
		else
			pagesize = 20
		end if
		if not isNumeric(pageCurrent) then
			pageCurrent = 1
		else
			pageCurrent = cint(pageCurrent)
		end if
		if pageCurrent < 1 then pageCurrent = 1
		dim sql
		sql = "select [id],[title],[pic],[url],[insert_date] as [date] from [t_news] where [show] > -1 and [pid] in("& sort.getChild &") order by [weight] desc,[id] desc"
		dim rs,strLoophtml
		i = 0
		set rs = e.objRs
			rs.open sql,db.conn,1,1
		if not rs.eof then
			rs.pagesize = pagesize
			rs.absolutepage = pageCurrent
			html = replace(html,"{#pagelist}",getPager(rs.pagecount,rs.absolutepage))
			'call set_pager(rs.pagecount,rs.absolutepage)
			while not rs.eof and i < rs.pagesize
				strLoophtml = strLoophtml & replaceFields(rs,ma.submatches(0),ma.submatches(2))
			i = i + 1
			rs.movenext
			wend
		end if
		html = replace(html,ma,strLoophtml)
		html = replace(html,"{#pagelist}","")
	end sub

	'获取分页条'
	private function getPager(intPagecount,intPagecurrent)
		pageurl = replace(thisUrl,CONFIG_HTML_EXT,"")
		dim s
			s = "<ul class='pagination'>"
		if intpagecurrent > 1 then
			s = s & "<li><a href='"& pageurl & CONFIG_PAGELIST_TAG & intpagecurrent - 1 & CONFIG_HTML_EXT &"'>上一页</a></li>"
		else
			s = s & "<li class='disabled'><a aria-label='previous'><span aria-hidden='true'>上一页</span></a></li>"
		end if
		dim intpagei
		for intpagei = 1 to intpagecount
			if intpagei = intpagecurrent then
				s = s & "<li class='active'><a>"& intpagei &" </a></li>"
			else
				s = s & "<li><a href='"& pageurl & CONFIG_PAGELIST_TAG & intpagei & CONFIG_HTML_EXT &"'>" & intpagei & "</a></li>"
			end if
		next
		if intpagecurrent < intpagecount then
			s = s & "<li><a href='"& pageurl & CONFIG_PAGELIST_TAG & pageCurrent + 1 & CONFIG_HTML_EXT & "'>下一页</a></li>"
		else
			s = s & "<li class='disabled'><a aria-label='next'><span aria-hidden='true'>下一页</span></a></li>"
		end if
		getPager = s & "</ul>"
	end function

	'模板目录'
	property get templatePath
		templatePath = site.dic("root") & CONFIG_TEMPLATES
	end property

	'读取模板文件'
	function getFile(strPath)
		if isNull(strPath) or strPath = "" or not e.fileExists(templatePath & strPath) then
			e.die "未找到模板文件：" & templatePath & strPath
			exit function
		end if
		getFile = e.fileRead(templatePath & strPath)
	end function

	'加载模板文件'
	function getTemplate(strPath)
		html = getFile(strPath)
		source = html
		call replaceInclude()	'替换模板中的包含文件'
		call replaceSrc()	'替换模板中的附件路径'
		call replaceTags()	'替换tag标签'
		getTemplate = html
	end function

	'替换包换文件'
	sub replaceInclude()
		dim ma,mas
		set mas = e.regExecute(html,"<\!--.?#include[ ]+?file[ ]*?=[ ]*?""(\S+?)""[ ]*--\>")
		for each ma in mas
			html = replace(html,ma.value,getFile(ma.submatches(0)))
		next
		set mas = nothing
	end sub

	'替换模板附件目录 images,css,js,bootstrap'
	sub replaceSrc()
		reg.pattern = "<(.*?)(src=|href=|value=|background=)""(images/|css/|js/|bootstrap/)(.*?)""(.*?)>"
		html = reg.replace(html,"<$1$2""" & templatePath  & "$3$4""$5>")
	end sub


	'执行asp语句'
	sub replaceAsp()
		dim s,tmp, a, b, t, mas, m
		s = html
		set mas = e.regExecute(s,"\{if\s+(.*)?\}")
		if mas.count > 0 then
			s = e.regReplace(s,"\{if\s+(.*?)\}","<"&"%if $1 then%"&">")
			s = e.regReplace(s,"\{else\}","<"&"%else%"&">")
			s = e.regReplace(s,"\{\/if\}","<"&"%end if%"&">")
			s = e.regReplace(s,"\{echo\s+(.*?)\}","<"&"%response.write $1%"&">")
		end if
		dim str : str = s
		'判断是否首页'
		tmp = "ishome = " & e.iif(isHome,true,false) & vbcrlf
		tmp = tmp & "thisurl = """& e.iif(thisUrl<>"",thisUrl,"") &""" " & vbcrlf
		tmp = tmp & "dim htm : htm = """"" & vbcrlf
		a = 1
		b = instr(a,str,"<%") + 2
		while b > a + 1
			t = mid(str,a,b-a-2)
			t = replace(t,vbcrlf,"{::vbcrlf}")
			t = replace(t,vbcr,"{::vbcr}")
			t = replace(t,"""","""""")
			tmp = tmp & "htm = htm & """ & t & """" & vbcrlf
			a = instr(b,str,"%\>") + 2
			tmp = tmp & e.regReplace(mid(str,b,a-b-2),"^\s*=","htm = htm & ") & vbcrlf
			b = instr(a,str,"<%") + 2
		wend
		t = mid(str,a)
		t = replace(t,vbcrlf,"{::vbcrlf}")
		t = replace(t,vbcr,"{::vbcr}")
		t = replace(t,"""","""""")
		tmp = tmp & "htm = htm & """ & t & """" & vbcrlf
		tmp = replace(tmp,"response.write","htm = htm & ",1,-1,1)
		on error resume next
		executeGlobal(tmp)
		if err then html = s : exit sub
		htm = replace(htm,"{::vbcrlf}",vbcrlf)
		htm = replace(htm,"{::vbcr}",vbcr)
		html = htm
	end sub
end class
%>
