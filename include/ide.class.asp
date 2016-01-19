<%
'通用类'
'2015-11-27'

dim e : set e = new ClassIDE
class ClassIDE

	'打印字符串'
	sub echo(s)
		response.write(s)
	end sub

	'打印字符串+换行'
	sub echoLine(s)
		response.write(s) & "<br>"
	end sub

	'打印字符串+结束'
	sub die(s)
		response.write(s) : response.end()
	end sub

	'弹出提示框'
	sub alert(s)
		die "<script type='text/javascript'>alert('"& s &"');history.go(-1);</script>"
	end sub

	'弹出提示框+跳转(提示信息,跳转网址)'
	sub alertUrl(s,surl)
		die "<script type='text/javascript'>alert('"& s &"');window.location.href='"& surl &"';</script>"
	end sub

	'三元表达式'
	function iif(byval a,byval b,byval c)
		if a then iif = b else iif = c
	end function

	'guid'string
	function guid()
		dim typelib
	    set typelib = server.createobject("scriptlet.typelib")
	    guid = mid(typelib.guid,2,36)
	    set typelib = nothing
	end function

	'随机字符串''string
	property get getRand()
		dim y,m,d,h,mm,s,r
		randomize
		y = year(now)
		m = right("0" & month(now),2)
		d = right("0" & day(now),2)
		h = right("0" & hour(now),2)
		mm = right("0" & minute(now),2)
		s = right("0" & second(now),2)
		r = 0
		r = cint(rnd() & 10000)
		s = right("0000" & r,4)
		getrand = y & m & d& h & mm & s & r
	end property

	'获取IP'string
	function getIP()
		s = request.serverVariables("http_x_forwarded_for")
		if s = "" then s = request.serverVariables("remote_addr")
		getIP = s
	end function

	'字符串转数字'int
	function strToInt(s)
		if isnull(s) then strToInt = 0 : exit function
		if s = "" then strToInt = 0 : exit function
		if not isNumeric(s) then strToInt = 0 : exit function
		strtoint = cint(s)
	end function

	'querystring' int
	function getInt(s)
		getint = strtoint([get](s))
	end function

	function getID(s)
		getid = getint(s)
	end function

	'form'int
	function postInt(s)
		postint = strtoint(post(s))
	end function

	'get'string
	function [get](s)
		[get] = request.Querystring(s)
	end function

	'post'string
	function post(s)
		post = request.Form(s)
	end function

	'过滤危险字符串'string
	function safe(s)
		safe = replace(s,"'","''")
	end function

	'获取application'string
	function getApp(s)
		if s = "" then exit function
		getApp = application.Contents.item(s)
	end function

	'设置application'
	sub setApp(sName,sValue)
		application.Lock()
		application.Contents(sName) = sValue
		application.Unlock()
	end sub

	'删除application'
	sub removeApp(s)
		if s = "" then exit sub
		application.lock()
		application.contents.remove(s)
		application.unlock()
	end sub

	'清空所有application'
	sub clearApp()
		application.lock()
		application.contents.removeall()
		application.unlock()
	end sub

	'获取当前url'string
	function getUrl()
		dim strTemp : strTemp = ""
		strTemp = request.serverVariables("script_name")
		strTemp = strTemp & iif(request.Querystring <> "","?" & request.Querystring,"")
		getUrl = enCode(strTemp)
		'url = server.urlencode(url)
	end function

	'分割 字符串,长度'string
	function cutString(s,intLen)
		cutString = s
		if isnull(s) then exit function
		if s = "" then exit function
		if not isnumeric(intLen) then exit function
		if len(s)*2 < cint(intLen) then exit function
		s = replace(s,chr(10),"")
		dim l,t,c,i,d
		l = len(s)
		t = 0
		d = "…"
		for i = 1 to l
			c = abs(ascw(mid(s,i,1)))
			t = iif(c>225 , t + 2 , t + 1)
			if t >= intLen then
				cutString = left(s,i) & d
				exit for
			else
				cutString = s
			end if
		next
	end function

	'格式化日期(日期,y-m-d)'string
	function dateTime(dDate,strTemp)
		if not isdate(dDate) or strTemp = "" then
			datetime = dDate
			exit function
		end if

		dim y,m,d,h,mm,s,tmp
		tmp = strTemp
		tmp = replace(tmp,"y",year(dDate))
        tmp = replace(tmp,"Y",right(year(dDate),2))
        tmp = replace(tmp,"m",month(dDate))
        tmp = replace(tmp,"M",right("00" & month(dDate),2))
        tmp = replace(tmp,"d",day(dDate))
        tmp = replace(tmp,"D",right("00" & day(dDate),2))
        tmp = replace(tmp,"h",hour(dDate))
        tmp = replace(tmp,"H",right("00" & day(dDate),2))
        tmp = replace(tmp,"mi",minute(dDate))
        tmp = replace(tmp,"MI",right("00" & minute(dDate),2))
        tmp = replace(tmp,"s",second(dDate))
        tmp = replace(tmp,"S",right("00" & second(dDate),2))
        datetime = tmp
	end function

	'获取+处理url
	'geturl("-page") 获取不包含page参数的url
	'geturl("page") 获取只包含page参数的url
	function get_url(s)
		dim s_url,s_pas
			s_url = request.servervariables("script_name")
			s_pas = request.querystring
		if s = "" or s_pas = "" then
			get_url = s_url
		else
			if left(s,1) = "-" then
				s_pas = "&" & s_pas
				s_pas = regREplace(s_pas,"[&]{1}"& right(s,len(s)-1) &"=[^&]*","")
				s_pas = replace(s_pas,"&","?",1,1)
				get_url = s_url & s_pas
			else
				dim mas : set mas = reg_execute(s_pas,s &"=[^&]*")
				get_url = s_url & "?" & iif(mas.count > 0,mas(0).value,s_pas)
			end if
		end if
	end function

	'判断是否安装了组件'html
	function isInstall(s)
		on error resume next
		dim objtest
		set objtest = server.createobject(s)
		if err.number = 0 then
			isInstall = "<i class='fa fa-check'></i>"
		else
			isInstall  ="<i class='fa fa-question'></i>"
		end if
		set objtest = nothing
	end function

	'正则验证类型(类型,字符串)'bool
	function test(sType,s)
		dim pa
		select case sType
			case "date"		test = iif(isdate(str),true,false) : exit function
			case "english"	pa = "^[a-za-z]+$"
			case "chinese"	pa = "^[\u0391-\uffe5]+$"
			case "username"	pa = "^[a-z]\w{2,19}$"
			case "email"	pa = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
			case "int"		pa = "^[-\+]?\d+$"
			case "number"	pa = "^\d+$"
			case "double"	pa = "^[-\+]?\d+(\.\d+)?$"
			case "price"	pa = "^\d+(\.\d+)?$"
			case "zip"		pa = "^[1-9]\d{5}$"
			case "qq"		pa = "^[1-9]\d{4,9}$"
			case "phone"	pa = "^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$"
			case "mobile"	pa = "^((\(\d{2,3}\))|(\d{3}\-))?(1[35][0-9]|189)\d{8}$"
			case "url"		pa = "^(http|https|ftp):\/\/[a-za-z0-9]+\.[a-za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\""])*$"
			case "domain"	pa = "^[a-za-z0-9\-]+\.([a-za-z]{2,4}|[a-za-z]{2,4}\.[a-za-z]{2})$"
			case "ip"		pa = "^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$"
			case else		pa = s
		end select
		test = regtest(s,pa)
	end function

	'压缩字符串，过滤换行'string
	function htmlZip(s)
		s = trim(s)
		s = replace(s,vbcrlf,"")
		s = regREplace(s,"\s"," ")
		s = regREplace(s,"  "," ")
		htmlZip = s
	end function

	'html编码'string
	function htmlEncode(s)
		s = replace(s, chr(38), "&#38;")
		s = replace(s, "<", "&lt;")
		s = replace(s, ">", "&gt;")
		s = replace(s, chr(39), "&#39;")
		s = replace(s, chr(32), "&nbsp;")
		s = replace(s, chr(34), "&quot;")
		s = replace(s, chr(9), "&nbsp;&nbsp; &nbsp;")
		s = replace(s, vbcrlf, "<br />")
		htmlEncode = s
	end function

	function encode(s)
		encode = s
		if isnull(s) or s = "" then exit function
		encode = server.htmlEncode(s)
	end function

	'html解码'string
	function htmlDecode(s)
		if isnull(s) then exit function
		s = replace(s,"<br\s*/?\s*>", vbcrlf)
		s = replace(s, "&nbsp;&nbsp; &nbsp;", chr(9))
		s = replace(s, "&quot;", chr(34))
		s = replace(s, "&nbsp;", chr(32))
		s = replace(s, "&#39;", chr(39))
		s = replace(s, "&apos;", chr(39))
		s = replace(s, "&gt;", ">")
		s = replace(s, "&lt;", "<")
		s = replace(s, "&amp;", chr(38))
		s = replace(s, "&#38;", chr(38))
		htmlDecode = s
	end function

	function decode(s)
		if isnull(s) then exit function
		decode = htmlDecode(s)
	end function

	'过滤html标签'string
	function htmlFilter(byval s)
		s = regReplace(s,"<[^>]+>","")
		s = replace(s, ">", "&gt;")
		htmlFilter = replace(s, "<", "&lt;")
	end function

	'正则测试(字符串，规则)'bool
	function regTest(a,b)
		dim oReg : set oReg = objReg
			oReg.pattern = b
			regTest = oReg.test(a)
		set oReg = nothing
	end function

	'正则搜索(字符串，规则)'obj-maches
	function regExecute(a,b)
		dim oReg : set oReg = objReg
			oReg.pattern = b
			set regExecute = oReg.execute(a)
		set oReg = nothing
	end function

	'正则替换(原字符串，规则，替换字符串)'string
	function regReplace(byval a,byval b,byval c)
		dim oReg : set oReg = objReg
		    oReg.pattern = b
			regReplace = oReg.replace(a,c)
		set oReg = nothing
	end function

	'regexp对象 'obj-regexp
	function objReg()
		set objReg = new regexp : objReg.ignorecase = true : objReg.global = true : objReg.multiline = true
	end function

	'dictionary对象'obj-dictionary
	function objDic()
		set objDic = server.createobject("scripting.dictionary")
	end function

	'recordset对象'obj-recordset
	function objRs()
		set objRs = server.createobject("adodb.recordset")
	end function

	'fso对象'obj-recordset
	function objFso()
		set objFso = server.createobject("sc" & "rip" & "ting" & ".fi" & "les" & "yste" & "mobj" & "ect")	'防止被判断为病毒
	end function

	'函数：获取当前脚本执行文件所在的磁盘目录
	function file_dir()
		dim tmp, arr
		tmp = request.ServerVariables("SCRIPT_NAME")
		arr = split(tmp,"/")
		tmp = arr(ubound(arr))
		arr = split(server.MapPath(tmp),"\")
		file_dir = arr(ubound(arr)-1)
	end function

	'函数：检测文件是否存在
	function fileExists(path)
		dim tmp : tmp = false
		if path = "" then exit function
		dim fso :set fso = objFso
		if fso.fileexists(server.MapPath(path)) then tmp = true
		if fso.folderexists(server.MapPath(path)) then tmp = true
		set fso = nothing
		fileExists = tmp
	end function

	'函数：删除文件/文件夹
	function fileDelete(path)
		dim tmp : tmp = false
		if path = "" then exit function
		dim fso : set fso = objFso
		if fso.fileexists(server.MapPath(path)) then'目标是文件
			fso.deletefile(server.MapPath(path))
			if not fso.fileexists(server.MapPath(path)) then tmp = true
		end if
		if fso.folderexists(server.MapPath(path)) then'目标是文件夹
			fso.deletefolder(server.MapPath(path))
			if not fso.folderexists(server.MapPath(path)) then tmp = true
		end if
		set fso = nothing
		fileDelete = tmp
	end function

	'函数：获取文件/文件夹信息
	function fileInfo(path)
		dim tmp(4)
		dim fso :set fso = objFso
		if fso.fileexists(server.MapPath(path)) then '目标是文件
			dim fl : set fl = fso.getfile(server.MapPath(path))
			tmp(0) = fl.type'类型
			tmp(1) = fl.attributes'属性
			tmp(2) = csize(fl.size,4)'大小
			tmp(3) = fl.datecreated'创建时间
			tmp(4) = fl.datelastmodified'最后修改时间
		elseif fso.folderexists(server.MapPath(path)) then '目标是文件夹
			dim fd : set fd = fso.getfolder(server.MapPath(path))
			tmp(0) = "folder"'类型
			tmp(1) = fd.attributes'属性
			tmp(2) = csize(fd.size,4)'大小
			tmp(3) = fd.datecreated'创建时间
			tmp(4) = fd.datelastmodified'最后修改时间
		end if
		set fso = nothing
		fileInfo = tmp
	end function

	'函数：复制文件/文件夹
	function fileCopy(file_start,file_end,model)
		if model<>0 and model<>1 then model = false else model = cbool(model)
		dim tmp : tmp = false
		dim fso :set fso = objFso
		if fso.fileexists(server.MapPath(file_start)) then '目标是文件
			fso.copyfile server.MapPath(file_start),server.MapPath(file_end),model
			if fso.fileexists(server.MapPath(file_end)) then tmp = true
		end if
		if fso.folderexists(server.MapPath(file_start)) then '目标是文件夹
			fso.copyfolder server.MapPath(file_start),server.MapPath(file_end),model
			if fso.folderexists(server.MapPath(file_end)) then tmp = true
		end if
		set fso = nothing
		fileCopy = tmp
	end function

	'函数：创建文件夹
	function fileCreate(path,model)
		if model<>0 and model<>1 then model = false else model = cbool(model)
		dim tmp : tmp = false
		dim fso :set fso = objFso
		if fso.folderexists(server.MapPath(path)) then
			if model then fso.deletefolder(server.MapPath(path)) : fso.createfolder server.MapPath(path)
		else
			fso.createfolder server.MapPath(path)
		end if
		if fso.folderexists(server.MapPath(path)) then tmp = true
		set fso = nothing
		fileCreate = tmp
	end function

	'函数：获取指定目录下所有文件及文件夹列表
	function fileList(path)
		if not fileExists(path) then fileList = array("") : exit function
		dim fso :set fso = objFso
		dim fdr : set fdr = fso.getfolder( server.MapPath(path) )
		dim files : set files = fdr.files
		for each f in files
			tmp = tmp & t & f.name : t = "|"
		next
		set fso = nothing
		fileList = split(tmp,"|")
	end function

	'函数：返回图片类型及尺寸
	function fileImgInfo(path)
		dim tmp : tmp = array("",0,0)
		dim fso : set fso = objFso
		if fso.fileexists(server.MapPath(path)) then
			dim img : set img = loadpicture(server.MapPath(path))
			select case img.type
				case 0 : tmp(0) = "none"'类型
				case 1 : tmp(0) = "bitmap"
				case 2 : tmp(0) = "metafile"
				case 3 : tmp(0) = "ico"
				case 4 : tmp(0) = "win32-enhanced metafile"
			end select
			tmp(1) = round(img.width/26.4583)'宽度
			tmp(2) = round(img.height/26.4583)'高度
			set img = nothing
			set fso = nothing
		end if
		fileImgInfo = tmp
	end function

	'函数：检测图片文件合法性
	function fileIsImg(path)
		dim tmp : tmp = false
		if not fileExists(path) then fileIsImg = tmp : exit function
		dim jpg(1):jpg(0)=cbyte(&HFF):jpg(1)=cbyte(&HD8)
		dim bmp(1):bmp(0)=cbyte(&H42):bmp(1)=cbyte(&H4D)
		dim png(3):png(0)=cbyte(&H89):png(1)=cbyte(&H50):png(2)=cbyte(&H4E):png(3)=cbyte(&H47)
		dim gif(5):gif(0)=cbyte(&H47):gif(1)=cbyte(&H49):gif(2)=cbyte(&H46):gif(3)=cbyte(&H39):gif(4)=cbyte(&H38):gif(5)=cbyte(&H61)
		dim fstream,fext,stamp,i
		fext = mid(path, instrrev(path,".")+1)
		set fstream = server.CreateObject("ADODB.Stream")
		fstream.open
		fstream.type = 1
		fstream.loadfromfile server.MapPath(path)
		fstream.position = 0
		select case fext
			case "jpg","jpeg":
				stamp = fstream.read(2)
				for i=0 to 1
					if ascb(midb(stamp,i+1,1))=jpg(i) then tmp=true else tmp=false
				next
			case "gif":
				stamp = fstream.read(6)
				for i=0 to 5
					if ascb(midb(stamp,i+1,1))=gif(i) then tmp=true else tmp=false
				next
			case "png":
				stamp = fstream.read(4)
				for i=0 to 3
					if ascb(midb(stamp,i+1,1))=png(i) then tmp=true else tmp=false
				next
			case "bmp":
				stamp = fstream.read(2)
				for i=0 to 1
					if ascb(midb(stamp,i+1,1))=bmp(i) then tmp=true else tmp=false
				next
		end select
		fstream.close : set fstream = nothing
		fileIsImg = tmp
	end function

	'函数：采集远程文件并保存到本地磁盘
	function file_savefromurl(fileurl,savepath,savetype)
		if savetype<>1 and savetype<>2 then savetype=2
		dim xmlhttp : set xmlhttp = server.CreateObject("MSXML2.XMLHTTP")
		with xmlhttp
			.open "get", fileurl, false, "", ""
			.send()
			dim fl : fl = .responsebody
		end with
		set xmlhttp = nothing
		dim stream : set stream = server.CreateObject("ADODB.Stream")
		with stream
			.type = savetype
			.open
			.write fl
			'if savetype=1 then .write fl else .writetext fl
			.savetofile server.MapPath(savepath), 2
			.cancel()
			.close()
		end with
		set stream = nothing
		file_savefromurl = fileExists(savepath)
	end function

	'函数：读取文件内容到字符串
	function fileRead(path)
		dim tmp : tmp = "false"
		if not fileExists(path) then fileRead = tmp : exit function
		dim stream : set stream = server.CreateObject("ADODB.Stream")
		with stream
			.type = 2 '文本类型
			.mode = 3 '读写模式
			.charset = "utf-8"
			.open
			.loadfromfile(server.MapPath(path))
			tmp = .readtext()
		end with
		stream.close : set stream = nothing
		fileRead = tmp
	end function

	'函数：保存字符串到文件
	function fileSave(str,path,model)
		if model<>0 and model<>1 then model=1
		if model=0 and fileExists(path) then fileSave=true : exit function
		dim stream : set stream = server.CreateObject("ADODB.Stream")
		with stream
			.type = 2 '文本类型
			.charset = "utf-8"
			.open
			.writetext str
			.savetofile(server.MapPath(path)),model+1
		end with
		stream.close : set stream = nothing
		fileSave = fileExists(path)
	end function


	function getHttpPage(url)
		dim objXML
		set objXML=server.createobject("MSXML2.XMLHTTP")
		objXML.open "get",url,false
		objXML.send()
		If objXML.readystate<>4 then
			exit function
		End If
		getHttpPage = bytesToBstr(objXML.responseBody)
		set objXML = nothing
		if err.number<>0 then err.Clear
	end function

	function bytesToBstr(body)
		dim objstream
		set objstream = Server.CreateObject("adodb.stream")
			objstream.Type = 1
			objstream.Mode =3
			objstream.Open
			objstream.write body
			objstream.Position = 0
			objstream.Type = 2
			objstream.Charset = "utf-8"
			BytesToBstr = objstream.ReadText
		objstream.Close
		set objstream = nothing
	end function

	'提交post'
	function post_data(url,data)
		on error resume next
		dim http
		set http = server.createobject("MSXML2.SERVERXMLHTTP.3.0")
			http.open "POST",url,false
			http.setRequestHeader "Content-Type","text/plain"
			http.send(data)
		if http.readystate <> 4 then exit function
		post_data = bytesToBstr(http.responseBody)
		set http=nothing
		if err.number > 0 then
			post_data = "error:" & err.description
		end if
	end function

	function md5(s)
		dim mClass
		set mClass = new md5_class
			s = mClass.md5(s & CONFIG_MD5_STRING)
		set mClass = nothing
		md5 = s
	end function

	'sendMail(服务器地址,发件人地址,发件人名称,邮件服务器用户名,邮件服务器密码,收件人邮箱,邮件标题,邮件内容)
	function sendMail(smtpServer, sFormMail, sFormName, smtpUsername, smtpPassword, targetMail, sTitle, sContent)
	    on error resume next
	    set jmail = server.CreateObject("JMAIL.Message")   '建立发送邮件的对象
	    if err.number <> 0 then sendMail = "服务器不支持jmail" : exit function
	    if smtpServer = "" or smtpUsername = "" or smtpPassword = "" then sendMail = "发送邮件失败(缺少必要信息)！" : exit function
	    jmail.silent = true    '屏蔽例外错误，返回FALSE跟TRUE两值
	    jmail.logging = false   '启用邮件日志
	    jmail.Charset = "GB2312"     '邮件的文字编码为中文
	    jmail.ISOEncodeHeaders = False '防止邮件标题乱码
	    jmail.ContentType = "text/html"    '邮件的格式为HTML格式
	    jmail.AddRecipient targetMail    '邮件收件人的地址
	    jmail.From = sFormMail  '发件人的E-MAIL地址
	    jmail.FromName = sFormName   '发件人姓名
	    jmail.MailServerUserName = smtpUsername    '登录邮件服务器所需的用户名
	    jmail.MailServerPassword = smtpPassword     '登录邮件服务器所需的密码
	    jmail.Subject = sTitle    '邮件的标题
	    jmail.Body = sContent      '邮件的内容
	    jmail.Priority = 1      '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
	    jmail.Send(smtpServer)     '执行邮件发送（通过邮件服务器地址）
	    jmail.Close()   '关闭对象
	    if jmail.errorCode <> 0 Then
	        sendMail = "邮件发送失败("& jmail.errorCode &")!"
	    Else
	        sendMail = "邮件发送至("& targetMail &")成功！"
	    End if
	end function

end class
%>

<%
class md5_class
	private bits_to_a_byte
	private bytes_to_a_word
	private bits_to_a_word
	private m_lonbits(30)
	private m_l2power(30)

	private sub class_initialize
		bits_to_a_byte = 8
		bytes_to_a_word = 4
		bits_to_a_word = 32
	end sub

	private function lshift(lvalue, ishiftbits)
		if ishiftbits = 0 then
			lshift = lvalue
			exit function
		elseif ishiftbits = 31 then
			if lvalue and 1 then
				lshift = &h80000000
			else
				lshift = 0
			end if
			exit function
		elseif ishiftbits < 0 or ishiftbits > 31 then
			err.raise 6
		end if
		if (lvalue and m_l2power(31 - ishiftbits)) then
			lshift = ((lvalue and m_lonbits(31 - (ishiftbits + 1))) * m_l2power(ishiftbits)) or &h80000000
		else
			lshift = ((lvalue and m_lonbits(31 - ishiftbits)) * m_l2power(ishiftbits))
		end if
	end function

	private function rshift(lvalue, ishiftbits)
		if ishiftbits = 0 then
			rshift = lvalue
			exit function
		elseif ishiftbits = 31 then
			if lvalue and &h80000000 then
				rshift = 1
			else
				rshift = 0
			end if
			exit function
		elseif ishiftbits < 0 or ishiftbits > 31 then
			err.raise 6
		end if
		rshift = (lvalue and &h7ffffffe) \ m_l2power(ishiftbits)
		if (lvalue and &h80000000) then
			rshift = (rshift or (&h40000000 \ m_l2power(ishiftbits - 1)))
		end if
	end function

	private function rotateleft(lvalue, ishiftbits)
		rotateleft = lshift(lvalue, ishiftbits) or rshift(lvalue, (32 - ishiftbits))
	end function

	private function addunsigned(lx, ly)
		dim lx4
		dim ly4
		dim lx8
		dim ly8
		dim lresult

		lx8 = lx and &h80000000
		ly8 = ly and &h80000000
		lx4 = lx and &h40000000
		ly4 = ly and &h40000000

		lresult = (lx and &h3fffffff) + (ly and &h3fffffff)

		if lx4 and ly4 then
			lresult = lresult xor &h80000000 xor lx8 xor ly8
		elseif lx4 or ly4 then
			if lresult and &h40000000 then
				lresult = lresult xor &hc0000000 xor lx8 xor ly8
			else
				lresult = lresult xor &h40000000 xor lx8 xor ly8
			end if
		else
			lresult = lresult xor lx8 xor ly8
		end if
		addunsigned = lresult
	end function

	private function md5_f(x, y, z)
		md5_f = (x and y) or ((not x) and z)
	end function

	private function md5_g(x, y, z)
		md5_g = (x and z) or (y and (not z))
	end function

	private function md5_h(x, y, z)
		md5_h = (x xor y xor z)
	end function

	private function md5_i(x, y, z)
		md5_i = (y xor (x or (not z)))
	end function

	private sub md5_ff(a, b, c, d, x, s, ac)
		a = addunsigned(a, addunsigned(addunsigned(md5_f(b, c, d), x), ac))
		a = rotateleft(a, s)
		a = addunsigned(a, b)
	end sub

	private sub md5_gg(a, b, c, d, x, s, ac)
		a = addunsigned(a, addunsigned(addunsigned(md5_g(b, c, d), x), ac))
		a = rotateleft(a, s)
		a = addunsigned(a, b)
	end sub

	private sub md5_hh(a, b, c, d, x, s, ac)
		a = addunsigned(a, addunsigned(addunsigned(md5_h(b, c, d), x), ac))
		a = rotateleft(a, s)
		a = addunsigned(a, b)
	end sub

	private sub md5_ii(a, b, c, d, x, s, ac)
		a = addunsigned(a, addunsigned(addunsigned(md5_i(b, c, d), x), ac))
		a = rotateleft(a, s)
		a = addunsigned(a, b)
	end sub

	private function converttowordarray(smessage)
		dim lmessagelength
		dim lnumberofwords
		dim lwordarray()
		dim lbyteposition
		dim lbytecount
		dim lwordcount
		dim modulus_bits : modulus_bits = 512
		dim congruent_bits : congruent_bits = 448
		lmessagelength = len(smessage)
		lnumberofwords = (((lmessagelength + ((modulus_bits - congruent_bits) \ bits_to_a_byte)) \ (modulus_bits \ bits_to_a_byte)) + 1) * (modulus_bits \ bits_to_a_word)
		redim lwordarray(lnumberofwords - 1)
		lbyteposition = 0
		lbytecount = 0
		do until lbytecount >= lmessagelength
			lwordcount = lbytecount \ bytes_to_a_word
			lbyteposition = (lbytecount mod bytes_to_a_word) * bits_to_a_byte
			lwordarray(lwordcount) = lwordarray(lwordcount) or lshift(asc(mid(smessage, lbytecount + 1, 1)), lbyteposition)
			lbytecount = lbytecount + 1
		loop
		lwordcount = lbytecount \ bytes_to_a_word
		lbyteposition = (lbytecount mod bytes_to_a_word) * bits_to_a_byte
		lwordarray(lwordcount) = lwordarray(lwordcount) or lshift(&h80, lbyteposition)
		lwordarray(lnumberofwords - 2) = lshift(lmessagelength, 3)
		lwordarray(lnumberofwords - 1) = rshift(lmessagelength, 29)
		converttowordarray = lwordarray
	end function

	private function wordtohex(lvalue)
		dim lbyte
		dim lcount
		for lcount = 0 to 3
			lbyte = rshift(lvalue, lcount * bits_to_a_byte) and m_lonbits(bits_to_a_byte - 1)
			wordtohex = wordtohex & right("0" & hex(lbyte), 2)
		next
	end function

	function md5(smessage)
		m_lonbits(0) = clng(1)
		m_lonbits(1) = clng(3)
		m_lonbits(2) = clng(7)
		m_lonbits(3) = clng(15)
		m_lonbits(4) = clng(31)
		m_lonbits(5) = clng(63)
		m_lonbits(6) = clng(127)
		m_lonbits(7) = clng(255)
		m_lonbits(8) = clng(511)
		m_lonbits(9) = clng(1023)
		m_lonbits(10) = clng(2047)
		m_lonbits(11) = clng(4095)
		m_lonbits(12) = clng(8191)
		m_lonbits(13) = clng(16383)
		m_lonbits(14) = clng(32767)
		m_lonbits(15) = clng(65535)
		m_lonbits(16) = clng(131071)
		m_lonbits(17) = clng(262143)
		m_lonbits(18) = clng(524287)
		m_lonbits(19) = clng(1048575)
		m_lonbits(20) = clng(2097151)
		m_lonbits(21) = clng(4194303)
		m_lonbits(22) = clng(8388607)
		m_lonbits(23) = clng(16777215)
		m_lonbits(24) = clng(33554431)
		m_lonbits(25) = clng(67108863)
		m_lonbits(26) = clng(134217727)
		m_lonbits(27) = clng(268435455)
		m_lonbits(28) = clng(536870911)
		m_lonbits(29) = clng(1073741823)
		m_lonbits(30) = clng(2147483647)
		m_l2power(0) = clng(1)
		m_l2power(1) = clng(2)
		m_l2power(2) = clng(4)
		m_l2power(3) = clng(8)
		m_l2power(4) = clng(16)
		m_l2power(5) = clng(32)
		m_l2power(6) = clng(64)
		m_l2power(7) = clng(128)
		m_l2power(8) = clng(256)
		m_l2power(9) = clng(512)
		m_l2power(10) = clng(1024)
		m_l2power(11) = clng(2048)
		m_l2power(12) = clng(4096)
		m_l2power(13) = clng(8192)
		m_l2power(14) = clng(16384)
		m_l2power(15) = clng(32768)
		m_l2power(16) = clng(65536)
		m_l2power(17) = clng(131072)
		m_l2power(18) = clng(262144)
		m_l2power(19) = clng(524288)
		m_l2power(20) = clng(1048576)
		m_l2power(21) = clng(2097152)
		m_l2power(22) = clng(4194304)
		m_l2power(23) = clng(8388608)
		m_l2power(24) = clng(16777216)
		m_l2power(25) = clng(33554432)
		m_l2power(26) = clng(67108864)
		m_l2power(27) = clng(134217728)
		m_l2power(28) = clng(268435456)
		m_l2power(29) = clng(536870912)
		m_l2power(30) = clng(1073741824)
		dim x
		dim k
		dim aa
		dim bb
		dim cc
		dim dd
		dim a
		dim b
		dim c
		dim d
		dim s11 : s11 = 7
		dim s12 : s12 = 12
		dim s13 : s13 = 17
		dim s14 : s14 = 22
		dim s21 : s21 = 5
		dim s22 : s22 = 9
		dim s23 : s23 = 14
		dim s24 : s24 = 20
		dim s31 : s31 = 4
		dim s32 : s32 = 11
		dim s33 : s33 = 16
		dim s34 : s34 = 23
		dim s41 : s41 = 6
		dim s42 : s42 = 10
		dim s43 : s43 = 15
		dim s44 : s44 = 21
		x = converttowordarray(smessage)
		a = &h67452301
		b = &hefcdab89
		c = &h98badcfe
		d = &h10325476
		for k = 0 to ubound(x) step 16
			aa = a
			bb = b
			cc = c
			dd = d
			md5_ff a, b, c, d, x(k + 0), s11, &hd76aa478
			md5_ff d, a, b, c, x(k + 1), s12, &he8c7b756
			md5_ff c, d, a, b, x(k + 2), s13, &h242070db
			md5_ff b, c, d, a, x(k + 3), s14, &hc1bdceee
			md5_ff a, b, c, d, x(k + 4), s11, &hf57c0faf
			md5_ff d, a, b, c, x(k + 5), s12, &h4787c62a
			md5_ff c, d, a, b, x(k + 6), s13, &ha8304613
			md5_ff b, c, d, a, x(k + 7), s14, &hfd469501
			md5_ff a, b, c, d, x(k + 8), s11, &h698098d8
			md5_ff d, a, b, c, x(k + 9), s12, &h8b44f7af
			md5_ff c, d, a, b, x(k + 10), s13, &hffff5bb1
			md5_ff b, c, d, a, x(k + 11), s14, &h895cd7be
			md5_ff a, b, c, d, x(k + 12), s11, &h6b901122
			md5_ff d, a, b, c, x(k + 13), s12, &hfd987193
			md5_ff c, d, a, b, x(k + 14), s13, &ha679438e
			md5_ff b, c, d, a, x(k + 15), s14, &h49b40821
			md5_gg a, b, c, d, x(k + 1), s21, &hf61e2562
			md5_gg d, a, b, c, x(k + 6), s22, &hc040b340
			md5_gg c, d, a, b, x(k + 11), s23, &h265e5a51
			md5_gg b, c, d, a, x(k + 0), s24, &he9b6c7aa
			md5_gg a, b, c, d, x(k + 5), s21, &hd62f105d
			md5_gg d, a, b, c, x(k + 10), s22, &h2441453
			md5_gg c, d, a, b, x(k + 15), s23, &hd8a1e681
			md5_gg b, c, d, a, x(k + 4), s24, &he7d3fbc8
			md5_gg a, b, c, d, x(k + 9), s21, &h21e1cde6
			md5_gg d, a, b, c, x(k + 14), s22, &hc33707d6
			md5_gg c, d, a, b, x(k + 3), s23, &hf4d50d87
			md5_gg b, c, d, a, x(k + 8), s24, &h455a14ed
			md5_gg a, b, c, d, x(k + 13), s21, &ha9e3e905
			md5_gg d, a, b, c, x(k + 2), s22, &hfcefa3f8
			md5_gg c, d, a, b, x(k + 7), s23, &h676f02d9
			md5_gg b, c, d, a, x(k + 12), s24, &h8d2a4c8a
			md5_hh a, b, c, d, x(k + 5), s31, &hfffa3942
			md5_hh d, a, b, c, x(k + 8), s32, &h8771f681
			md5_hh c, d, a, b, x(k + 11), s33, &h6d9d6122
			md5_hh b, c, d, a, x(k + 14), s34, &hfde5380c
			md5_hh a, b, c, d, x(k + 1), s31, &ha4beea44
			md5_hh d, a, b, c, x(k + 4), s32, &h4bdecfa9
			md5_hh c, d, a, b, x(k + 7), s33, &hf6bb4b60
			md5_hh b, c, d, a, x(k + 10), s34, &hbebfbc70
			md5_hh a, b, c, d, x(k + 13), s31, &h289b7ec6
			md5_hh d, a, b, c, x(k + 0), s32, &heaa127fa
			md5_hh c, d, a, b, x(k + 3), s33, &hd4ef3085
			md5_hh b, c, d, a, x(k + 6), s34, &h4881d05
			md5_hh a, b, c, d, x(k + 9), s31, &hd9d4d039
			md5_hh d, a, b, c, x(k + 12), s32, &he6db99e5
			md5_hh c, d, a, b, x(k + 15), s33, &h1fa27cf8
			md5_hh b, c, d, a, x(k + 2), s34, &hc4ac5665
			md5_ii a, b, c, d, x(k + 0), s41, &hf4292244
			md5_ii d, a, b, c, x(k + 7), s42, &h432aff97
			md5_ii c, d, a, b, x(k + 14), s43, &hab9423a7
			md5_ii b, c, d, a, x(k + 5), s44, &hfc93a039
			md5_ii a, b, c, d, x(k + 12), s41, &h655b59c3
			md5_ii d, a, b, c, x(k + 3), s42, &h8f0ccc92
			md5_ii c, d, a, b, x(k + 10), s43, &hffeff47d
			md5_ii b, c, d, a, x(k + 1), s44, &h85845dd1
			md5_ii a, b, c, d, x(k + 8), s41, &h6fa87e4f
			md5_ii d, a, b, c, x(k + 15), s42, &hfe2ce6e0
			md5_ii c, d, a, b, x(k + 6), s43, &ha3014314
			md5_ii b, c, d, a, x(k + 13), s44, &h4e0811a1
			md5_ii a, b, c, d, x(k + 4), s41, &hf7537e82
			md5_ii d, a, b, c, x(k + 11), s42, &hbd3af235
			md5_ii c, d, a, b, x(k + 2), s43, &h2ad7d2bb
			md5_ii b, c, d, a, x(k + 9), s44, &heb86d391
			a = addunsigned(a, aa)
			b = addunsigned(b, bb)
			c = addunsigned(c, cc)
			d = addunsigned(d, dd)
		next
		'md5=lcase(wordtohex(b) & wordtohex(c))
		md5 = ucase(wordtohex(a) & wordtohex(b) & wordtohex(c) & wordtohex(d))
	end function
end class
%>
