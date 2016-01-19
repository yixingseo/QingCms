<%
dim site : set site = new ClassSite
class ClassSite

	public dic
	dim rs

    sub class_initialize()
    	set dic = e.objDic
		call reader()
	end sub

	sub reader()
		set rs = db.query("select config_name,config_value from t_config")
		while not rs.eof
			if isnull(rs("config_value").value) then
				dic.add rs("config_name").value,""	'处理null'
			else
				dic.add rs("config_name").value,rs("config_value").value
			end if
		rs.movenext
		wend
		rs.close
		set rs = nothing
		dic.add "templates",CONFIG_TEMPLATES
	end sub

	property get newsAttArray
		if dic("news_att") = "" then
			newsAttArray = split(CONFIG_ATTR,",")
		else
			newsAttArray = split(dic("news_att"),",")
		end if
	end property

	function getMeta(sKeywords,sDescription)
		dim strTemp : strTemp = ""
		if sKeywords <> "" then
			strTemp = strTemp & "<meta name=""keywords"" content="""& sKeywords &""" />"
		end if
		if sDescription <> "" Then
			strTemp = strTemp & "<meta name=""description"" content="""& sDescription &""" />"
		end if
		getMeta = strTemp & Dic("meta")
	end function

	'模板下拉列表
	function TemplateFileList
		dim strTemp : strTemp = ""
		dim strPath : strPath = dic("root") & "content/templates/"
		if not e.fileExists(strPath) then exit function
		dim oFso : set oFso = e.objFso
		dim oDir : set oDir = oFso.getFolder(server.mapPath(strPath))
		dim oFiles : set oFiles = oDir.files
		dim f
		for each f in oFiles
			if right(f.name,5) = ".html" then
				strTemp = strTemp & "<option value='"& f.name &"'>"& f.name &"</option>"
			end if
		next
		TemplateFileList = strTemp
	end function

	property get templateFiles
		templateFiles = TemplateFileList
	end property

	'检测是否超级管理员'
	function isSuper()
		isSuper = false
		if session("admin") = "" then exit function
		dim arrUsers : arrUsers = split(CONFIG_ADMIN_SUPER,",")
		for i = 0 to ubound(arrUsers)
			if session("admin") = arrUsers(i) then isSuper = true : exit function
		next
	end function

	'监测注入信息'
	sub checkString(sArray)
		dim oReg : set oReg = e.objReg
		oReg.pattern = CONFIG_SAFE_STRING
		dim sItem,sValue
		for each sItem in sArray
			sValue = sArray(sItem)
			if oReg.test(sValue) then
				if CONFIG_SAFE_LOG then call log.add("危险注入",sValue)
				e.die "非法注入！你的行为已经被记入日志！"
			end if
		next
		set oReg = nothing
	end sub

	sub safe()
		if request.querystring <> "" then call checkString(request.querystring)
		if request.form <> "" then call checkString(request.form)
	end sub

	'发送邮件'
	'Jmail发送邮件'
	function sendMail(sName,sMail,sTitle,sBody)
		s_SMTPServer = JMAIL_SMTP
		s_FromMail = JMAIL_USER
		s_FromName = sName
		s_MailServerUserName = JMAIL_USER
		s_MailServerPassword = JAMIL_PASS
		s_ToEmail = sMail
		s_Subject = sTitle
		s_Body = sBody
	    sendMail = e.SendMail_JMail(s_SMTPServer, s_FromMail, s_FromName, s_MailServerUserName, s_MailServerPassword, s_ToEmail, s_Subject, s_Body)
	end function

end class
%>
