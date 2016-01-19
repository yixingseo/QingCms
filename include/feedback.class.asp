<%
dim feedback : set feedback = new ClassFeedback
class ClassFeedback

	public id,nid,title,content,insert_date,insert_ip,person,address,phone,mobile,company,email,reply,reply_date
	public table
	dim rs

	sub class_initialize()
		id = 0
		logintimes = 0
		insert_ip = e.getIP()
		insert_date = now()
		reply_date = now()
		table = "[t_feedback]"
	end sub

	function read(intID)
		read = false
		if isNull(intID) or not isNumeric(intID) then exit function
		set rs = db.query("select top 1 * from "& table &" where id = " & intID)
			if rs.eof then exit function
				id = rs("id")
				nid = rs("nid")
				title = rs("title")
				content = rs("content")
				insert_date = rs("insert_date")
				insert_ip = rs("insert_ip")
				person = rs("person")
				address = rs("address")
				phone = rs("phone")
				mobile = rs("mobile")
				company = rs("company")
				email = rs("email")
				reply = rs("reply")
				reply_date = rs("reply_date")
			rs.close
			set rs = nothing
		read = true
	end function

	'新增'
	function insert()
			id = e.postInt("id")
			nid = e.postInt("pid")
			title = e.post("title")
			content = e.post("content")
			person = e.post("person")
			address = e.post("address")
			phone = e.post("phone")
			mobile = e.post("mobile")
			company = e.post("company")
			email = e.post("email")
			if title = "" and person = "" then e.alert("请检查表单，输入完整信息")
			set rs = e.objRs
				rs.open "select top 1 * from " & table,db.conn,1,3
				rs.addnew()
				rs("nid") = nid
				rs("title") = title
				rs("content") = content
				rs("person") = person
				rs("address") = address
				rs("phone") = phone
				rs("mobile") = mobile
				rs("company") = company
				rs("email") = email
				rs("insert_ip") = insert_ip
				rs("insert_date") = insert_date
				rs.update()
				rs.close
			set rs = nothing
			if site.dic("mail_to_admin") = "true" then 
				call sendMailToAdmin()
				call sendMailToUser()
			end if	
	end function

	'回复'
	function rep()
		id = e.postint("id")
		reply = e.post("reply")
		set rs = e.objRs
			rs.open "select top 1 * from "& table &" where id = " & id,db.conn,1,3
			if rs.eof then exit function
				rs("reply") = reply
				rs("reply_date") = reply_date
				rs.update()
				rs.close
		set rs = nothing
		if site.dic("mail_to_user") = "true" then 
			dim sTitle,sContent
			sTitle = "订购/留言/评论的回复信息 " & site.dic("domain")
			sContent = reply & "<br>"
			sContent = sContent & "来自：" & site.dic("title") & " / " & site.dic("domain") & "<br />"
			sContent = sContent & "时间：" & now()
			call sendMail(sTitle,sContent)
		end if
	end function

	'删除'
	function delete()
		id = e.get("id")
		if not isNumeric(id) then exit function
		delete = db.exec("delete from "& table &" where id in("& id &")")
	end function

	'发送邮件给管理员'
	function sendMailToAdmin()		
		if site.dic("smtp_server") = "" or site.dic("smtp_username") = "" or site.dic("smtp_password") = "" then exit function
		sTitle = "来自网站的反馈信息-" & site.dic("domain") & "/" & now()
		'发送邮件给管理员'
		dim strTemp : strTemp = site.dic("mail_admin_template")
		strTemp = e.regReplace(strTemp,"\n","<br>")
		strTemp = replace(strTemp,"{标题}",title)
		strTemp = replace(strTemp,"{公司}",company)
		strTemp = replace(strTemp,"{姓名}",person)
		strTemp = replace(strTemp,"{电话}",phone)
		strTemp = replace(strTemp,"{手机}",mobile)
		strTemp = replace(strTemp,"{邮箱}",email)
		strTemp = replace(strTemp,"{备注}",content)
		strTemp = replace(strTemp,"{网址}",site.dic("domain"))
		strTemp = replace(strTemp,"{时间}",now())
		sContent = strTemp
		sendMailToAdmin = sendMail(sTitle,sContent)
	end function

	'发送邮件给用户'
	function sendMailToUser()
		if site.dic("smtp_server") = "" or site.dic("smtp_username") = "" or site.dic("smtp_password") = "" then exit function
		if email <> "" then 
			strTemp = site.dic("mail_user_template")
			strTemp = e.regReplace(strTemp,"\n","<br>")
			strTemp = replace(strTemp,"{客服电话}",site.dic("phone"))
			strTemp = replace(strTemp,"{客服邮箱}",site.dic("email"))
			strTemp = replace(strTemp,"{公司名称}",site.dic("title"))
			strTemp = replace(strTemp,"{时间}",now())
			targetMail = email
			sTitle = "您的反馈信息我们已经收到-" & site.dic("domain") & "/" & now()
			sContent = strTemp		
			call sendMail(sTitle,sContent)
		end if
	end function

	'邮件测试'
	function testMail()
		title = "这是一封测试邮件"
		company = "客户的公司名称"
		person = "联系人"
		phone = "010-123456789"
		mobile = "013912345678"
		email = "web@company.com"
		content = "这是一封测试邮件，一下是内容正文：测试邮件测试邮件测试邮件测试邮件"
		dim strTemp : strTemp = site.dic("mail_admin_template")
		strTemp = e.regReplace(strTemp,"\n","<br>")
		strTemp = replace(strTemp,"{标题}",title)
		strTemp = replace(strTemp,"{公司}",company)
		strTemp = replace(strTemp,"{姓名}",person)
		strTemp = replace(strTemp,"{电话}",phone)
		strTemp = replace(strTemp,"{手机}",mobile)
		strTemp = replace(strTemp,"{邮箱}",email)
		strTemp = replace(strTemp,"{备注}",content)
		strTemp = replace(strTemp,"{网址}",site.dic("domain"))
		strTemp = replace(strTemp,"{时间}",now())
		sContent = strTemp
		sTitle = "这是一封测试邮件-" & site.dic("domain") & "/" & now()
		testMail = sendMail(sTitle,sContent)	
	end function

	'发送邮件'
	function sendMail(sTitle,sContent)
		smtpServer = site.dic("smtp_server")
		smtpUsername = site.dic("smtp_username")
		smtpPassword = site.dic("smtp_password")
		targetMail = site.dic("email")				
		sFormMail = site.dic("smtp_username")
		sFormName = site.dic("title") & site.dic("domain")
		sendMail = e.sendMail(smtpServer, sFormMail, sFormName, smtpUsername, smtpPassword, targetMail, sTitle, sContent)		
	end function

end class
%>
