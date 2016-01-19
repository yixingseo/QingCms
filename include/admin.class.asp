<%
class t_admin
	
	public id,title,username,password,email,phone,qq,logintimes
	public table
	sub class_initialize()
		id = 0
		logintimes = 0
		table = "[t_admin]"
	end sub

	'读取'
	function reader(n)
		reader = false
		if n = 0 then exit function
		set rs = db.query("select top 1 * from [t_admin] where id = " & n)
		if rs.eof then exit function
			id = rs("id")
			title = rs("title")
			username = rs("username")
			password = rs("password")
			email = rs("email")
			phone = rs("phone")
			qq = rs("qq")
			logintimes = rs("logintimes")							
		rs.close
		set rs = nothing
		reader = true
	end function

	'读取(用户名)'
	function readerBySession()
		readerBySession = false
		username = session("admin")	
		if username = "" then exit function		
		set rs = db.query("select top 1 * from "& table &" where [username] = '"& username &"'")
		if rs.eof then exit function
			id = rs("id")
			title = rs("title")
			username = rs("username")
			password = rs("password")
			email = rs("email")
			phone = rs("phone")
			qq = rs("qq")
			logintimes = rs("logintimes")							
		rs.close
		set rs = nothing
		readerBySession = true
	end function

	'登录'
	sub login()
		'if e.get("yzm") <> session("yzm") then e.alert("验证码错误") : exit sub
		username = e.post("username")
		password = e.post("password")
		username = e.safe(username)
		password = e.safe(password)
		if username = "" or password = "" then e.alert("请输入用户名和密码") : exit sub
		set rs = e.objRs
			rs.open "select top 1 * from "& table &" where username = '"& username &"' and password = '"& e.md5(password) &"'",db.conn,1,3
			if rs.eof then 				
				call log.add("登录失败",username)
				e.alert("用户名或密码错误")
				exit sub
			else
				rs("logintimes") = rs("logintimes") + 1
				rs.update()
				call log.add("登录成功",username)
				session("admin_is_login") = "true"
				session("admin") = username				
				response.redirect("index.asp")
			end if					
			rs.close
			set rs = nothing		
	end sub
	
	'保存'
	sub saveData()
		if not site.isSuper then e.alert("权限不足")
		id = e.postint("id")
		title = e.post("title")
		if title = "" then e.alert("请输入管理员昵称") : exit sub		

		username = e.post("username")
		if id = 0 then
			if username = "" then e.alert("请输入用户名") : exit sub	
			if not e.test("username",username) then e.alert("用户名只能包含字母和数字，且以字母开头不能纯数字")
			if not checkUsername then e.alert("该用户名已经存在") : exit sub
		end if

		password = e.post("password")
		if password = "" then e.alert("请输入登陆密码") : exit sub
		if id = 0 then
			confirm_password = e.post("confirm_password")		
			if confirm_password = "" then e.alert("请再输入一次登陆密码") : exit sub
			if confirm_password <> password then e.alert("两次输入的密码一不一致") : exit sub				
		end if

		qq = e.post("qq")
		if qq <> "" and not e.test("qq",qq) then e.alert("QQ号码格式不正确")
		phone = e.post("phone")
		'if not e.test("phone",phone) and not e.test("mobile",phone) then e.alert("电话号码不正确")		
		email = e.post("email")
		if email <> "" and not e.test("email",email) then e.alert("邮箱格式不正确")		
		set rs = e.objRs
			rs.open "select top 1 * from "& table &" where id = " & id,db.conn,1,3			
			if rs.eof then
				rs.addnew()
				rs("username") = username							
				rs("logintimes") = logintimes
			end if
			rs("password") = e.md5(password)
			rs("title") = title			
			rs("qq") = qq
			rs("phone") = e.post("phone")
			rs("email") = e.post("email")			
			log.add "修改管理员" , username & "成功" 
		rs.update()
		rs.close
		set rs = nothing		
	end sub

	sub changePassword()
		dim old_password,new_password,confirm_password
		old_password = e.post("old_password")
		if old_password = "" then e.alert("请输入原密码") : exit sub				
		if db.query("select count(*) from "& table &" where [username] = '"& session("admin") &"' and [password] = '"& e.md5(old_password) &"'")(0) < 1 then e.alert("原密码错误") : exit sub
		new_password = e.post("new_password")
		if new_password = "" then e.alert("请输入新密码") : exit sub		
		confirm_password = e.post("confirm_password")
		if confirm_password = "" then e.alert("请再输入一次新密码") : exit sub		
		if new_password <> confirm_password then e.alert("两次输入的密码不一致") : exit sub		
		if db.exec("update [t_admin] set [password] = '"& e.md5(new_password) &"' where [username] = '"& session("admin") &"'") > 0 then
			session.abandon()
			log.add "修改密码" , session("admin") 
			e.echo "<script>"
			e.echo "alert('修改成功，请重新登录');"
			e.echo "parent.window.location.href='login.asp'"
			e.echo "</script>"
			response.end			
		else
			e.alert("操作失败")
		end if
	end sub

	function checkUsername()
		checkUsername = false
		if db.query("select count(*) from "& table &" where [username] = '"& username &"'")(0) < 1 then checkUsername = true
	end function

	sub del()
		id = e.getint("id")
		if id = 0 then exit sub
		if not reader(id) then exit sub
		dim arrAdmin : arrAdmin = split(CONFIG_ADMIN_SUPER,"|")
		dim is_del : is_del = false
		for i = 0 to ubound(arrAdmin)
			'e.echoline arrAdmin(i) & "-" & session("admin")
			if arrAdmin(i) = session("admin") then 				
				if db.exec("delete from "& table &" where id = " & id) > 0 then 
					is_del = true
					log.add "删除用户" , id & "/成功"
					exit sub
				end if	
			end if
		next
		if not is_del then			
			log.add "删除用户：" , id & "/失败"
			e.alert("删除失败，可能是权限不足!")
		end if
	end sub

end class
%>