<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html"-->

<!-- 系统设置 -->
<h3>邮箱设置</h3>
<form action="config.action.asp?action=save&redirect=<%=e.getUrl%>" method="post">
<table width="100%" class="table table-bordered table-hover table-admin">
	<!-- smtp服务器地址 -->
	<tr>
    	<th><label>smtp服务器地址</label></th>
          <td style="width:650px;">
              <input name="smtp_server" type="text" class="form-control" id="smtp_server" value="<%=site.dic("smtp_server")%>" maxlength="255" />
          </td>
          <td>
            <div class="tip">smtp服务器 <abbr title="例如smtp.163.com smtp.qq.com">地址 </abbr></div>
          </td>
    </tr>

	<!-- smtp服务器地址 -->
	<tr>
    	<th><label>smtp服务器用户名</label></th>
          <td style="width:650px;">
              <input name="smtp_username" type="text" class="form-control" id="smtp_username" value="<%=site.dic("smtp_username")%>" maxlength="255" />
          </td>
          <td>
            <div class="tip">smtp服务器登陆 <abbr title="一般为邮箱地址 例如web@company.com">用户名</abbr></div>
          </td>
    </tr>

    <!-- smtp服务器地址 -->
	<tr>
    	<th><label>smtp服务器密码</label></th>
          <td style="width:650px;">
              <input name="smtp_password" type="text" class="form-control" id="smtp_password" value="<%=site.dic("smtp_password")%>" maxlength="255" />
          </td>
          <td>
            <div class="tip">smtp服务器 <abbr title="一般为邮箱登陆密码">登陆密码</abbr></div>
          </td>
    </tr>

	<!-- 订购留言评论邮件 -->
	<tr>
    	<th><label>订购/留言/评论</label></th>
          <td style="width:650px;">
			<label><input type="radio" name="mail_to_admin" id="mail_to_admin1" value="true" <%if site.dic("mail_to_admin") = "true" then%>checked<%end if%>> 是</label> &nbsp;&nbsp;
			<label><input type="radio" name="mail_to_admin" id="mail_to_admin2" value="false" <%if site.dic("mail_to_admin") = "false" then%>checked<%end if%>> 否</label>
          </td>
          <td>
            <div class="tip">收到 <abbr title="订购/留言/评论等信息">反馈信息</abbr> 是否发送邮件至 <abbr title="网站设置里修改">管理员邮箱</abbr> 和客户邮箱</div>
          </td>
    </tr>

	<!-- 回复邮件 -->
	<tr>
    	<th><label>留言回复</label></th>
          <td style="width:650px;">
			<label><input type="radio" name="mail_to_user" id="mail_to_user1" value="true" <%if site.dic("mail_to_user") = "true" then%>checked<%end if%>> 是</label> &nbsp;&nbsp;
			<label><input type="radio" name="mail_to_user" id="mail_to_user2" value="false" <%if site.dic("mail_to_user") = "false" then%>checked<%end if%>> 否</label>
          </td>
          <td>
            <div class="tip">留言回复信息是否发送邮件至 <abbr title="如果客户在留言时填写了邮箱，我的回复信息会发送至客户邮箱">客户邮箱</abbr></div>
          </td>
    </tr>

	<!-- 邮件模板 -->
	<tr>
		<th><label>管理员邮件模板</label></th>
		<td>
		    <textarea name="mail_admin_template" rows="10" class="form-control code" id="mail_admin_template"><%=site.dic("mail_admin_template")%></textarea>
		</td>
		<td>
		  <div class="tip">发送给管理员的邮件模板</div>
		</td>
	</tr>

	<!-- 底部代码 -->
	<tr>
		<th><label>客户邮件模板</label></th>
		<td>
		    <textarea name="mail_user_template" rows="10" class="form-control code" id="mail_user_template"><%=site.dic("mail_user_template")%></textarea>
		</td>
		<td>
		  <div class="tip">发送给客户的邮件模板</div>
		</td>
	</tr>

	<tr>
		<th>&nbsp;</th>
		<td>&nbsp;
		<button type="submit" class="btn btn-primary"> <i class="fa fa-check"></i> 保存设置</button>		
		<button type="button" class="btn btn-default" onclick="mail_test()"><i class="fa fa-share"></i> 发送测试邮件</button>				  
		</td>
		<td></td>
	</tr>
</table>
</form>
</body>
</html>
<!--#include file="cms.footer.asp" -->
