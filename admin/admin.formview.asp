<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/admin.class.asp" -->
<!--#include file="cms.meta.html"-->

<%
dim admin : set admin = new t_admin
if e.getint("id") > 0 then
  if not admin.reader(e.getint("id")) then e.alert("未找到改用户")  
end if  
%>
<h3>管理员管理</h3>
<%if not site.isSuper then%>
<div class="alert alert-danger" role="alert">需要超级管理员权限，您的权限不够。你可以 <a href="admin.password.asp">修改密码</a></div>
<%else%>
<form action="admin.action.asp?action=save" method="post" name="userform" id="userform">
<table class="table table-hover table-bordered table-admin">
  <tr>
    <th><label>管理员名称</label></th>
    <td style="width:550px;"><input name="title" type="text" id="title" value="<%=admin.title%>" class="form-control require" /></td>
    <td><div class="tip">*管理员昵称</div></td>
  </tr>
 
  <tr>
    <th><label>登陆用户名</label></th>
    <td><input name="username" type="text" id="username" value="<%=admin.username%>" class="form-control require" <%if e.getint("id") > 0 then%>disabled<%end if%> /></td>
    <td><div class="tip">*登录用户名，只能包含字母和数字，且以字母开头不能纯数字</div></td>
  </tr> 

  <tr>
    <th><label>登陆密码</label></th>
    <td><input name="password" type="password" id="password" value="" class="form-control require" /></td>
    <td><div class="tip">*登陆密码</div></td>
  </tr>

  <%If admin.ID = 0 Then%>
  <tr>
    <th><label>确认密码</label></th>
    <td><input name="confirm_password" type="password" id="confirm_password" value="" class="form-control require" /></td>
    <td><div class="tip">*再次输入登陆密码</div></td>
  </tr>
  <%End If%>

  <tr>
    <th><label>手机号码</label></th>
    <td><input name="phone" type="text" id="phone" value="<%=admin.phone%>" class="form-control" /></td>
    <td><div class="tip">手机号码</div></td>
  </tr>

  <tr>
    <th><label>QQ号码</label></th>
    <td><input name="qq" type="text" id="qq" value="<%=admin.qq%>" class="form-control" /></td>
    <td><div class="tip">QQ号码</div></td>
  </tr>    

  <tr>
    <th><label>邮箱</label></th>
    <td><input name="email" type="text" id="email" value="<%=admin.email%>" class="form-control" /></td>
    <td><div class="tip">管理员邮箱</div></td>
  </tr>    
  <tr>
    <th>&nbsp;</th>
    <td>
    <input name="id" type="hidden" id="id" value="<%=admin.id%>"> 
    <button type="submit" class="btn btn-primary"> <i class="fa fa-check"></i> 保存设置</button>
    </td>
    <td></td>
  </tr>
</table>
</form>
<%end if%>



</body>
</html>
<%set admin = nothing%>
<!--#include file="cms.footer.asp" -->