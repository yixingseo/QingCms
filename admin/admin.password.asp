<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/admin.class.asp" -->
<!--#include file="cms.meta.html"-->
<body class="abody">

<%dim admin : set admin = new t_admin%>
<h3>修改密码</h3>
<%if not admin.readerBySession() then %>
<div class="alert alert-danger" role="alert">您尚未登录，或者登录已经超时！<a href="login.asp" target="_parent">重新登录</a></div>
<%end if%>
<form action="admin.action.asp?action=password" method="post" name="adminform" id="adminform">
<table class="table table-hover table-bordered table-admin">
  <tr>
    <th><label>管理员名称</label></th>
    <td style="width:550px;"><input name="title" type="text" id="title" value="<%=admin.title%>" "" class="form-control require" readonly /></td>
    <td><div class="tip">管理员的名称/昵称</div></td>
  </tr>
 
  <tr>
    <th><label>登陆用户名</label></th>
    <td><input name="username" type="text" id="username" value="<%=admin.username%>" class="form-control require" readonly /></td>
    <td><div class="tip">用于登陆的用户名</div></td>
  </tr> 

  <tr>
    <th><label>旧密码</label></th>
    <td><input name="old_password" type="text" id="old_password" value="" class="form-control require" /></td>
    <td><div class="tip">原来的密码</div></td>
  </tr>
 
  <tr>
    <th><label>新密码</label></th>
    <td><input name="new_password" type="text" id="new_password" value="" class="form-control require" /></td>
    <td><div class="tip">新密码</div></td>
  </tr>

  <tr>
    <th><label>确认密码</label></th>
    <td><input name="confirm_password" type="password" id="confirm_password" value="" class="form-control require" /></td>
    <td><div class="tip">再次输入新密码</div></td>
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
<%set admin = nothing%>
</body>
</html>
<!--#include file="cms.footer.asp" -->
