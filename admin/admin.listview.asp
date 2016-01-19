<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html" -->
<body class="abody">
<h3>管理员管理</h3>
<div class="tools">
<div class="form-inline">
  <div class="form-group">
    <a class="btn btn-primary" href="admin.formview.asp"><i class="fa fa-plus-circle"></i> 添加管理员</a>
  </div>
</div>
</div>
<table class="table table-bordered table-hover">
<tr>
    <th style="width:70px;">编号</th>
    <th>昵称</th>
    <th>登录名</th>
    <td>电话</td>
    <td>QQ</td>
    <td>邮箱</td>
    <th>登陆次数</th>
    <th>操作</th>
</tr>
<%
dim rs
set rs = db.query("select * from [t_admin] order by id desc")
while not rs.eof
%>
<tr id="<%=rs("id")%>">
    <td class="listview_id"><%=rs("id")%></td>
    <td><span class="text-primary"><%=rs("title")%></span></td>
    <td><span class="text-success"><%=rs("username")%></span></td>
    <td><%=e.encode(rs("phone"))%></td>
    <td><%=rs("qq")%></td>
    <td><%=rs("email")%></td>
    <td><%=rs("logintimes")%>次</td>
    <td class="listview_action">
    <a href="admin.formview.asp?id=<%=rs("id")%>&burl=<%=e.getUrl%>"><i class="fa fa-pencil"></i> 修改</a>&nbsp;&nbsp;
    <a href="admin.action.asp?action=del_data&id=<%=rs("ID")%>&burl=<%=e.getUrl%>" onclick="return confirm('确定要删除吗')"><i class="fa fa-trash"></i> 删除</a>
    </td>
</tr>
<%
Rs.Movenext
Wend
%>
</table>
<%
Rs.Close
Set Rs = Nothing
%>
</body>
</html>

<!--#include file="cms.footer.asp" -->
