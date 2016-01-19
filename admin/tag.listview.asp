<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html" -->
<h3>区块管理</h3>
<div class="tools">
<div class="form-inline">
  <div class="form-group"><a class="btn btn-primary" href="tag.formview.asp"><i class="fa fa-plus-circle"></i> 添加标签</a></div>
</div>
</div>
<table class="table table-bordered table-hover">
  <tr>
    <th style="width:70px;">编号</th>
    <th>描述</th>
    <th>标签名称</th>
    <th>调用名称</th>
    <th>操作</th>
  </tr>
<%
Dim Rs : Set Rs = db.listpage("Select ID,Title,Description From [t_tag] Order By ID Desc",20)
i = 0
While Not Rs.Eof And i < Rs.Pagesize
%>
  <tr id="<%=Rs("ID")%>">
    <td class="listview_id"><%=rs("id")%></td>
    <td><%=e.encode(rs("Description"))%></td>
    <td><span class="text-primary"><%=Rs("Title")%></span></td>
    <td><span class="text-success">{#tag.<%=Rs("Title")%>}</span></td>
    <td class="listview_action">
    <a href="tag.formview.asp?id=<%=rs("id")%>"><i class="fa fa-pencil"></i> 修改</a>&nbsp;&nbsp;
    <a href="tag.action.asp?action=delete&id=<%=rs("ID")%>" onclick="return confirm('确定要删除吗')"><i class="fa fa-trash"></i> 删除</a></td>
  </tr>
<%
i = i + 1
Rs.Movenext
Wend
%>
</table>

<nav class="pagelist"><%=db.getpager("")%></nav>

</body>
</html>
<!--#include file="cms.footer.asp" -->
