<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/feedback.class.asp" -->
<!--#include file="cms.meta.html"-->
<script src="../include/ueditor/ueditor.config.js"></script>
<script src="../include/ueditor/ueditor.all.min.js"></script>
<%
dim id : id = e.getInt("id")
if id > 0 then feedback.read(id)
%>
<h3>留言评论</h3>
<form action="feedback.action.asp?action=rep&redirect=<%=e.get("redirect")%>" method="post" name="form" id="form">
<table class="table table-hover table-bordered table-admin">

  <tr>
    <th>姓名：</th>
    <td><%=feedback.person%></td>
  </tr>
  <tr>
    <th>标题：</th>
    <td><%=feedback.title%></td>
  </tr>
  <tr>
    <th>内容：</th>
    <td><%=feedback.content%></td>
  </tr>
  <tr>
    <th>地址：</th>
    <td><%=feedback.address%></td>
  </tr>
  <tr>
    <th>电话：</th>
    <td><%=feedback.phone%></td>
  </tr>
  <tr>
    <th>手机：</th>
    <td><%=feedback.mobile%></td>
  </tr>
  <tr>
    <th>公司名称：</th>
    <td><%=feedback.company%></td>
  </tr>
  <tr>
    <th>邮箱：</th>
    <td><%=feedback.email%></td>
  </tr>
  <tr>
    <th>添加日期：</th>
    <td><%=feedback.insert_date%></td>
  </tr>
  <tr>
    <th>添加IP：</th>
    <td><%=feedback.insert_ip%></td>
  </tr>
  <tr>
    <th><label>你的回复</label></th>
    <td colspan="2">
      <script id="reply" name="reply" type="text/plain" style="width:760px;height:300px;"><%=feedback.reply%></script>
    </td>
  </tr>

  <tr>
    <th>&nbsp;</th>
    <td><input name="id" type="hidden" id="id" value="<%=feedback.id%>"> <input type="submit" name="button" id="save" value="保 存" class="btn btn-primary">&nbsp;&nbsp;
  <%if e.get("redirect") <> "" then%><a href="<%=e.get("redirect")%>" onClick="return confirm('确定要放弃修改吗？')">返回</a><%end if%>
    </td>
  </tr>
</table>
</form>
<script>
//编辑器
var ue = UE.getEditor('reply');
</script>
<%set feedback = nothing%>
</body>
</html>
<!--#include file="cms.footer.asp" -->
