<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/tag.class.asp" -->
<!--#include file="cms.meta.html"-->
<script src="../include/ueditor/ueditor.config.js"></script>
<script src="../include/ueditor/ueditor.all.min.js"></script>
<%
dim id : id = e.getInt("id")
if id > 0 then tag.read(id)
%>
<h3>区块管理</h3>
<form action="tag.action.asp?action=<%=e.iif(tag.id > 0,"update","create")%>" method="post" name="tag_form" id="tag_form">
<table class="table table-hover table-bordered table-admin">
  <tr>
    <th><label>标签名称</label></th>
    <td style="width:550px;"><input name="title" type="text" id="title" value="<%=tag.title%>" class="form-control" /></td>
    <td><div class="tip">标签名称，仅支持字母和数字，且不能纯数字</div></td>
  </tr>

  <tr>
    <th><label>标签介绍</label></th>
    <td>
    <textarea name="description" class="form-control" id="description"><%=tag.description%></textarea>
    </td>
    <td><div class="tip">标签介绍</div></td>
  </tr>

  <tr>
    <th><label>标签内容</label></th>
    <td colspan="2">
      <script id="content" name="content" type="text/plain" style="width:760px;height:300px;"><%=tag.content%></script>
    </td>
  </tr>

  <tr>
    <th>&nbsp;</th>
    <td>
    <input name="id" type="hidden" id="id" value="<%=tag.id%>"> 
    <button type="submit" class="btn btn-primary"> <i class="fa fa-check"></i> 保存设置</button>
    </td>
    <td></td>
  </tr>
</table>
</form>
<script>
//编辑器
var ue = UE.getEditor('content');
$(function(){
    $('#tag_form').submit(function(){
      if($('#title').val() == ''){
          $('#title').parent().parent().addClass('has-error');
          return false;
        }
    })
})
</script>
<%set tag = nothing%>
</body>
</html>
<!--#include file="cms.footer.asp" -->
