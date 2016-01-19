<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html"-->
<%
dim strSql
strSql = "select * from [t_feedback] order by id desc"
%>

<h3>留言评论</h3>
<div class="tools">
<div class="form-inline">
  <div class="form-group"><button type="button" class="btn btn-danger" id="del_all"><i class="fa fa-trash"></i> 批量删除</button></div>
</div>
</div>
<table class="table table-bordered table-hover">
  <tr>
    <th width="50" class="text-center"><input name="select_all" id="select_all" type="checkbox" value=""></th>
    <th style="width:70px;">编号</th>
    <th>标题</th>
    <th>姓名</th>
    <th>手机</th>
    <th>添加时间 / 回复时间</th>
    <th>操作</th>
  </tr>
<%
Dim Rs : Set Rs = db.listpage(strSql,20)
i = 0
while not rs.eof and i < rs.Pagesize
%>
  <tr id="<%=rs("id")%>">
    <td align="center"><input name="ids" type="checkbox" class="ck" value="<%=rs("id")%>" /></td>
    <td class="listview_id"><%=rs("id")%></td>
    <td><%=e.encode(Rs("title"))%></span></td>
    <td><%=e.encode(rs("person"))%></td>
    <td><%=e.encode(rs("mobile"))%></td>
    <td class="text-success"><i class="fa fa-clock-o"></i> <%=rs("insert_date")%> <i class="fa fa-reply"></i> <%=e.iif(isnull(rs("reply_date")),"未回复",rs("reply_date"))%></td>
    <td class="listview_action">
    <a href="feedback.formview.asp?id=<%=rs("id")%>&redirect=<%=e.getUrl%>"><i class="fa fa-reply"></i> 浏览</a>&nbsp;&nbsp;
    <a href="feedback.action.asp?action=delete&id=<%=rs("id")%>&redirect=<%=e.getUrl%>" onClick="return confirm('确定要删除吗')"><i class="fa fa-trash"></i> 删除</a>
    </td>
  </tr>
<%
i = i + 1
Rs.Movenext
Wend
%>
</table>
<nav class="pagelist"><%=Db.GetPager("")%></nav>
<%
Rs.Close
Set Rs = Nothing
%>

<script>
  $(function(){

    //批量删除
    $('#del_all').click(function(){
        //alert('删除选中');
        var ids = "";
        $("input[name='ids']").each(function(){
          if($(this).prop("checked")){
            ids += $(this).val() + ",";
          }
        })
        if(ids == "")
          return;
        //alert(ids);
        $.get("feedback.action.asp", { action:'delete',id: ids} ,function(){
          document.location.href='<%=e.getUrl%>'
        });
      })
    //全选
    $('#select_all').click(function(){
      $('.ck').prop('checked',this.checked);
      })
  })
</script>
</body>
</html>
<!--#include file="cms.footer.asp" -->
