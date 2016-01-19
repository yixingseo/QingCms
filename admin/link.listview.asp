<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html"-->
<%
dim strSql
strSql = "select * from [t_link] order by id desc"
%>

<h3>友情链接</h3>
<div class="tools">
<table>
  <tr>
    <td>
      <button type="button" class="btn btn-danger" id="del_all"><i class="fa fa-trash"></i> 批量删除</button>
      <a class="btn btn-primary" href="link.formview.asp"><i class="fa fa-plus-circle"></i> 添加链接</a>
    </td>
  </tr>
</table>
</div>
<table class="table table-bordered table-hover">
  <tr>
    <th width="50" class="text-center"><input name="select_all" id="select_all" type="checkbox" value=""></th>
    <th style="width:70px;">编号</th>
    <th>标题</th>
    <th>图片</th>
    <th>网址</th>
    <th>打开方式</th>
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
    <td><%=Rs("title")%></span></td>
    <td><%if rs("pic") <> "" then%><i class="fa fa-photo showpic"></i> <img src="<%=rs("pic")%>" class="list_thumail"><%end if%> &nbsp; <%=rs("pic")%></td>
    <td><span class="text-primary"><a href="<%=rs("url")%>" target="_blank"><%=rs("url")%></a></span></td>
    <td>
    <%
    if rs("target") = "_blank" then
      e.echo "<i class='fa fa-share-square-o'></i> 新窗口"
    else
      e.echo "<i class='fa fa-share-square'></i> 原窗口"
    end if
    %></td>
    <td class="listview_action">
    <a href="link.formview.asp?id=<%=rs("id")%>&redirect=<%=e.getUrl%>"><i class="fa fa-pencil"></i> 修改</a>&nbsp;&nbsp;
    <a href="link.action.asp?action=delete&id=<%=rs("id")%>&redirect=<%=e.getUrl%>" onClick="return confirm('确定要删除吗')"><i class="fa fa-trash"></i> 删除</a>
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

    $('.showpic').mouseover(function(){
    $(this).next('img').toggle();
    })
    $('.showpic').mouseout(function(){
      $(this).next('img').toggle();
    })

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
        $.get("action.asp", { action:'link_del',id: ids} ,function(){
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
