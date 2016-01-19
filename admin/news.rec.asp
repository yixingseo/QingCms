<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/sort.class.asp" -->
<!--#include file="cms.meta.html"-->
<%

dim strSql
strSql = "select news.id as id,news.title as title,news.insert_date as insert_date,news.hits,news.weight,news.insert_date,sort.id,sort.title from [t_news] as news left join [t_sort] as sort on news.pid = sort.id"
	strSql = strSql & " where show < 0"
	strSql = strSql & " order by news.weight desc,news.id desc"
%>

<h3>回收站</h3>
<div class="tools">
	<div class="form-inline">
		<div class="form-group">
			<button type="button" class="btn btn-primary" id="rec_all"><i class="fa fa-reply"></i> 批量恢复</button>
			<button type="button" class="btn btn-danger" id="del_all"><i class="fa fa-trash"></i> 批量删除</button>
		</div>
	</div>
</div>
<table class="table table-bordered table-hover">
  <tr>
  	<th width="50" class="text-center"><input name="select_all" id="select_all" type="checkbox" value=""></th>
    <th style="width:70px;">编号</th>
    <th>标题</th>
    <th class="text-center">分类</th>
    <th>添加时间</th>
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
    <td><span class="ajax_title" newsid="<%=rs("id")%>"><%=server.htmlencode(Rs("title"))%></span></td>
    <td align="center"><%=E.Iif(Rs("sort.title") <> "",Rs("sort.title"),"根目录")%></td>
    <td class="text-success"><i class="fa fa-clock-o"></i> <%=Rs("insert_date")%></td>
    <td class="listview_action">
    <a href="news.action.asp?action=display&id=<%=rs("id")%>&redirect=<%=e.getUrl%>"><i class="fa fa-reply"></i> 恢复</a>&nbsp;&nbsp;
    <a href="news.action.asp?action=del&id=<%=rs("id")%>&redirect=<%=e.getUrl%>" onClick="return confirm('确定要删除吗')"><i class="fa fa-trash"></i> 彻底删除</a>
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
    //批量恢复
    $('#rec_all').click(function(){
        var ids = "";
        $("input[name='ids']").each(function(){
          if($(this).prop("checked")){
            ids += $(this).val() + ",";
          }
        })
        if(ids == "")
          return;
        $.get("news.action.asp", { action:'display',id: ids} ,function(data){
          //alert(data);
          document.location.href='<%=e.getUrl%>'
        });
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
        $.get("news.action.asp", { action:'delete',id: ids} ,function(){
          document.location.href='<%=e.getUrl%>'
        });
      })
    //全选
    $('#select_all').click( function(){$('.ck').prop('checked',this.checked);} )
  })
</script>
</body>
</html>
<!--#include file="cms.footer.asp" -->
