<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html" -->
<h3>日志管理</h3>
<!-- 操作 -->
<div class="tools">        
<table>
  <tr>
    <td>
      <button type="button" class="btn btn-danger" id="del_all"><i class="fa fa-trash"></i> 清空日志</button>
    </td>
    <td>
      <div class="input-group">          
          <input type="text" class="form-control" name="ip" id="ip" placeholder="123.123.123.123" value="<%=e.get("ip")%>">
          <div class="input-group-btn">
            <button class="btn btn-default btn-primary" type="button" id="search-btn">搜索IP</button>
          </div>
      </div>
    </td>
  </tr>
</table>
</div>
<!-- 操作栏 -->
<%
dim sql
  sql = "select * from [t_logs]"
if e.get("ip") <> "" then
  sql = sql & " where insert_ip like '%"& e.safe(e.get("ip"))&"%'"
end if
sql = sql & " order by id desc"
%>
<table class="table table-bordered table-hover">
  <tr>        
    <th>IP</th>
    <th>类型</th>    
    <th>用户</th>   
    <th>操作信息</th>
    <th>操作时间</th>
    <th>操作</th>
  </tr>
<%
Dim Rs : Set Rs = db.listpage(Sql,20)
i = 0
While Not Rs.Eof And i < Rs.Pagesize
%>  
  <tr id="<%=Rs("ID")%>">    
    <td><span class="text-success"><%=rs("insert_ip")%></span></td>
    <td><%=rs("title")%></td>
    <td><span class="text-primary"><%=rs("user")%></span></span></td>
    <td>
      <%
        if len(rs("content"))  > 70 then
          e.echo e.cutstring(rs("content"),70)
          e.echo "<a href='#' id='view' logid='"& rs("id") &"' data-toggle='modal' data-target='#myModal'>[查看]</a>"
        else
          e.echo rs("content")
        end if
      %>

    </td>
    <td><i class="fa fa-clock-o"></i> <%=rs("insert_date")%></td>
    <td>
    <a href="log.action.asp?action=del&id=<%=rs("id")%>&burl=<%=e.url%>" onclick="return confirm('确定要删除吗')"><i class="fa fa-trash"></i> 删除</a     
    </td>
  </tr>
<%
i = i + 1
Rs.Movenext
Wend
%>  
</table>
<nav class="pagelist"><%=Db.GetPager("")%></nav>
<script>
$(function(){
  $('#search-btn').click(function(){
    document.location.href='log.listview.asp?ip=' + $('#ip').val();
  })

  $('#del_all').click(function(){
    document.location.href='log.action.asp?action=clear&burl=<%=e.url%>';
  })

  $('#view').click(function(){
      var id = $(this).attr("logid");
      $.get("log.action.asp", { id: id ,action : 'ajax'} ,function(data){
        $('#log_content').html(data);
      });
  })
})  
</script>

<!-- Modal -->
<div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
        <h4 class="modal-title" id="myModalLabel">日志详情</h4>
      </div>
      <div class="modal-body">
        <div id="log_content"></div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">关闭</button>        
      </div>
    </div>
  </div>
</div>
</body>
</html>
<!--#include file="cms.footer.asp" -->

    