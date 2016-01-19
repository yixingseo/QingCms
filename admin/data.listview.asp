<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html" -->
<h3>数据库备份</h3>
<!-- 操作 -->
<%
dim objFso,o_file,s_path,s_datafile,arr_path
  s_datafile = site.dic("root") & CONFIG_DATA
set objFso = e.objFso
  set o_file = objFso.getfile(server.mappath(s_datafile))  
%>
<div class="tools">        
  <table>
    <tr>
      <td>
        <a class="btn btn-primary" href="action.asp?action=bak" id="bak"> 备份当前数据库 </a>
        <a class="btn btn-success" href="action.asp?action=zip" id="zip"> 压缩当前数据库 <%=formatNumber(o_file.Size / 1024 / 1024,5,true,,0)%>MB</a>
      </td>
    </tr>
  </table>
</div>
<%
set o_file = nothing
s_path = site.dic("root") & CONFIG_DATA_BAK 
dim obj_dir : set obj_dir = objFso.getfolder(server.mappath(s_path))
dim obj_files : set obj_files = obj_dir.files
%>
<table class="table table-bordered table-hover">
  <tr>
    <th>名称</th>
    <th>大小</th>
    <th>创建时间</th>    
    <th>操作</th>
  </tr>
<%
'On Error Resume Next
dim f
for each f in obj_files
%>  
  <tr>
    <td><span class="text-primary"><%=f.name%></span></td>
    <td><%=f.size / 1024%> KB (<%=formatNumber(f.Size / 1024 / 1024,5,true,,0)%>MB) </td>
    <td><%=f.dateCreated%></td>    
    <td class="listview_action">
    <a href="action.asp?action=down&file=<%=f.name%>"><i class="fa fa-download"></i> 下载</a>&nbsp;&nbsp;
    <a href="action.asp?action=delbak&file=<%=f.name%>&burl=<%=e.url%>" onclick="return confirm('确定要删除吗')"><i class="fa fa-trash"></i> 删除</a>
    </td>
  </tr>
<%
next
set f = nothing
set objFso = nothing
%>  
</table>
</body>
</html>
<!--#include file="cms.footer.asp" -->