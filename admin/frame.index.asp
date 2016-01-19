<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html"-->
<h3>服务器信息</h3>
<table class="table table-hover table-admin table-bordered">
      <tr>
        <th>服务器地址</th>
        <td>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>) 端口:<%=Request.ServerVariables("SERVER_PORT")%></td>
      </tr>
      <tr>
        <th>服务器时间</th>
        <td><%=Now()%></td>
      </tr>
      <tr>
        <th>IIS版本</th>
        <td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
      </tr>
      <tr>
        <th>脚本超时</th>
        <td><%=Server.ScriptTimeout%>分</td>
      </tr>
      <tr>
        <th>服务器路径</th>
        <td><%=Request.ServerVariables("PATH_TRANSLATED")%></td>
      </tr> 
      <tr>
        <th>Application</th>
        <td><%=Application.Contents.Count%></td>
      </tr>
      <tr>
        <th>Session</th>
        <td><%=Session.Contents.Count%></td>
      </tr>  
      <tr>
        <th>Cookies</th>
        <td><%=Request.Cookies.Count%></td>
      </tr>  
      <tr>
        <th>GUID</th>
        <td><%=e.guid()%></td>
      </tr>  
</table>

<h3>组件检测</h3>
<table class="table table-hover table-admin table-bordered">
  <tr>
    <th>MSWC.AdRotator</th>
    <td><%=e.isinstall("MSWC.AdRotator")%></td>
  </tr>
  <tr>
    <th>MSWC.BrowserType</th>
    <td><%=e.isinstall("MSWC.BrowserType")%></td>
  </tr>
  <tr>
    <th>MSWC.NextLink</th>
    <td><%=e.isinstall("MSWC.NextLink")%></td>
  </tr>
  <tr>
    <th>Scripting.FileSystemObject</th>
    <td><%=e.isinstall("Scripting.FileSystemObject")%></td>
  </tr>
  <tr>
    <th>Adodb.Stream</th>
    <td><%=e.isinstall("Adodb.Stream")%></td>
  </tr> 
  <tr>
    <th>JMail.SmtpMail</th>
    <td><%=e.isinstall("JMail.SmtpMail")%></td>
  </tr>
  <tr>
    <th>IISSample.PageCounter</th>
    <td><%=e.isinstall("IISSample.PageCounter")%></td>
  </tr>  
  <tr>
    <th>Persits.Jpeg</th>
    <td><%=e.isinstall("Persits.Jpeg")%></td>
  </tr>  
</table>

<form action="?" method="get" id="checkForm">
<div style="width:400px;">
	<div class="input-group">
            <input name="objname" type="text" class="form-control require" placeholder="组件名称"> 
            <span class="input-group-btn">
                <input name="" type="submit" class="btn btn-primary" value="检测">
            </span>
    </div>    
    <p style="padding-top:15px;">
<%
if e.get("objname") <> "" then
  if e.isinstall(e.get("objname")) = "<i class='fa fa-check'></i>" then
    e.echo "<i class='fa fa-check'></i>" & "服务器支持该组件：" & e.get("objname")
  else
    e.echo "不支持！"
  end if
end if
%>
  </p>
</div>
</form>

</body>
</html>
<!--#include file="cms.footer.asp" -->
