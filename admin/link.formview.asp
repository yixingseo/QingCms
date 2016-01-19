<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/link.class.asp" -->
<!--#include file="cms.meta.html"-->
<%
dim id : id = e.getInt("id")
if id > 0 then link.read(id)
%>

<h3>链接管理</h3>
<form action="link.action.asp?action=<%=e.iif(link.id > 0,"update","create")%>&backurl=link.listview.asp" method="post" id="linkform">
<table width="100%" border="0" cellspacing="1" cellpadding="0" class="table table-bordered table-hover table-admin">

  <tr>
    <th scope="row"><label>链接名称</label></th>
    <td style="width:550px"><input name="title" type="text" class="form-control require" id="title" value="<%=link.Title%>" maxlength="255" /></td>
    <td><div class="tip">*链接名称</div></td>
  </tr>

  <tr>
    <th scope="row"><label>链接网址</label></th>
    <td>
    	<input name="url" type="text" class="form-control url require" id="url" value="<%=link.url%>" maxlength="255" />
    </td>
    <td><div class="tip">请输入网址 http://xxx.com</div></td>
  </tr>

  <tr>
    <th><label>缩略图</label></th>
    <td>
    <div class="input-group">
    <input name="pic" type="text" id="pic" value="<%=link.pic%>" class="form-control" />
        <div class="input-group-btn" style="position:relative">
            <input name="" type="button" class="btn btn-default" value="上传图片" id="upload_btn" onClick="selectFile()">
            <iframe class="upload_btn" src="upload.asp" style="position:absolute;top:0;z-index:999;right:0px;width:90px;height:35px;" scrolling="no" frameborder="0"></iframe>
        </div>
    </div>
    </td>
    <td></td>
  </tr>

<tr>
    <th scope="row"><label>打开方式</label></th>
    <td>
        <div class="radio">
            <label>
                <input type="radio" name="target" id="target" value="_blank" <%if link.target="_blank" then%>checked<%end if%>>新窗口
            </label>
            <label>
                <input type="radio" name="target" id="target1" value="_self" <%if link.target<>"_blank" then%>checked<%end if%>>原窗口
            </label>
        </div>
    </td>
    <td><div class="tip">打开方式</div></td>
  </tr>

  <tr>
    <th scope="row"><label>排序</label></th>
    <td>
    <input name="weight" type="text" id="weight" value="<%=link.weight%>" class="form-control weight" />
    </td>
    <td>
    	<div class="tip">*越大排越前</div>
    </td>
  </tr>

  <tr>
    <th scope="row">&nbsp;</th>
    <td>
    <input name="id" type="hidden" value="<%=link.id%>" />
    <button type="submit" class="btn btn-primary"> <i class="fa fa-check"></i> 保存设置</button>
    </td>
    <td></td>
  </tr>
</table>
</form>
</body>
</html>
<!--#include file="cms.footer.asp" -->
