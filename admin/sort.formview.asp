<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/sort.class.asp" -->
<!--#include file="cms.meta.html"-->
<script src="../include/ueditor/ueditor.config.js"></script>
<script src="../include/ueditor/ueditor.all.min.js"></script>
<%
dim id : id = e.getint("id")
if id > 0 then sort.read(id)

'如果添加小类
if e.getInt("parent_id") > 0 then
    dim sortParent : set sortParent = new ClassSort
        sortParent.read(e.getint("parent_id"))
        sort.sort_template = sortParent.sort_template
        sort.content_template = sortParent.content_template
        sort.pid = sortParent.id
end if
%>

<script>
$(function(){
	<%if sort.pid > 0 then%>
		$('#pid').val(<%=sort.pid%>);
	<%end if%>

	<%if sort.sort_template <> "" then%>
		$('#sort_template').val('<%=sort.sort_template%>');
	<%End If%>

	<%If Sort.content_template <> "" Then%>
		$('#content_template').val('<%=sort.content_template%>');
	<%End If%>

	$('#form').submit(function(){
			if($('#title').val() == ''){
					$('#title').parent().parent().addClass('has-error');
					return false;
				}
		})
	})
</script>

<h3>分类管理</h3>
<form action="sort.action.asp?action=<%=e.iif(sort.id > 0,"update","create")%>" method="post" id="form">
<table width="100%" border="0" cellspacing="1" cellpadding="0" class="table table-bordered table-hover table-admin">

  <tr>
    <th scope="row"><label>上级分类</label></th>
    <td style="width:450px;">
    <select name="pid" id="pid" class="form-control" style="width:200px;"<%if sort.id > 0 and sort.hasChild then%> disabled<%end if%>>
    	<option value="0">根目录</option>
		<%=sort.getSelect()%>
    </select>
    </td>
    <td><div class="tip">上级分类</div></td>
  </tr>

  <tr>
    <th scope="row"><label>分类名称</label></th>
    <td><input name="title" type="text" class="form-control" id="title" value="<%=Sort.Title%>" maxlength="255" /></td>
    <td><div class="tip">*分类名称</div></td>
  </tr>

<tr>
    <th><label>缩略图</label></th>
    <td>
    <div class="input-group">
      <input name="pic" type="text" id="pic" value="<%=sort.pic%>" class="form-control" />
      <div class="input-group-btn" style="position:relative">
          <input name="" type="button" class="btn btn-default" value="上传图片" id="upload_btn" onClick="selectFile()">
          <iframe class="upload_btn" src="upload.asp" style="position:absolute;top:0;z-index:999;right:0px;width:90px;height:35px;" scrolling="no" frameborder="0"></iframe>
        </div>
    </div>
    </td>
    <td>
      <div class="tip">
        <%=CONFIG_PIC_EXE%>格式，<%=CONFIG_PIC_MAXSIZE/1000%>KB以下
      </div>
    </td>
  </tr>

  <tr>
    <th scope="row"><label>URL别名</label></th>
    <td>
    	<input name="urlname" type="text" class="form-control url" id="urlname" value="<%=Sort.Urlname%>" maxlength="255" />
    </td>
    <td><div class="tip">URL名称，跳转请输入网址 http://</div></td>
  </tr>

  <tr>
    <th scope="row"><label>模板</label></th>
    <td>
    	<div class="row">
        	<div class="col-sm-6">
            	<select name="sort_template" id="sort_template" class="form-control">
                    <option value="">分类模板</option>
                    <%=Site.templateFiles()%>
                </select>
            </div>
            <div class="col-sm-6">
            	<select name="content_template" id="content_template" class="form-control">
                    <option value="">内容模板</option>
                    <%=site.templateFiles()%>
                </select>
            </div>
        </div>

		<%if sort.id > 0 and sort.hasChild then%>
        <div class="checkbox"><label><input type="checkbox" name="changeChild" value="True" id="changeChild" />修改小类模板</label></div>
        <%end if%>
    </td>
    <td>
    	<div class="tip">*列表/内容显示模板</div>
    </td>
  </tr>

  <tr>
    <th><label>分类内容</label></th>
    <td colspan="2">
    <%if sort.content = "" then%>
        <button id="loadingEditor" type="button" class="btn btn-default"><i class="fa fa-edit"></i> 使用编辑器</button>
        <script id="content" name="content" type="text/plain" style="width:760px;height:300px;"><%=sort.content%></script>
    <%else%>
        <script id="content" name="content" type="text/plain" style="width:760px;height:300px;"><%=sort.content%></script>
        <script>
            var ue = UE.getEditor('content');
        </script>
    <%end if%>
    </td>
  </tr>

  <tr>
    <th scope="row"><label>排序</label></th>
    <td>
    <input name="weight" type="text" id="weight" value="<%=sort.weight%>" class="form-control weight" />
    </td>
    <td>
    	<div class="tip">*越大排越前</div>
    </td>
  </tr>

  <tr>
    <th scope="row">&nbsp;</th>
    <td>
    <input name="id" type="hidden" value="<%=sort.id%>" />
    <button type="submit" class="btn btn-primary"> <i class="fa fa-check"></i> 保存设置</button>
    </td>
    <td></td>
  </tr>
</table>
</form>
</body>

<script>
$('#loadingEditor').click(function(){
    var ue = UE.getEditor('content');
    $(this).hide();
})
</script>
<%set sort = nothing%>
</body>
</html>
<!--#include file="cms.footer.asp" -->
