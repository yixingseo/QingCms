<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/sort.class.asp" -->
<!--#include file="cms.meta.html"-->
<h3>批量添加</h3>
<form id="uploadForm" action="news.upload.asp" method="post" enctype="multipart/form-data">
<table class="table table-bordered table-admin">
  <tr>
    <th><label>内容分类</label></th>
    <td style="width:550px">
    <select name="pid" id="pid" class="form-control" style="width:250px;">
      <option value="0">根目录</option>
      <%=sort.getSelect()%>
    </select>
    <%set sort = nothing%>
    </td>
    <td><div class="tip">内容所属栏目/分类</div></td>
  </tr>
  <tr>
    <th><label>内容标题</label></th>
    <td>

    <input type="text" name="title" id="title" class="form-control" aria-label="...">

    </td>
    <td><div class="tip">内容的标题，默认使用图片名称</div></td>
  </tr>
  <tr>
    <th><label>选择图片</label></th>
    <td>
      <input id="fileImage" type="file" size="30" name="fileselect" multiple /><br>
    </td>
    <td><div class="tip">选择需要上传图片 建议使用Chrome浏览器或者Firefox浏览器</div></td>
  </tr>
  <tr>
    <th>图片列表</th>
    <td>
    <div id="preview" class="upload_preview">尚未添加图片</div>
    </td>
    <td><div class="tip">可以添加或者删除图片</div></td>
  </tr>
   <tr>
    <th>&nbsp;</th>
    <td><button type="submit" id="fileSubmit" class="upload_submit_btn">确认上传图片</button>&nbsp;&nbsp;
    <%if e.get("burl") <> "" then%><a href="<%=e.get("burl")%>">返回</a><%end if%>
    </td>
    <td></td>
  </tr>
</table>
</form>

<script>
  $(function(){
    var imgList = ''
    $('#fileImage').on('change',function(){
      var files = $(this).prop('files');
      for(var i=0;i<files.length;i++){
        imgList += '<div class="url">' + files[i].name + '<div>';
      }
      $('#preview').html(imgList);
    })

  })
</script>
</body>
</html>
<!--#include file="cms.footer.asp" -->
