<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("admin") = "" then response.redirect("login.asp")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>上传图片</title>
<style>
*{margin:0;padding:0;} 
a{text-decoration:none;} 
.btn_addPic{ 
	display: block; 
	position: relative; 
	width: 140px; 
	height: 39px; 
	overflow: hidden; 
	border: 1px solid #EBEBEB; 
	background: none repeat scroll 0 0 #F3F3F3; 
	color: #999999; 
	cursor: pointer; 
	text-align: center; 
} 
.filePrew { 
	display: block; 
	position: absolute; 
	top: 0; 
	left: 0; 
	width: 140px; 
	height: 39px; 
	font-size: 100px; /* 增大不同浏览器的可点击区域 */ 
	opacity: 0; /* 实现的关键点 */ 
	filter:alpha(opacity=0);/* 兼容IE */ 
} 
</style>
</head>

<body>
<form action="upload.action.asp?input=<%=request.querystring("input")%>" method="post" enctype="multipart/form-data" id="upload_pic_form" name="upload_pic_form">
	<a class="btn_addPic" href="javascript:void(0);"><input type="file" name="pic_upload" id="pic_upload" class="filePrew" onchange="upload();"></a>
</form>
<script type="text/javascript">
function upload(){
	document.getElementById('upload_pic_form').submit();
	}
</script>
</body>
</html>
