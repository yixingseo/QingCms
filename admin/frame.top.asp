<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="../include/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="../include/css/font-awesome.min.css">
<link rel="stylesheet" type="text/css" href="../include/css/admin.css">
</head>

<body>

<div class="header">
	<div class="logo">
    	<div class="pull-right" style="padding-top:10px;">
    		<a href="../" class="btn btn-info btn-sm" target="_blank"><i class="fa fa-home"></i> 网站首页</a>
    		<a href="action.asp?action=clear_cache" class="btn btn-success btn-sm" target="mainFrame"><i class="fa fa-refresh"></i> 清空缓存</a>
    		<a href="admin.password.asp" target="mainFrame" class="btn btn-default btn-sm"><i class="fa fa-edit"></i> <%=session("admin")%> 已登录</a>
    		<a href="javascript:parent.window.location.href='login.action.asp?exit'" class="btn btn-default btn-sm"><i class="fa fa-power-off"></i> 注销</a>    		
    	</div>
    </div>
</div>

</body>
</html>
