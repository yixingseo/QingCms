<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
<meta name="renderer" content="webkit">
<title>管理员登录-企业网站管理系统</title>
<link rel="stylesheet" type="text/css" href="../include/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="../include/css/font-awesome.min.css">
<link rel="stylesheet" type="text/css" href="../include/css/admin.css">
<script src="../include/js/jquery-1.7.2.min.js"></script>
</head>

<body>
<!--有任何问题可以联系QQ：66866111 或者发邮件66866111@qq.com 交流-->
<div class="header">
	<div class="logo">
    	<img src="../include/images/logo.jpg" alt="企业网站管理系统">       
    </div>
</div>

<form action="login.action.asp?login" method="post" id="loginForm">
<div class="login">
    <div class="panel panel-default">
		<div class="panel-heading">
        	<div class="panel-title"><i class="fa fa-user"></i> <strong>管理员登陆</strong></div>
        </div>        
        <div class="panel-body">
        	<div class="form-group">
            	<label for="username">用户名：</label>
                <input name="username" type="text" id="username" size="20" maxlength="20" class="form-control" />
                <p class="help-block">请输入用户名</p>
            </div>
            <div class="form-group">
            	<label for="password">密码：</label>
                <input name="password" type="password" id="password" size="20" maxlength="20" class="form-control" />
                <p class="help-block">请输入密码</p>
            </div>
            <div class="form-group">
            	<input type="submit" class="btn btn-primary btn-block" value="登录系统 LOGIN" />
            </div>
		</div>
        <!--[if lt IE 9]>
        <div class="panel-footer">
        	<div class="text-danger text-center"><a href="http://windows.microsoft.com/zh-CN/internet-explorer/downloads/ie?ocid=ie6_countdown_bannercode"><i class="fa fa-download"></i> 您使用的浏览器版本过低，可以考虑升级了。</a></div>
        </div>
        <![endif]-->
	</div>
    	    
</div>
</form>  

<script>
$('#loginForm').submit(function(){
		if( $('#username').val() == ''){
			$("#username").parent(".form-group").addClass("has-error");
			$("#username").focus();
			return false
			}
		if( $('#password').val() == ''){
			$("#password").parent(".form-group").addClass("has-error");
			$("#password").focus();
			return false
			}					
	})
</script>

<!--[if lt IE 9]>
<script src="http://cdn.bootcss.com/html5shiv/3.7.2/html5shiv.min.js"></script>
<script src="http://cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>
<![endif]-->    
</body>
</html>