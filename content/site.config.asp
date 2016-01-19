<%
'模板目录'
CONST CONFIG_TEMPLATES = "content/templates/"
'首页模板文件名'
CONST CONFIG_TEMPLATES_INDEX = "index.htm"
'生成HTML后缀名
CONST CONFIG_HTML_EXT = ".html"
'分页符'
CONST CONFIG_PAGELIST_TAG = "__"
'文件缓存文件夹'
CONST CONFIG_CACHE_PATH = "cache/"
'启用HTML文件压缩
CONST CONFIG_HTMLZIP = FALSE
'HTML分类url'
CONST CONFIG_HTML_SORT = "?{id}{ext}"
'HTML内容url'
CONST CONFIG_HTML_CONTENT = "?content-{id}{ext}"
'HTML分隔符'
CONST CONFIG_HTML_TAG = "-"
'数据库相对路径
CONST CONFIG_DATA = "content/data/e9data.mdb"
'数据库备份地址'
CONST CONFIG_DATA_BAK = "content/bak/"
'记录错误到日志'
CONST CONFIG_ERROR_LOG = TRUE
'留言评论后是否发送邮件给管理员'
CONST FEEDBACK_EMAIL = TRUE
'超级管理员名称，多个之间英文逗号隔开'
CONST CONFIG_ADMIN_SUPER = "drupal"
'MD5混淆字符串'
CONST CONFIG_MD5_STRING = "卿卿"
'注入字符串'
CONST CONFIG_SAFE_STRING = "'|;|#|([/s/b+()]+([email=select%7Cupdate%7Cinsert%7Cdelete%7Cdeclare%7C@%7Cexec%7Cdbcc%7Calter%7Cdrop%7Ccreate%7Cbackup%7Cif%7Celse%7Cend%7Cand%7Cor%7Cadd%7Cset%7Copen%7Cclose%7Cuse%7Cbegin%7Cretun%7Cas%7Cgo%7Cexists)[/s/b]select|update|insert|delete|declare|@|exec|dbcc|alter|drop|create|backup|if|else|end|and|or|add|set|open|close|use|begin|retun|as|go|exists)[/s/b[/email]+]*)"
'危险注入信息，是否记入日志'
CONST CONFIG_SAFE_LOG = TRUE
'默认排序
CONST CONFIG_WEIGHT = 99
'默认特性'
CONST CONFIG_ATTR = "特性,幻灯片,推荐"
'默认点击
CONST CONFIG_NEWS_HITS = 0
'默认文章是否显示 小于0会隐藏
CONST CONIFG_NEWS_SHOW = 0
'是否自动生成摘要
CONST CONFIG_NEWS_AUTOINFO = TRUE
'摘要截取字符数
CONST CONFIG_NEWS_AUTOINFO_LENGTH = 200
'内容缩略图格式
CONST CONFIG_PIC_EXE = "*.jpg;*.png"
'内容缩略图大小
CONST CONFIG_PIC_MAXSIZE = 512000
'缩略图保存文件夹
CONST CONFIG_PIC_PATH = "content/upload/thumail/"

'自动生成缩略图
CONST CONFIG_ASPJPEG_THUMAIL = FALSE
'生成缩略图后，原图是否删除
CONST CONFIG_ASPJPEG_THUMAIL_DELETE = FALSE
%>