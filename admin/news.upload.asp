<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="cms.meta.html"-->
<!--#include file="../include/upload.class.asp" -->
<!--#include file="../include/sort.class.asp" -->
<body class="abody">


<%
'===============================================================================
'创建类实例
set Upload=new AnUpLoad

'设置单个文件最大上传限制,按字节计；默认为不限制
Upload.SingleSize = CONFIG_NEWS_PIC_MAXSIZE

'设置最大上传限制,按字节计；默认为不限制
Upload.MaxSize = clng(1000 * 1024) * 1024

'设置合法扩展名,以|分割,忽略大小写
Upload.Exe="*.jpg;"

'设置文本编码，默认为gb2312
Upload.Charset="utf-8"

'获取并保存数据,必须调用本方法
Upload.GetData()
'stop

'===============================================================================
'判断错误号,如果myupload.Err<=0表示正常
if Upload.ErrorID>0 then
	e.alert(Upload.description)
	'response.Write("{err:true,msg:'" & Upload.description & "'}")
else
	if Upload.files(-1).count>0 then '这里判断你是否选择了文件
		savepath = site.dic("root") & "content/upload/thumail/"
		dim pid : pid = upload.forms("pid")
		if not isnumeric(pid) then pid = 0
		pid = cint(pid)
		if pid = 0 then e.alert("必须选择一个分类")
		sort.read(pid)
		e.echo "<center><a href='news.add.asp' class='btn btn-success'>确 定</a></center>"
		for each f in upload.files(-1)
			dim file
			set file = upload.files(f)
			if file.isfile then
				result = file.saveToFile(savepath,0,true)
				if result then
					dim title : title = upload.forms("title")
					if title = "" then
						if instr(file.LocalName,".") > 0 then
							title = split(file.LocalName,".")(0)
						else
							title = file.LocalName
						end if
						title = e.safe(title)
					end if
					guid = e.guid
					url = sort.url & guid & CONFIG_HTML_EXT
					values = sort.id 	'pid'
					values = values & ",'" & title & "',"	'title'
					values = values & "'" & savepath & file.filename & "',"	'pic'
					values = values & "'<center><img src="""&  savepath & file.filename &""" alt="""& title &"""></center>',"	'content'
					values = values & "'" & url & "',"	'url'
					values = values & CONIFG_NEWS_SHOW & ","	'show'
					values = values & CONFIG_WEIGHT & "," 'weight'
					values = values & "'"& guid &"',"
					values = values & "'"& session("admin") &"'"
					sql= "insert into [t_news](pid,title,pic,content,url,show,weight,[guid],insert_user) values("& values &")"
					ok = false
					if db.exec(sql) > 0 then ok = true
					e.echo "<p style='width:600px;margin:auto'>"
					e.echoline "状态：" & e.iif(ok,"上传成功","上传失败") & "<a href='"& site.dic("root") & url &"'>"& url &"</a>"
					e.echoline "文件：" & file.LocalName
					e.echoline "上传：" & savepath & file.filename
					e.echoline "大小：" & file.size & "字节"
					e.echoline "分类：" & sort.title
					e.echo "</p><hr>"
					'call e.alerturl("上传完成!","news.add.asp")
				else
					response.Write file.Exception
				end if
			else
				response.Write f & "<br />"
			end if
			set file = nothing
		next
		e.echo "<center><a href='news.add.asp' class='btn btn-success'>确 定</a></center>"
		set sort = nothing
	end if
end if
set Upload=nothing
%>
</body>
