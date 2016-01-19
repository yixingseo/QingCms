<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/upload.class.asp" -->
<%
'上传图片'
on error resume next
dim objUpload,strPath,strInput,result
set objUpload = new anUpload

objUpload.exe = CONFIG_PIC_EXE
objUpload.maxSize = CONFIG_PIC_MAXSIZE

objUpload.getData()

if objUpload.errorID > 0 then
	e.alert("上传失败：" & objUpload.description)
else
	dim files,savepath
	savepath = site.dic("root") & CONFIG_PIC_PATH
	set files = objUpload.files("pic_upload")
	if files.isfile then
		result = files.savetofile(savepath,0,true)
		if result then
			strPath = savepath & files.filename
		else
			alert("上传失败")
		end if
	end if
end if

strInput = e.iif(e.get("input") <> "",e.get("input"),"pic")
set objUpload = nothing
%>
<script>
<%if result then%>
	window.parent.document.getElementById('<%=strInput%>').value = '<%=strPath%>';
	document.location.href='upload.asp?input=<%=strInput%>';
<%else%>	
	alert("上传失败：" & objUpload.description)
<%end if%>
</script>
<%set objUpload = nothing%>
<!--#include file="cms.footer.asp" -->