<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/feedback.class.asp" -->
<!--#include file="../include/link.class.asp" -->

<%
dim action : action = e.get("action")

select case lcase(action)

	case "clear_cache"
		on error resume next
		e.echoline "缓存文件夹：" & server.mappath(site.dic("root") & CONFIG_CACHE_PATH)
		dim objFso,o_dir,o_files
		set objFso = e.objFso
		set	o_dir = objFso.getfolder(server.mappath(site.dic("root") & CONFIG_CACHE_PATH))
		set	o_files = o_dir.files
		dim f
		for each f in o_files
			e.echoline "正在删除：" & f.name
			if objFso.FileExists(f.path) then objFso.deletefile(f.path)
			if not objFso.FileExists(f.path) then e.echo "-OK"
		next
		e.alert("清空完成")

	case "bak"
		dim str_file
			str_file = site.dic("root") & CONFIG_DATA_BAK & e.guid & ".bak"
		On Error Resume Next
		call e.file_copy(site.dic("root") & CONFIG_DATA, str_file ,1)
		if err then
			e.alert "备份失败，可能是磁盘空间不够"
		else
			log.add "备份数据"
			e.alerturl "备份成功","data.listview.asp"
		end if

	case "delbak"
		strFile = site.dic("root") & CONFIG_DATA_BAK & E.Get("file")
		if e.file_delete(strFile) then
		else
			if err.number > 0 then e.alert("删除失败，可能权限不够")
		end if
		log.add "删除备份"

	case "down"
		if not site.super_user then e.alert("权限不够") : response.end()
		response.redirect site.dic("root") & CONFIG_DATA_BAK & e.get("file")

	case "zip"
		strFile = site.dic("root") & CONFIG_DATA
		Const strConnPartial = "provider=microsoft.jet.oledb.4.0; data source="
		Dim oJE : Set oJE = Server.CreateObject("jro.JetEngine")
		Dim oFso : Set oFso = E.objFso
		Db.Conn.Close
		Dim strNewFile
		strFile = Server.Mappath(strFile)
		If Not oFso.FileExists(strFile) Then E.Alert("数据库文件不存在")
		strNewFile = strFile & ".tmp"
		If oFso.FileExists(strNewFile) Then oFso.DeleteFile(strNewFile)
		oJE.CompactDatabase strConnPartial & strFile,strConnPartial & strNewFile
		oFso.DeleteFile strFile
		oFso.MoveFile strNewFile,strFile
		Set oFso = Nothing
		Set oJE = Nothing
		e.alerturl "压缩成功","data.listview.asp"

end select

redirect "action.asp"
%>

<!--#include file="cms.footer.asp" -->
