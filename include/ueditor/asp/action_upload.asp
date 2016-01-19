<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="ASPJson.class.asp"-->
<!--#include file="config_loader.asp"-->
<!--#include file="Uploader.class.asp"-->
<%
	'获取根目录
	response.Charset = "utf-8"
	response.Codepage = 65001
	Dim strUrl : strUrl = Lcase(Request.ServerVariables("SCRIPT_NAME"))
	Dim rootPath : rootPath = Split(strUrl,"include/")(0)	
		rootPath =  rootPath & "content/"


    uploadTemplateName = Session.Value("ueditor_asp_uploadTemplateName")
	
    Set up = new Uploader
    up.MaxSize = config.Item( uploadTemplateName & "MaxSize" )
    up.FileField = config.Item( uploadTemplateName & "FieldName" )
    up.PathFormat = rootPath & config.Item( uploadTemplateName & "PathFormat" )

    If Not IsEmpty( Session.Value("base64Upload") ) Then
        up.UploadBase64( Session.Value("base64Upload") )
    Else
        up.AllowType = config.Item( uploadTemplateName & "AllowFiles" )
        up.UploadForm()
    End If

    Set json = new ASPJson

    With json.data
        .Add "url", up.FilePath
        .Add "original", up.OriginalFileName
        .Add "state", up.State
        .Add "title", up.OriginalFileName
    End With
    
    json.PrintJson()
%>