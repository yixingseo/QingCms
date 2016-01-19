<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%response.codepage = 65001%>
<%response.charset = "utf-8"%>
<%dim startTimer : starTimer = timer()%>
<%'on error resume next%>
<!--#include file="content/site.config.asp" -->
<!--#include file="include/ide.class.asp" -->
<!--#include file="include/db.class.asp" -->
<%db.path = CONFIG_DATA%>
<!--#include file="include/site.class.asp" -->
<!--#include file="include/log.class.asp" -->
<!--#include file="include/sort.class.asp" -->
<!--#include file="include/news.class.asp" -->
<!--#include file="include/tpl.class.asp" -->
<!--#include file="include/feedback.class.asp" -->
<%
if site.dic("enable") = "false" then e.die "网站暂时关闭，请联系管理员。"
dim tpl : set tpl = new class_tpl
call tpl.route()
set tpl = nothing : set db = nothing : set e =  nothing
%>
