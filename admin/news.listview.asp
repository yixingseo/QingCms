<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/sort.class.asp" -->
<!--#include file="cms.meta.html"-->
<%
sort.id = e.getInt("sort_id")
if sort.id > 0 then sort.read(sort.id)

dim strSql
strSql = "select [news.id] as id,[news.title] as title,[news.url] as url,[news.pic] as pic,[news.att] as att,[news.hits] as hits,[news.weight] as weight,[news.insert_date] as insert_date,[sort.id] as sortID,[sort.title] as sortTitle"
strSql = strSql & " from [t_news] as news left join [t_sort] as sort on news.pid = sort.id"

dim sqlWhere,strKeywords
	strKeywords = e.get("search_key")
	strKeywords = e.safe(strKeywords)
	sqlWhere = " news.show > -1"
	sqlWhere = sqlWhere & e.iif(sort.id > 0," and [news.pid] in("& sort.getChild() &")","")
	sqlWhere = sqlWhere & e.iif(strKeywords <> ""," and [news.title] like('%"& strKeywords &"%')","")
	sqlWhere = sqlWhere & e.iif(e.get("att") <> ""," and instr([news.att],'"& E.Safe(E.Get("att")) &"') > 0","")
	strSql = strSql & " where " & sqlWhere & " order by [news.weight] desc,[news.id] desc"
%>
<form action="news.listview.asp" method="get">
<h3>内容管理</h3>
<!-- 操作 -->
<div class="tools">
	<div class="form-inline">

		<div class="form-group">
			<button type="button" class="btn btn-danger" id="del_all"><i class="fa fa-recycle"></i> 回收站</button>
			<a class="btn btn-success" href="news.add.asp"><i class="fa fa-plus-circle"></i> 批量添加</a>
			<a class="btn btn-primary" href="news.formview.asp"><i class="fa fa-plus-circle"></i> 添加内容</a>
		</div>

		<div class="form-group">
			<div class="input-group">
					<input name="search_key" type="text" class="form-control">
					<span class="input-group-btn">
							<input name="search_btn" type="submit" class="btn btn-default" value="搜索">
					</span>
			</div>
		</div>

		<div class="form-group">
			<select name="sort_jump_box" id="sort_jump_box" class="form-control">
      	<option value="0">按分类查看</option>
      	<%=sort.getSelect()%>
      </select>
		</div>

		<div class="form-group">
				<select name="att_jump_box" id="att_jump_box" class="form-control">
        	<option value="">按特性查看</option>
        	<%
					attArray =  site.newsAttArray
					for k = 0 To Ubound(attArray)%>
        		<option value="<%=attArray(k)%>"><%=attArray(k)%></option>
        	<%next%>
        </select>
		</div>

	</div>
</div>
<!-- 操作栏 -->
</form>

<%
dim rs : set rs = db.listpage(strSql,20)
i = 0
%>
<table class="table table-bordered table-hover">
  <tr>
  	<th width="50" class="text-center"><input name="select_all" id="select_all" type="checkbox"></th>
    <th style="width:70px">编号</th>
    <th>标题</th>
    <th style="width:70px">浏览</th>
    <th style="width:50px" class="text-center"><i class="fa fa-photo"></i</th>
    <th class="text-center">分类</th>
    <th class="text-center">特性</th>
    <th>点击</th>
    <th>排序</th>
    <th>操作</th>
  </tr>
<%
while not rs.eof and i < rs.pagesize
%>
  <tr id="<%=rs("id")%>">
  	<td align="center"><input name="ids" type="checkbox" class="ck" value="<%=rs("id")%>" /></td>
    <td class="listview_id"><%=rs("id")%></td>
    <td><span class="ajax_title" newsid="<%=rs("id")%>"><%=e.encode(e.cutString(rs("title"),40))%></span></td>
    <td><a href="<%if left(rs("url"),1) = "?" then%>../<%=rs("url")%><%else%><%=rs("url")%><%end if%>" target="_blank"><i class="fa fa-link"></i> 浏览</a></td>
		<td class="text-center"><%if rs("pic") <> "" then%><i class="fa fa-photo showpic"></i> <img src="<%=rs("pic")%>" class="list_thumail"><%end if%> &nbsp;</td>
    <td align="center"><a href="?pid=<%=rs("sortID")%>"><%=E.Iif(rs("sortTitle") <> "",e.encode(rs("sortTitle")),"根目录")%></a></td>
    <td>
<%
for k = 0 To Ubound(attArray)
	If InStr(rs("att"),attArray(k)) > 0 Then
	%>
	<a href="news.action.asp?action=delatt&id=<%=rs("id")%>&att=<%=Server.UrlEncode(attArray(k))%>&redirect=<%=e.getUrl%>"><%=attArray(k)%></a>
	<%else%>
	<a href="news.action.asp?action=addatt&id=<%=rs("id")%>&att=<%=Server.UrlEncode(attArray(k))%>&redirect=<%=e.getUrl%>" class="att-false"><%=attArray(k)%></a>
<%
	end if
next
%>
    </td>
    <td><%=rs("Hits")%></td>
    <td class="ajax_weight" newsid="<%=rs("id")%>"><%=rs("Weight")%></td>
    <td class="listview_action">
    	<a href="news.formview.asp?id=<%=rs("id")%>&redirect=<%=e.getUrl%>"><i class="fa fa-pencil"></i> 修改</a>&nbsp;&nbsp;
    	<a href="news.action.asp?action=rec&id=<%=rs("id")%>&redirect=<%=e.getUrl%>" onclick="return confirm('确定要删除吗')"><i class="fa fa-recycle"></i> 回收站</a>
    </td>
  </tr>
<%
i = i + 1
rs.Movenext
Wend
%>
</table>


<nav class="pagelist"><%=Db.GetPager("")%></nav>
<%
rs.Close
Set rs = Nothing
%>


<script>
$(document).ready(function() {
	$('.showpic').hover(function(){
		//alert($(this).next('img'))	;
		$(this).next('img').show();
	},function(){
		$(this).next('img').hidden();
	})

	//批量删除
	$('#del_all').click(function(){
			//alert('删除选中');
			var ids = "";
			$("input[name='ids']").each(function(){
				if($(this).prop("checked")){
					ids += $(this).val() + ",";
				}
			})
			if(ids == "")
				return;
			//alert(ids);
			$.get("news.action.asp", { action:'rec',id: ids} ,function(){
				document.location.href='<%=e.getUrl%>'
			});
		})

	//分类跳转
	$('#sort_jump_box').change(function(){
			document.location.href='news.listview.asp?sort_id=' + $(this).val();
		})

	<%If Sort.ID > 0 Then%>
	$('#sort_jump_box').val('<%=Sort.ID%>');
	<%End If%>

	//特性跳转
	$('#att_jump_box').change(function(){
			document.location.href='news.listview.asp?att=' + $(this).val();
	})
	<%If Request.QueryString("att") <> "" Then%>
	$("#att_jump_box").val('<%=e.get("att")%>')
	<%End If%>

	//weight ajax
	var $tds = $(".ajax_weight");
	$tds.click(function() {
				var $td = $(this);
				var newsid = $td.attr('newsid');
				if ($td.children("input").length > 0) {
					return false;
				}
				var text = $td.text();
				var $input = $("<input type='text'>").css("background-color",'#FFFFFF').width('70px');
				$input.val(text);
				$td.html("");
				$td.append($input);

				$input.keyup(function(event) {
							var key = event.which;
							if (key == 13) {
								var value = $input.val();
								$td.html(value);
								$.get("news.action.asp", { action:'ajax_weight',id: newsid, weight: value } );
							} else if (key == 27) {
								$td.html(text);
							}
						});
			});

	//title ajax
	var $tds = $(".ajax_title");
	$tds.click(function() {
				var $td = $(this);
				var newsid = $td.attr('newsid');
				if ($td.children("input").length > 0) {
					return false;
				}
				var text = $td.text();
				var $input = $("<input type='text'>").css("background-color",'#FFFFFF').width('220px');
				$input.val(text);
				$td.html("");
				$td.append($input);

				$input.keyup(function(event) {
							var key = event.which;
							if (key == 13) {
								var value = $input.val();
								$td.html(value);
								$.get("news.action.asp", { action:'ajax_title',id: newsid, title: value});
							} else if (key == 27) {
								$td.html(text);
							}
						});
			});
});
</script>
</body>
</html>
<!--#include file="cms.footer.asp" -->
