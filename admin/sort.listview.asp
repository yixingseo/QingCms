<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/sort.class.asp" -->
<!--#include file="cms.meta.html"-->
<%
dim arrayList : arrayList = sort.getRows()
%>
<h3>分类管理</h3>
<div class="tools">
<div class="form-inline">
  <div class="form-group"><a class="btn btn-primary" href="sort.formview.asp"><i class="fa fa-plus-circle"></i> 添加分类</a></div>
</div>
</div>
<table width="100%" class="table table-bordered table-hover">
  <tr>
    <th style="width:80px;">编号</th>
    <th>分类名称</th>
    <th style="width:70px;">浏览</th>
    <th>分类模板</th>
    <th>内容模板</th>
    <th>排序</th>
    <th>操作</th>
  </tr>
	<%=listview_sort(arrayList,0,0)%>
</table>

</body>

<%
'归递分类
Function listview_sort(arrayList,Pid,K)
	Dim i
	For i = 0 To Ubound(arrayList,2)
		if cint(arrayList(2,i)) = cint(pid) then
			listview_sort = listview_sort & "<tr class='"& E.IIF(Cint(arrayList(2,i)) = 0, "active" ,"") &"'>" &vbcrlf
			'id
			listview_sort = listview_sort & "<td class=""listview_id"">"& arrayList(0,i) &"</td>" &vbcrlf
			'title
			listview_sort = listview_sort & "<td>"& GetTag(k) & "<span class=""ajax_title"" sortid="""& arrayList(0,i) &""">" & arrayList(1,i) &"</span></td>" &vbcrlf
			listview_sort = listview_sort & "<td><a href='../"& arrayList(3,i) &"' target='_blank'><i class='fa fa-link'></i> 浏览</a></td>" &vbcrlf
			'sort_template
			listview_sort = listview_sort & "<td><span class=""url"">&nbsp;"& arrayList(5,i) &"</span></td>" &vbcrlf
			'content_template
			listview_sort = listview_sort & "<td><span class=""url"">&nbsp;"& arrayList(6,i) &"</span></td>" &vbcrlf
			'weight
			listview_sort = listview_sort & "<td class=""ajax_weight"" sortid="""& arrayList(0,i) &""">"& arrayList(11,i) &"</td>" &vbcrlf
			'action
			listview_sort = listview_sort & "<td class=""listview_action"">"
			listview_sort = listview_sort & "<a href='sort.formview.asp?parent_id="& arrayList(0,i) &"'><i class='fa fa-plus-circle'></i> 添加小类</a>&nbsp;&nbsp;"
			listview_sort = listview_sort & "<a href='sort.formview.asp?id="& arrayList(0,i) &"'><i class='fa fa-pencil'></i> 修改</a>&nbsp;&nbsp;"
			listview_sort = listview_sort & "<a href='sort.action.asp?action=delete&id="& arrayList(0,i) &"' onclick=""return confirm('确定要删除吗？')""><i class='fa fa-trash'></i> 删除</a>"
			listview_sort = listview_sort & "</td>" &vbcrlf
			listview_sort = listview_sort & "</tr>" &vbcrlf
			listview_sort = listview_sort & listview_sort(arrayList,arrayList(0,i),k + 1)
		end if
	next
End Function

function GetTag(k)
	if k = 0 then
		GetTag = "┋┅" & "&nbsp;"
		exit function
	end if
	dim i
	for i = 0 to k
		GetTag = GetTag & "&nbsp;&nbsp;"
	next
	GetTag = GetTag & "┊┈" & "&nbsp;"
end function
%>
<script>
$(document).ready(function() {
	//ajax_weight
	var $tds = $(".ajax_weight");
	$tds.click(function() {
				var $td = $(this);
				var sortid = $td.attr('sortid');
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
								$.get("sort.action.asp", { action:'ajax_weight',id: sortid, weight: value } );
							} else if (key == 27) {
								$td.html(text);
							}
						});
			});

	//ajax_title
	var $tds = $(".ajax_title");
	$tds.click(function() {
				var $td = $(this);
				var sortid = $td.attr('sortid');
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
								$.get("sort.action.asp", { action:'ajax_title',id: sortid, title: value } );
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
