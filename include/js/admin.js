$(function(){

//必填项
$('form').on("submit",function(e){
	$(this).find('.require').each(function(){
		var obj = $(this);
		var objVal = obj.val();
		if(obj.val() == '' || obj.val() == '0'){
			obj.parent().parent().addClass('has-error');
			obj.on('change',function(){
				obj.parent().parent().removeClass('has-error');
			})
			e.preventDefault();
		}
	})
})

//全选	
$('#select_all').click(function(){	
	$('.ck').prop('checked',this.checked);
	})

//显示缩略图
$('.showpic').hover(function(){				
	$(this).next('img').toggle();
})
	
})


//发送邮件的测试
function mail_test(){
	$.get("feedback.action.asp", { action: "mail_test"},
	  function(data){
	    alert(data);
	  });
}