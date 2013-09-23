// JavaScript Document
$(document).ready(function() {
    $(window).scroll(function() {    // 页面发生scroll事件时触发  
        var bodyTop = 0;
        if (typeof window.pageYOffset != 'undefined') {
            bodyTop = window.pageYOffset;
        } else if (typeof document.compatMode != 'undefined' && document.compatMode != 'BackCompat') {
            bodyTop = document.documentElement.scrollTop;
        }
        else if (typeof document.body != 'undefined') {
            bodyTop = document.body.scrollTop;
        }

        var mainHeight = $("#footerf").offset().top;

        var divTop = $("#navfuntop").offset().top;
        var divHeight = $("#navfun").height();
        var funTop = divTop + bodyTop;

        if ((funTop + divHeight) > mainHeight) {
            funTop = mainHeight - divHeight - 10;
        }

        $("#navfun").css("top", funTop);   // 设置层的CSS样式中的top属性, 注意要是小写，要符合“标准”  
        //$("#navfun").text(funTop);   // 设置层的内容，这里只是显示当前的scrollTop  
    });

    $("#navfun").html("<div id=\"navfun1\"><a href=\"tencent://message/?uin=954904356&amp;Site=沭阳亚菲&amp;Menu=yes\"><img src=\"Images/funQQ.gif\" alt=\"客服QQ\" width=\"109\" height=\"50\"  border=\"0\" /></a><br /><br /><a href=\"tencent://message/?uin=954904356&amp;Site=沭阳亚菲&amp;Menu=yes\"><img src=\"Images/funQQ.gif\" alt=\"设计QQ\" width=\"109\" height=\"50\"  border=\"0\" /></a><br /></div>");
});