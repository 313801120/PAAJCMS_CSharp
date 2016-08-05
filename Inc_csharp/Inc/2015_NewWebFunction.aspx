<%
/************************************************************
作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
版权：源代码免费公开，各种用途均可使用。 
创建：2016-08-05
联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
*                                    Powered by PAAJCMS 
************************************************************/
%>
<script runat="server" language="c#">
//新网站函数

//特殊字符替换
string specialStrReplace( string content){
    content= replace(content, "\\|", "[$特殊字符A]$");
    content= replace(content, "\\-", "[$特殊字符B]$");
    content= replace(content, "\\,", "[$特殊字符C]$");
    content= replace(content, "\\'", "[$特殊字符D]$");
    content= replace(content, "\\\"", "[$特殊字符E]$");
    return content;
}
//解特殊字符替换
string unSpecialStrReplace( string content, string startStr){
    content= replace(content, "[$特殊字符A]$", startStr + "|");
    content= replace(content, "[$特殊字符B]$", startStr + "-");
    content= replace(content, "[$特殊字符C]$", startStr + ",");
    content= replace(content, "[$特殊字符D]$", startStr + "'");
    content= replace(content, "[$特殊字符E]$", startStr + "\"");
    return content;
}

//栏目类型处理 首页|文本|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单
string handleColumnType(string columnName){
    string s="";
    switch ( columnName ){
        case "首页" : s= "home";break;
        case "文本" : s= "text";break;
        case "产品" : s= "product";break;
        case "新闻" : s= "news";break;
        case "视频" : s= "video";break;
        case "下载" : s= "download";break;
        case "案例" : s= "case";break;
        case "留言" : s= "message";break;
        case "反馈" : s= "feedback";break;
        case "招聘" : s= "job";break;
        case "订单" : s= "order";
        break;
    }
    return s;
}

</script>
