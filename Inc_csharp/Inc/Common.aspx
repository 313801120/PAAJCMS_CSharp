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
//常用函数大全 (2013,9,27Option Explicit)

//成功永远都是缓慢的 2013,10,4,悟

//显示弹窗对话框20150312
void msgBox( string content){
    content= replace(replace(content, chr(10).ToString(), "\\n"), chr(13).ToString(), "\\n");
    Response.Write("<script>alert('" + content + "');</"+"script>");
}
//提示20150729
void MBInfo(string title, string content){
    msgBox(title);
}
//给Queststring赋值

//ASP自带的跳转
void rr(string url){
    Response.Redirect(url);
}
//替换Request.Form对象
string rf(string str){
    return cStr(Request.Form[str]);
}
//获得传值
string rq(string str){
    return cStr(Request.QueryString[str]);
}
//获得传值
string rfq(string str){
    return cStr(Request[str]);
}
string ADSqlRf(string inputName){
    return replace(cStr(Request.Form[inputName]), "'", "''");
}
//处理成Access数据库值   辅助上面
string ADSql(string valueStr){
    return replace(valueStr, "'", "''");
}
//替换Response.Write对象
void rw(string str){
    Response.Write(str);
}
//输出内容加换行
void rwBr(string str){
    Response.Write(str + vbCrlf());
}
//替换Response.Write对象 + Response.End()
void rwEnd(string str){
    Response.Write(str);
    Response.End();
}
//HTML结束
void htmEnd(string str){
    rwEnd(str);
}
//替换Response.Write对象+Response.End()
void die(string str){
    rwEnd(str);
}
//替换Response.Write对象
void debug(string str){
    Response.Write("<div  style=\"border:solid 1px #000000;margin-bottom:2px;\">调试" + str + "</div>" + vbCrlf());
}
//跟踪
void trace(string str){
    debug(str);
}
//测试显示信息
void echo(string word, string str){
    echoPrompt(word, str);
}
//测试显示信息+红色
void echoRed(string word, string str){
    Response.Write("<font color=red>" + word + "</font>：" + str + "<br>");
}
//测试显示信息+红色+粗
void echoRedB(string word, string str){
    Response.Write("<b><font color=red>" + word + "</font>：" + str + "</b><br>");
}
//测试显示信息+黄色
void echoYellow(string word, string str){
    Response.Write("<font color=#B38704>" + word + "</font>：" + str + "<br>");
}
//测试显示信息+黄色+粗
void echoYellowB(string word, string str){
    Response.Write("<b><font color=#B38704>" + word + "</font>：" + str + "</b><br>");
}
//回显内容
void echoPrompt(string word, string str){
    Response.Write("<font color=Green>" + word + "</font>：" + str + "<br>");
}
//回显内容
void echoStr(string word, string str){
    echoPrompt(word, str);
}
//打印数组 打印PHP里用到  打印 [0] => aa [1] => bb [2] => cc

//测试显示信息 颜色
void setColorEcho(string color, string word, string str){
    Response.Write("<font color=" + color + ">" + word + "</font>：" + str + "<br>");
}
//测试显示信息暂停
void eerr(string word, string str){
    //Response.Write(TypeName(Word) & "-" & TypeName(Str)):Response.End()
    Response.Write("<font color=red>" + word + "</font>：" + str);
    Response.End();
}
//立即回显内容
void doEvents(){
    Response.Flush();
}
//功能:ASP里的IIF 如：IIf(1 = 2, "a", "b")

//Hr
void hr(){
    Response.Write("<hr size='1' color='#666666'> ");
}
//BR 20160517
void br(){
    Response.Write("<br/>");
}

//输出字符串 引用别人20141217
//Public Sub Echo(ByVal s) : Response.Write(s) : End Sub
//输出字符串和一个换行符
void print( string s){
    Response.Write(s + vbCrlf()) ; Response.Flush();
}
//输出字符串和一个html换行符
void println( string s){
    Response.Write(s + "<br />" + vbCrlf()) ; Response.Flush();
}
//输出字符串并将HTML标签转为普通字符
void printHtml( string s){
    Response.Write(replace(replace(s, "<", "&lt;"), ">", "&gt;") + vbCrlf());
}
void printlnHtml( string s){
    Response.Write(replace(replace(s, "<", "&lt;"), ">", "&gt;") + "<br />" + vbCrlf());
}
//将任意变量直接输出为字符串(Json格式)
//Public Sub PrintString(ByVal s) : Response.Write(Str.ToString(s) & VbCrLf) : End Sub
//Public Sub PrintlnString(ByVal s) : Response.Write(Str.ToString(s)) & "<br />" & VbCrLf : End Sub
//输出经过格式化的字符串
//Public Sub PrintFormat(ByVal s, ByVal f) : Response.Write(Str.Format(s, f)) & VbCrLf : End Sub
//Public Sub PrintlnFormat(ByVal s, ByVal f) : Response.Write(Str.Format(s, f)) & "<br />" & VbCrLf : End Sub
//输出字符串并终止程序运行
void printEnd( string s){
    Response.Write(s) ; Response.End();
}




//判断是否一样，一样返回checked,否者返回空值
string setIsChecked( string str, string str2){
    return IIF(str== str2,"checked='checked'","");
}
//判断是否一样，一样返回selected,否者返回空值
string setIsSelected( string str, string str2){
    return IIF(str== str2,"selected='selected'","");
}


void doError(string s, string msg){
    //On Error Resume Next
    int nRnd=0; string c="";

    nRnd= cInt(rnd() * 29252); //29252888
    c= "<br />";
    c= c + "<div style=\"width:100%; font-size:12px;;line-height:150%\">";
    c= c + "  <label onClick=\"ERRORDIV" + nRnd + ".style.display=(ERRORDIV" + nRnd + ".style.display=='none'?'':'none')\">";
    c= c + "  <span style=\"background-color:#820222;color:#FFFFFF;height:23px;font-size:14px;cursor:pointer\">〖 出错 提示信息 ERROR 〗</span><br />";
    c= c + "  </label>";
    c= c + "  <div id=\"ERRORDIV" + nRnd + "\" style=\"width:100%;border:1px solid #820222;padding:5px;overflow:hidden;\">";
    c= c + " <span style=\"color:#FF0000;\">出错描述</span> " + s + "<br />";
    c= c + " <span style=\"color:#FF0000;\">回显信息</span> " + msg + "<br />";
    c= c + "  </div>";
    c= c + "</div>";
    c= c + "<br />";
    Response.Write(c);
    Response.End(); //终止，程序停止
}

//执行ASP脚本


</script>
