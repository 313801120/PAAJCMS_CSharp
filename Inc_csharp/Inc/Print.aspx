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
//输出内容(2014,03,28)


//空网页内容 20160108
string getBlankHtmlBody(string webTitle, string webBody){
    string c="";
    c= "<!DOCTYPE html PUBLIC>" + vbCrlf();
    c= c + "<html xmlns=\"http://www.w3.org/1999/xhtml\">" + vbCrlf();
    c= c + "<head>" + vbCrlf();
    c= c + "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\" />" + vbCrlf();
    c= c + "<title>" + webTitle + "</title>" + vbCrlf();
    c= c + "</head>" + vbCrlf();
    c= c + "<body>" + vbCrlf();
    c= c + webBody + vbCrlf();
    c= c + "</body>" + vbCrlf();
    c= c + "</html>" + vbCrlf();
    return c;
}



//错误提示 之前网站用到
void errorText(int nRefresh, string str, string url){
    if( nRefresh>-1 ){
        rw("<meta http-equiv=\"refresh\" content=\"" + nRefresh + ";URL=" + url + "\">" + vbCrlf() );
    }
    rw("<fieldset>" + vbCrlf() );
    //rw("<legend></legend>") & vbCrLf
    rw("<div style=\"padding-left:20px;padding-top:10px;color:red;font-weight:bold;text-align:center;\">" + str + "</div>" + vbCrlf() );
    rw("<div style=\"height:200p;text-align:center;\"><P>" + vbCrlf() );
    rw("<a href=\"" + url + "\">如果您的游览器没有自动跳转,请点这里>></a><P>" + vbCrlf() );
    rw("</div></fieldset>") ; Response.End();
}

//原型显示内容
string printPre(string content){
    content= replace(content, "<", "&lt;");
    return "<pre>" + content + "</pre>";
}


//折叠菜单
string foldingMenu(int nId, string s, string msg){
    //On Error Resume Next
    int nRnd=0; string c="";
    if( nId== -1 ){

        nRnd= cInt(rnd() * 29252);	 //29252888
    }else{
        nRnd= nId;
    }
    c= "<div style=\"width:100%; font-size:12px;;line-height:150%;margin-bottom:4px;\">";
    c= c + "  <label onClick=\"ERRORDIV" + nRnd + ".style.display=(ERRORDIV" + nRnd + ".style.display=='none'?'':'none')\">";
    c= c + "  <span style=\"background-color:#666;color:#FFFFFF;height:23px;font-size:14px;cursor:pointer\">〖 " + s + " 〗</span><br />";
    c= c + "  </label>";
    c= c + "  <div id=\"ERRORDIV" + nRnd + "\" style=\"width:100%;border:1px solid #820222;padding:5px;overflow:hidden;display:none;\">";
    c= c + msg;
    c= c + "  </div>";
    c= c + "</div>";
    return c;
}
//返回标题红色
string returnRen(string Lable, string msg){
    return "<font color=red>" + Lable + "</font>：" + msg;
}
//返回Hr
string returnHr(){
    return "<hr size='1' color='#666666'> ";
}
//返回Hr
string returnRenHr(){
    return "<hr size='1' color='red'> ";
}
//处理错误信息 第二种
void showErr( string ErrCode, string ErrDesc){
    string c="";
    c= "<style>.ab-showerr{width:400px;font-size:12px;font-family:Consolas;margin:10px auto;padding:0;background-color:#FFF;}";
    c= c + ".ab-showerr h3,.ab-showerr h4{font-size:12px;margin:0;line-height:24px;text-align:center;background-color:#999;border:1px solid #555;color:#FFF;border-bottom:none;}.ab-showerr h4{padding:5px;line-height:1.5em;text-align:left;background-color:#FFC;color:#000; font-weight:normal;}";
    c= c + ".ab-showerr h4 strong{color:red;}.ab-showerr table{width:100%;margin:0;padding:0;border-collapse:collapse;border:1px solid #555;border-bottom:none;}.ab-showerr th{background-color:#EEE;white-space:nowrap;}.ab-showerr thead th{background-color:#CCC;}.ab-showerr th,.ab-showerr td{font-size:12px;border:1px solid #999;padding:6px;line-height:20px;word-break:break-all;}.ab-showerr span.info{color:#F30;}</style>";
    c= c + "<div class=\"ab-showerr\"><h3>Microsoft VBScript 编译器错误</h3><h4>程序出错了，错误代码： <strong>" + ErrCode + "</strong> ，以下是错误描述：</h4><table><tr><td>" + ErrDesc + "</td></tr></table></div>";
    Response.Write(c) ; Response.End();
}

//打印Form表单转送的值，方便写程序  因为它不能与VB软件共存  从Print.Asp引用过来



//DedeCms回显样式 20150113
string dedeCMSMsg(){
    string c="";
    c= "<style> " + vbCrlf();
    c= c + ".msgbox {" + vbCrlf();
    c= c + "    width: 450px;" + vbCrlf();
    c= c + "    border: 1px solid #DADADA;" + vbCrlf();
    c= c + "    margin:0 auto;" + vbCrlf();
    c= c + "    margin-top:20px;" + vbCrlf();
    c= c + "    line-height:20px;" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + ".msgbox .ptitle{" + vbCrlf();
    c= c + "    padding: 6px;" + vbCrlf();
    c= c + "    font-size: 12px;" + vbCrlf();
    c= c + "    border-bottom: 1px solid #DADADA;" + vbCrlf();
    c= c + "    background: #DBEEBD url(/plus/img/wbg.gif);" + vbCrlf();
    c= c + "    font-weight:bold;" + vbCrlf();
    c= c + "    text-align:center;" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + ".msgbox .pcontent{" + vbCrlf();
    c= c + "    height: 100px;" + vbCrlf();
    c= c + "    font-size: 10pt;" + vbCrlf();
    c= c + "    background: #ffffff;" + vbCrlf();
    c= c + "    text-align:center;" + vbCrlf();
    c= c + "    padding-top:30px;" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "</style> " + vbCrlf();
    c= c + "<div class=\"msgbox\">" + vbCrlf();
    c= c + "    <div class=\"ptitle\">提示信息！</div>" + vbCrlf();
    c= c + "    <div class=\"pcontent\">" + vbCrlf();
    c= c + "        成功登录，正在转向管理管理主页！<br>" + vbCrlf();
    c= c + "        <a href=\"#\">如果你的浏览器没反应，请点击这里...</a>" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "    </div>" + vbCrlf();
    c= c + "</div>" + vbCrlf();
    return c;
}
</script>
