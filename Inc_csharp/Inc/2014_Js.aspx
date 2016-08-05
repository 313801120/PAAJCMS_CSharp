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
//Js

//Asp代码混淆处理 20160624
string jsCodeConfusion(string content){
    string[] splStr; int i=0; bool yesJs; bool yesWord; string sx=""; string s=""; string wc=""; string zc=""; string s1=""; string aspCode=""; string upWord="";
    string upWordn=""; string tempS=""; string dimList="";
    bool yesFunction; //函数是否为真
    bool isStartFunction; //开始函数 目的是为了让function default 处理函数后面没有()   20150218
    string wcType=""; //输入文本类型，如 " 或 '
    bool isAddToSYH; //是否累加双引号
    string beforeStr=""; string afterStr=""; string endCode=""; int nSYHCount=0;

    isStartFunction= false; //默认开始函数为假
    //If nType="" Then  nType  = 0
    yesJs= false; //是ASP 默认为假
    yesFunction= false; //是函数 默认为假
    yesWord= false; //是单词 默认为假
    nSYHCount= 0; //双引号默认为0
    splStr= aspSplit(content, vbCrlf()); //分割行
    //循环分行
    foreach(var eachs in splStr){
        s=eachs;
        //循环每个字符
        for( i= 1 ; i<= len(s); i++){
            sx= mid(s, i, 1);
            //Asp开始
            if( sx== "<" && wc== "" ){ //输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                if( mid(s, i + 1, 6)== "script" ){
                    yesJs= true; //ASP为真
                    i= i + 1; //加1而不能加2，要不然<%function Test() 就截取不到
                    sx= mid(s, i, 1);
                    aspCode= aspCode + "<";
                }
                //ASP结束
            }else if( sx== "<" && mid(s, i + 1, 8)== "/script>" && wc== "" ){ //Wc为输出内容
                yesJs= false; //ASP为假
                i= i + 1; //不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假
                sx= mid(s, i, 8);
                aspCode= aspCode + "/script>";
            }
            if( yesJs== true ){

                beforeStr= right(replace(mid(s, 1, i - 1), " ", ""), 1); //上一个字符
                afterStr= left(replace(mid(s, i + 1,-1), " ", ""), 1); //下一个字符
                endCode= mid(s, i + 1,-1); //当前字符往后面代码 一行
                //输入文本
                if((sx== "\"" || sx== "'" && wcType== "") || sx== wcType || wc != "" ){
                    isAddToSYH= true;
                    //这是一种简单的方法，等完善(20150914)
                    if( isAddToSYH== true && beforeStr== "\\" ){

                        if( len(wc) >= 1 ){
                            if( isStrTransferred(wc)== true ){ //为转义字符为真
                                //call echo(wc,isStrTransferred(wc))
                                isAddToSYH= false;
                            }
                        }else{
                            isAddToSYH= false;
                        }
                        //call echo(wc,isAddToSYH)
                    }
                    if( wc== "" ){
                        wcType= sx;
                    }

                    //双引号累加
                    if( sx== wcType && isAddToSYH== true ){ nSYHCount= nSYHCount + 1 ;}//排除上一个字符为\这个转义字符(20150914)


                    //判断是否"在最后
                    if( nSYHCount % 2== 0 && beforeStr != "\\" ){
                        if( mid(s, i + 1, 1) != wcType ){
                            wc= wc + sx;
                            aspCode= aspCode + wc; //行代码累加
                            //call echo("wc",wc)
                            nSYHCount= 0 ; wc= ""; //清除
                            wcType= "";
                        }else{
                            wc= wc + sx;
                        }
                    }else{
                        wc= wc + sx;
                    }

                }else if( sx== "'" ){ //注释则退出
                    aspCode= aspCode + mid(s, i,-1);
                    break;
                    //字母
                }else if( checkABC(sx)== true ||(sx== "_" && zc != "") || zc != "" ){
                    zc= zc + sx;
                    s1= lCase(mid(s + " ", i + 1, 1));
                    if( inStr("abcdefghijklmnopqrstuvwxyz0123456789", s1)== 0 && (s1== "_" && zc != "") ){//最简单判断
                        tempS= mid(s, i + 1,-1);

                        if( inStr("|function|sub|", "|" + lCase(zc) + "|")>0 ){
                            //函数开始
                            if( yesFunction== false && lCase(upWord) != "end" ){
                                yesFunction= true;
                                dimList= getFunDimName(tempS);
                                isStartFunction= true;
                            }else if( yesFunction== true && lCase(upWord)== "end" ){//获得上一个单词
                                yesFunction= false;
                            }
                        }else if( yesFunction== true && lCase(zc)== "var" ){
                            dimList= dimList + "," + getVarName(tempS);
                        }else if( yesFunction== true ){
                            //排除函数后面每一个名称
                            if( isStartFunction== false ){
                                zc= replaceDim2(dimList, zc);
                            }
                            isStartFunction= false;
                        }
                        upWord= zc; //记住当前单词
                        aspCode= aspCode + zc;
                        zc= "";
                    }
                }else{
                    aspCode= aspCode + sx;
                }
            }else{
                aspCode= aspCode + sx;
            }
            doEvents();
        }
        aspCode= aspRTrim(aspCode); //去除右边空格
        aspCode= aspCode + vbCrlf(); //Asp换行
        doEvents();
    }
    return aspCode;
}


//删除JS注释 20160602
string delJsNote(string content){
    string[] splStr; string s=""; string c=""; bool isMultiLineNote; string s2="";
    isMultiLineNote= false; //多行注释默认为假
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        s2= phpTrim(s);
        if( isMultiLineNote== true ){
            if( len(s2) >= 2 ){
                if( right(s2, 2)== "*/" ){
                    isMultiLineNote= false;
                }
            }
            s= "";
        }else{
            if( left(s2, 2)== "/*" ){
                if( right(s2, 2) != "*/" ){
                    isMultiLineNote= true;
                }
                s= "";
            }else if( left(s2, 2)== "//" ){
                s= "";
            }
        }
        c= c + s + vbCrlf();
    }
    return c;
}

//JS转换，引用别人



//远程网站会员统计2010330
//<script>document.writeln("<script src=\'http://127.0.0.1/web_soft/R.Asp?act=Stat&GoToUrl="+escape(document.referrer)+"&ThisUrl="+escape(window.location.href)+"&screen="+escape(window.screen.width+"x"+window.screen.height)+"&co="+escape(document.cookie)+" \'><\/script>");<'/script>
string showStatJSCode(string url){
    return "<script>document.writeln(\"<script src=\\'" + url + "act=Stat&GoToUrl=\"+escape(document.referrer)+\"&ThisUrl=\"+escape(window.location.href)+\"&screen=\"+escape(window.screen.width+\"x\"+window.screen.height)+\"&co=\"+escape(document.cookie)+\" \\'><\\/script>\");</"+"script>";
}


//Js定时跳转 Timing = 定时 时间测定 例：Call Rw("账号或密码错误，" & JsTiming("返回", 5))
string jsTiming(string url, int nSeconds){
    string c="";
    c= c + "<span id=mytimeidboyd>倒计时</span>" + vbCrlf();
    c= c + "<script type=\"text/javascript\">" + vbCrlf();
    c= c + "//配置Config" + vbCrlf();
    c= c + "var coutnumb" + vbCrlf();
    c= c + "coutnumb=" + nSeconds + "" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "//定时跳转" + vbCrlf();
    c= c + "function Countdown(){" + vbCrlf(); //Countdown=倒数计秒
    c= c + "    coutnumb-=1" + vbCrlf();
    c= c + "    mytimeidboyd.innerHTML=\"倒计时<font color=#000000>\"+coutnumb+\"</font>\"" + vbCrlf();
    c= c + "    if(coutnumb<1){    " + vbCrlf();

    if( url== "back" || url== "返回" ){ //当Action为back是返回上页
        c= c + "        history.back();" + vbCrlf();
    }else{
        c= c + "        location.href='" + url + "';" + vbCrlf();
    }


    c= c + "    }else{" + vbCrlf();
    c= c + "        setTimeout(\"Countdown()\",1000);" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}setTimeout(\"Countdown()\",1)" + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    return c;
}
//JS弹窗 Call Javascript("返回", "操作成功", "")
string javascript(string action, string msg, string url){
    if( msg != "" ){ msg= "alert('" + msg + "');" ;}//当Msg不为空则弹出信息
    if( action== "back" || action== "返回" ){ //当Action为back是返回上页
        Response.Write("<script>" + msg + "history.back();</"+"script>");
    }else if( url != "" ){ //当Url不为空
        Response.Write("<script>" + msg + "location.href='" + url + "';</"+"script>"); //跳转Url页
    }else{
        Response.Write("<script>" + msg + "</"+"script>");
    }
    Response.End();
    return "";
}
//创建Ajax对象实例
string createAjax(){
    string c="";
    c= "<script language=\"javascript\">" + vbCrlf();
    c= c + "//AjAX XMLHTTP对象实例" + vbCrlf();
    c= c + "function createAjax() { " + vbCrlf();
    c= c + "    var _xmlhttp;" + vbCrlf();
    c= c + "    try {    " + vbCrlf();
    c= c + "        _xmlhttp=new ActiveXObject(\"Microsoft.XMLHTTP\");    //IE的创建方式" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "    catch (e) {" + vbCrlf();
    c= c + "        try {" + vbCrlf();
    c= c + "            _xmlhttp=new XMLHttpRequest();    //FF等浏览器的创建方式" + vbCrlf();
    c= c + "        }" + vbCrlf();
    c= c + "        catch (e) {" + vbCrlf();
    c= c + "            _xmlhttp=false;        //如果创建失败，将返回false" + vbCrlf();
    c= c + "        }" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "    return _xmlhttp;    //返回xmlhttp对象实例" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "//Ajax" + vbCrlf();
    c= c + "function Ajax(URL,ShowID) {  " + vbCrlf();
    c= c + "    var xmlhttp=createAjax();" + vbCrlf();
    c= c + "    if (xmlhttp) {" + vbCrlf();
    c= c + "        URL+= \"&n=\"+Math.random() " + vbCrlf();
    c= c + "        xmlhttp.open('post', URL, true);//基本方法" + vbCrlf();
    c= c + "        xmlhttp.setRequestHeader(\"cache-control\",\"no-cache\"); " + vbCrlf();
    c= c + "        xmlhttp.setRequestHeader(\"Content-Type\", \"application/x-www-form-urlencoded\");         " + vbCrlf();
    c= c + "        xmlhttp.onreadystatechange=function() {        " + vbCrlf();
    c= c + "            if (xmlhttp.readyState==4 && xmlhttp.status==200) {     " + vbCrlf();
    c= c + "                document.getElementById(ShowID).innerHTML = \"操作完成\"// unescape(xmlhttp.responseText); " + vbCrlf();
    c= c + "            }" + vbCrlf();
    c= c + "            else {                " + vbCrlf();
    c= c + "                document.getElementById(ShowID).innerHTML = \"正在加载中...\"" + vbCrlf();
    c= c + "            }" + vbCrlf();
    c= c + "        }" + vbCrlf();
    //c=c & "alert(document.all.TEXTContent.value)" & vbcrlf
    c= c + "        xmlhttp.send(\"Content=\"+escape(document.all.TEXTContent.value)+\"\");    " + vbCrlf();
    c= c + "        //alert(\"网络错误\");" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "function GetIDHTML(Root){" + vbCrlf();
    c= c + "    alert(document.all[Root].innerHTML)" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    return c;
}
//JS在线编辑
string onLineEditJS(){
    string c="";
    c= c + "<script language=\"javascript\">" + vbCrlf();
    c= c + "//显示与编辑内容，但不修改，ASP代码分析器用 创作于2013,10,5" + vbCrlf();
    c= c + "function TestInput(Root){ " + vbCrlf();
    c= c + "    var TempContent" + vbCrlf();
    c= c + "    TempContent=\"\" " + vbCrlf();
    c= c + "    " + vbCrlf();
    c= c + "    document.all[Root].title=\"\"" + vbCrlf();
    c= c + "    if(document.all[Root].innerHTML.indexOf(\"<TEXTAREA\")==-1){" + vbCrlf();
    c= c + "            TempContent=document.all[Root].innerHTML" + vbCrlf();
    c= c + "            TempContent=TempContent.replace(/<BR><BR>/g,\"<BR>\");     " + vbCrlf();
    c= c + "            TempContent=TempContent.replace(/<BR>/g,\"\\n\");     " + vbCrlf();
    c= c + "            if(TempContent==\"&nbsp;\"){TempContent=\"\"}" + vbCrlf();
    c= c + "            document.all[Root].innerHTML=\"<textarea name=TEXT\"+Root+\" style='width:50%;height:50%' onblur=if(this.value!=''){document.all.\"+Root+\".title='点击可编辑';document.all.\"+Root+\".innerHTML=ReplaceNToBR(this.value)}else{document.all.\"+Root+\".innerHTML='&nbsp;'};>\" + TempContent + \"</textarea>\";" + vbCrlf();
    c= c + "            document.all[\"TEXT\"+Root].focus();" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "function ReplaceNToBR(Content){" + vbCrlf();
    c= c + "    return Content.replace(/\\n/g,\"<BR>\")" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    return c;
}
//在线编辑
string editTXT(string content, string jsId){
    content= IIF(content== "", "&nbsp;", content);
    return "<span id='" + jsId + "' onClick=\"TestInput('" + jsId + "');\" title='点击可编辑'>" + content + "</span>";
}
//在线编辑  (辅助)
string onLineEdit(string content, string jsId){
    return editTXT(content, jsId);
}
//****************************************************
//函数名：JSGoTo
//作  用：显示文本
//时  间：2013年12月14日
//参  数：Url
//*       SetTime
//返回值：字符串
//调  试：Call Echo("测试函数 JSGoTo", JSGoTo("", "",""))
//****************************************************
string jsGoTo(string title, string url, int nSetTime){
    string c="";
    if( title== "" ){ title= "添加成功" ;}
    //if nSetTime = "" then nSetTime = 4                                                '默认为4秒
    c= c + "<script>" + vbCrlf();
    c= c + "//通用定时器 如：MyTimer('Show', 'alert(1+1)', 5)" + vbCrlf();
    c= c + "var StopTimer = \"\"" + vbCrlf();
    c= c + "function MyTimer(ID, ActionStr,TimeNumb){" + vbCrlf();
    c= c + "    if(StopTimer == \"停止\" || StopTimer == \"停止定时器\"){" + vbCrlf();
    c= c + "        StopTimer = \"\"" + vbCrlf();
    c= c + "        return false" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "    TimeNumb--" + vbCrlf();
    c= c + "    document.all[ID].innerHTML = \"倒计时：\" + TimeNumb" + vbCrlf();
    c= c + "    if(TimeNumb<1){" + vbCrlf();
    c= c + "        setTimeout(ActionStr,100);" + vbCrlf();
    c= c + "    }else{" + vbCrlf();
    c= c + "        setTimeout(\"MyTimer('\"+ID+\"', '\"+ActionStr+\"',\"+TimeNumb+\")\",1000);" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "function GotoURL(){" + vbCrlf();
    c= c + "    location.href='" + url + "'" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    c= c + "<div id=\"Show\">Loading...</div><script>MyTimer('Show', 'GotoURL()', " + nSetTime + ")</"+"script>" + vbCrlf();
    return c;
}

//JS图片滚动
string jsPhotoScroll(string id, string width, string height){
    string c="";
    c= c + "<script type=\"text/javascript\">" + vbCrlf();
    c= c + "    var marqueeB = new Marquee(\"" + id + "\")    " + vbCrlf();
    c= c + "    marqueeB.Direction =2;" + vbCrlf();
    c= c + "    marqueeB.Step = 1;" + vbCrlf();
    c= c + "    marqueeB.Width = " + width + ";" + vbCrlf();
    c= c + "    marqueeB.Height = " + height + ";" + vbCrlf();
    c= c + "    marqueeB.Timer = 1;" + vbCrlf();
    c= c + "    marqueeB.DelayTime = 0;" + vbCrlf();
    c= c + "    marqueeB.WaitTime = 0;" + vbCrlf();
    c= c + "    marqueeB.ScrollStep = 20;" + vbCrlf();
    c= c + "    marqueeB.Start();    " + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    return c;
}
//图片向左滚动（暂不用）
string photoLeftScroll(string demo, string demo1, string demo2){
    string c="";
    c= c + "<!--图片向左轮番滚动-->" + vbCrlf();
    c= c + "<script language=\"javascript\">" + vbCrlf();
    c= c + "var speed=30" + vbCrlf();
    c= c + "" + demo2 + ".innerHTML=" + demo1 + ".innerHTML" + vbCrlf();
    c= c + "function Marquee(){" + vbCrlf();
    c= c + "    if(" + demo2 + ".offsetWidth-" + demo + ".scrollLeft<=0)" + vbCrlf();
    c= c + "        " + demo + ".scrollLeft-=" + demo1 + ".offsetWidth" + vbCrlf();
    c= c + "    else{" + vbCrlf();
    c= c + "        " + demo + ".scrollLeft++" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "var MyMar=setInterval(Marquee,speed)" + vbCrlf();
    c= c + "" + demo + ".onmouseover=function() {clearInterval(MyMar)}" + vbCrlf();
    c= c + "" + demo + ".onmouseout=function() {MyMar=setInterval(Marquee,speed)}" + vbCrlf();
    c= c + "</"+"script> " + vbCrlf();
    return c;
}
</script>
