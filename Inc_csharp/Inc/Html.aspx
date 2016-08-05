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
//Html 处理HTML代码 (2014,1,3)
//显示HTML结构        call rw(displayHTmL("<br>aasdfds<br>"))
//关闭显示HTML结构   call rwend(unDisplayHtml("&lt;br&gt;aasdfds&lt;br&gt;"))

//显示HTML结构
string displayHtml(string str){
    str= replace(str, "<", "&lt;");
    str= replace(str, ">", "&gt;");
    return str;
}
//关闭显示HTML结构
string unDisplayHtml(string str){
    str= replace(str, "&lt;", "<");
    str= replace(str, "&gt;", ">");
    return str;
}

//处理闭合HTML标签(20150902)  比上面的更好用 第二种
string handleCloseHtml(string content, bool isImgAddAlt, string action){
    int i=0; string endStr=""; string s=""; string s2=""; string c=""; string labelName=""; string startLabel=""; string endLabel="";
    action= "|" + action + "|";
    startLabel= "<";
    endLabel= ">";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        endStr= mid(content, i,-1);
        if( s== "<" ){
            if( inStr(endStr, ">") > 0 ){
                s= mid(endStr, 1, inStr(endStr, ">"));
                i= i + len(s) - 1;
                s= mid(s, 2, len(s) - 2);
                s= phpTrim(s);
                if( right(s, 1)== "/" ){
                    s= phpTrim(left(s, len(s) - 1));
                }
                endStr= right(endStr, len(endStr) - len(s) - 2); //最后字符减去当前标签  -2是因为它有<>二个字符
                //注意之前放在labelName下面
                labelName= mid(s, 1, inStr(s + " ", " ") - 1);
                labelName= lCase(labelName);
                //call eerr("s",s)

                if( inStr(action, "|处理A链接|") > 0 ){
                    s= handleHtmlAHref(s, labelName, "http://127.0.0.1/TestWeb/Web/", "处理A链接"); //处理干净html标签
                }else if( inStr(action, "|处理A链接第二种|") > 0 ){
                    s= handleHtmlAHref(s, labelName, "http://127.0.0.1/debugRemoteWeb.asp?url=", "处理A链接"); //处理干净html标签
                }
                //call echo(s,labelName)   param与embed是Flash用到，不过embed有结束标签的
                if( inStr("|meta|link|embed|param|input|img|br|hr|rect|line|area|script|div|span|a|", "|" + labelName + "|") > 0 ){
                    s= replace(replace(replace(replace(s, " class=\"\"", ""), " alt=\"\"", ""), " title=\"\"", ""), " name=\"\"", ""); //临时这么做一下，以后要完整系统的做
                    s= replace(replace(replace(replace(s, " class=''", ""), " alt=''", ""), " title=''", ""), " name=''", "");

                    //给vb.net软件用的 要不然它会报错，晕
                    if( labelName== "img" && isImgAddAlt== true ){
                        if( inStr(s, " alt")== 0 ){
                            s= s + " alt=\"\"";
                        }
                        s= aspTrim(s);
                        s= s + " /";
                        //补齐<script>20160106  暂时不能用这个，等改进
                    }else if( labelName== "script" ){
                        if( inStr(s, " type")== 0 ){
                            s= s + " type=\"text/javascript\"";
                        }
                    }else if( right(aspTrim(s), 1) != "/" && inStr("|meta|link|embed|param|input|img|br|hr|rect|line|area|", "|" + labelName + "|") > 0 ){
                        s= aspTrim(s);
                        s= s + " /";
                    }
                }
                s= startLabel + s + endLabel;
                //处理javascript script部分
                if( labelName== "script" ){
                    s2= mid(endStr, 1, inStr(endStr, "</"+"script>") + 8);

                    //call eerr("",s2)
                    i= i + len(s2);
                    s= s + s2;
                }
                //call echo("s",replace(s,"<","&lt;"))
            }
        }
        c= c + s;
    }
    return c;
}
//处理htmlA标签的Href链接  配合上面函数
string handleHtmlAHref( string content, string labelName, string addToHttpUrl, string action){
    int i=0; string s=""; string c=""; string temp="";
    bool isValue; //是否为内容值
    string valueStr=""; //存储内容值
    string yinghaoLabel=""; //引号类型如'"
    string parentName=""; //参数名称
    string behindStr=""; //后面全部字符
    string sNotDanYinShuangYinStr=""; //不是单引号和双引号字符
    action= "|" + action + "|";
    content= replace(content + " ", vbTab(), " "); //退格替换成空格，最后加一个空格，方便计算
    content= replace(replace(content, " =", "="), " =", "=");
    isValue= false; //默认内容为假，因为先是获得标签名称
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1); //获得当前一个字符
        behindStr= mid(content, i,-1); //后面字符
        if( s== "=" && isValue== false ){ //不是内容值，并为=号
            isValue= true;
            valueStr= "";
            yinghaoLabel= "";
            if( c != "" && right(c, 1) != " " ){ c= c + " " ;}
            parentName= lCase(temp); //参数名称转小写
            c= c + parentName + s;
            temp= "";
            //获得值第一个字符，因为它是引号类型
        }else if( isValue== true && yinghaoLabel== "" ){
            if( s != " " ){
                if( s != "'" && s != "\"" ){
                    sNotDanYinShuangYinStr= s; //不是单引号和双引号字符
                    s= " ";
                }
                yinghaoLabel= s;
                //call echo("yinghaoLabel",yinghaoLabel)
            }
        }else if( isValue== true && yinghaoLabel != "" ){
            //为引号结束
            if( yinghaoLabel== s ){
                isValue= false;
                if( labelName== "a" && parentName== "href" && inStr(action, "|处理A链接|") > 0 ){
                    //处理
                    if( inStr(valueStr, "?") > 0 ){
                        valueStr= replace(valueStr, "?", "WenHao") + ".html";
                    }
                    if( inStr("|asp|php|aspx|jsp|", "|" + lCase(mid(valueStr, inStrRev(valueStr, ".") + 1,-1)) + "|") > 0 ){
                        valueStr= valueStr + ".html";
                    }
                    valueStr= addToOrAddHttpUrl(addToHttpUrl, valueStr, "替换");

                }
                //call echo("labelName",labelName)
                if( yinghaoLabel== " " ){
                    c= c + "\"" + sNotDanYinShuangYinStr + valueStr + "\" "; //追加 不是单引号和双引号字符            补全
                }else{
                    c= c + yinghaoLabel + valueStr + yinghaoLabel; //追加 不是单引号和双引号字符
                }
                yinghaoLabel= "";
                sNotDanYinShuangYinStr= ""; //不是单引号和双引号字符 清空
            }else{
                valueStr= valueStr + s;
            }
            //为 分割
        }else if( s== " " ){
            //暂存内容不为空
            if( temp != "" ){
                if( left(aspTrim(behindStr) + " ", 1)== "=" ){
                    //后面一个字符等于=不处理
                }else{
                    //为标签
                    if( isValue== false ){
                        temp= lCase(temp) + " "; //标签类型名称转小写
                    }
                    c= c + temp;
                    temp= "";
                }
            }
        }else{
            temp= temp + s;
        }

    }
    c= aspTrim(c);
    return c;
}
//追加或替换网址(20150922) 配合上面   addToOrAddHttpUrl("http://127.0.0.1/aa/","http://127.0.0.1/4.asp","替换") = http://127.0.0.1/aa/4.asp
string addToOrAddHttpUrl(string httpurl, string url, string action){
    string s="";
    action= "|" + action + "|";
    if( inStr(action, "|替换|") > 0 ){
        s= getWebSite(url);
        if( s != "" ){
            url= replace(url, s, "");
        }
    }
    if( inStr(url, httpurl)== 0 ){
        if( right(httpurl, 1)== "/" &&(left(url, 1)== "/" || left(url, 1)== "\\") ){
            url= mid(url, 2,-1);
        }
        url= httpurl + url;
    }

    return url;
}

//获得HTML标签名 call rwend(getHtmlLableName("<img src><a href=>",0))    输出  img
string getHtmlLableName(string content, int nThisLabel){
    int i=0; string endStr=""; string s=""; string c=""; string labelName=""; int nLabelCount=0;
    nLabelCount= 0;
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        endStr= mid(content, i,-1);
        if( s== "<" ){
            if( inStr(endStr, ">") > 0 ){
                s= mid(endStr, 1, inStr(endStr, ">"));
                i= i + len(s) - 1;
                s= mid(s, 2, len(s) - 2);
                s= phpTrim(s);
                if( right(s, 1)== "/" ){
                    s= phpTrim(left(s, len(s) - 1));
                }
                endStr= right(endStr, len(endStr) - len(s) - 2); //最后字符减去当前标签  -2是因为它有<>二个字符
                //注意之前放在labelName下面
                labelName= mid(s, 1, inStr(s + " ", " ") - 1);
                labelName= lCase(labelName);
                if( nThisLabel== nLabelCount ){
                    break;
                }
                nLabelCount= nLabelCount + 1;
            }
        }
        c= c + s;
    }
    return labelName;
}

//删除html里空行 最笨的方法 删除空行
string removeBlankLines(string content){
    string s=""; string c=""; string[] splStr;
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( replace(replace(s, vbTab(), ""), " ", "") != "" ){
            if( c != "" ){ c= c + vbCrlf() ;}
            c= c + s;
        }
    }
    return c;
}


//call echo("webtitle",getHtmlValue(content,"webtitle"))
//call echo("webdescription",getHtmlValue(content,"webdescription"))
//call echo("webkeywords",getHtmlValue(content,"webkeywords"))
//获得html里面指定值20160520 call echo("webtitle",getHtmlValue(content,"webtitle"))
string getHtmlValue(string content, string sType){
    int i=0; string endStr=""; string s=""; string labelName=""; string startLabel=""; string endLabel=""; string LCaseEndStr=""; string paramName="";
    startLabel= "<";
    endLabel= ">";
    string getHtmlValue="";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        endStr= mid(content, i,-1);
        if( s== "<" ){
            if( inStr(endStr, ">") > 0 ){
                s= mid(endStr, 1, inStr(endStr, ">"));
                i= i + len(s) - 1;
                s= mid(s, 2, len(s) - 2);
                s= phpTrim(s);
                if( right(s, 1)== "/" ){
                    s= phpTrim(left(s, len(s) - 1));
                }
                endStr= right(endStr, len(endStr) - len(s) - 2); //最后字符减去当前标签  -2是因为它有<>二个字符
                //注意之前放在labelName下面
                labelName= mid(s, 1, inStr(s + " ", " ") - 1);
                labelName= lCase(labelName);

                if( labelName== "title" && sType== "webtitle" ){
                    LCaseEndStr= lCase(endStr);
                    if( inStr(LCaseEndStr, "</title>") > 0 ){
                        s= mid(endStr, 1, inStr(LCaseEndStr, "</title>") - 1);
                    }else{
                        s= "";
                    }
                    getHtmlValue= s;
                    return getHtmlValue;
                }else if( labelName== "meta" &&(sType== "webkeywords" || sType== "webdescription") ){
                    LCaseEndStr= lCase(endStr);
                    paramName= phpTrim(lCase(getParamValue(s, "name")));
                    if( "web" + paramName== sType ){
                        getHtmlValue= getParamValue(s, "content");
                        return getHtmlValue;
                    }


                }

            }
        }
    }
    return "";
}

//获得参数值20160520  call rwend(getParamValue("meta name=""keywords"" content=""{$web_keywords$}""","name"))
string getParamValue(string content, string paramName){
    string LCaseContent=""; string s="";   int i=0; string startStr=""; string endStr="";
    LCaseContent= lCase(content);
    string getParamValue="";

    string[] sArrayStart= {"=\"", "='", "="};
    string[] sArrayEnd= {"\"", "'", ">"};
    for( i= 0 ; i<= uBound(sArrayStart); i++){
        startStr= paramName + sArrayStart[i];
        endStr= sArrayEnd[i];
        if( inStr(LCaseContent, startStr) > 0 && inStr(LCaseContent, endStr) > 0 ){
            s= strCut(content, startStr, endStr, 2);
            if( s != "" ){
                getParamValue= s;
                return getParamValue;
            }
        }
    }
    return getParamValue;
}

</script>
