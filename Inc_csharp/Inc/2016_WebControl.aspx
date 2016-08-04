<%
/************************************************************
作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
版权：源代码免费公开，各种用途均可使用。 
创建：2016-08-04
联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
*                                    Powered by PAAJCMS 
************************************************************/
%>
<script runat="server" language="c#">
//网站控制 20160223



//处理模块替换数组


//去掉模板里不需要显示内容 删除模板中我的注释代码
string delTemplateMyNote(string code){
    string startStr=""; string endStr=""; int i=0; string s=""; int nHandleNumb=0; string[] splStr; string block=""; string id="";
    string content=""; string dragSortCssStr=""; string dragSortStart=""; string dragSortEnd=""; string dragSortValue=""; string c="";
    string lableName=""; string lableStartStr=""; string lableEndStr="";
    nHandleNumb= 99; //这里定义很重要

    //加强版  对这个也可以<!--#aaa start#--><!--#aaa end#-->
    startStr= "<!--#" ; endStr= "#-->";
    for( i= 1 ; i<= nHandleNumb; i++){
        if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
            lableName= strCut(code, startStr, endStr, 2);
            if( inStr(lableName, " start") > 0 ){
                lableName= mid(lableName, 1, len(lableName) - 6);
            }

            s= startStr + lableName + endStr;
            lableStartStr= startStr + lableName + " start" + endStr;
            lableEndStr= startStr + lableName + " end" + endStr;
            if( inStr(code, lableStartStr) > 0 && inStr(code, lableEndStr) > 0 ){
                s= strCut(code, lableStartStr, lableEndStr, 1);
                //call echo(">>",s)
            }
            code= replace(code, s, "");
            //call echo("s",s)
            //call echo("lableName",lableName)
            //call echo("lableStartStr",replace(lableStartStr,"<","&lt;"))
            //call echo("lableEndStr",replace(lableEndStr,"<","&lt;"))
        }else{
            break;
        }
    }



    //清除ReadBlockList读出块列表内容  不过有个不足的地方，读出内容可以从外部读出内容，这个以后考虑
    //Call Eerr("ReadBlockList",ReadBlockList)
    //写于20141118
    //splStr = Split(ReadBlockList, vbCrLf)                 '不用这种，复杂了
    //修改于20151230
    for( i= 1 ; i<= nHandleNumb; i++){
        startStr= "<R#读出内容" ; endStr= " start#>";
        block= strCut(code, startStr, endStr, 2);
        if( block != "" ){
            startStr= "<R#读出内容" + block + " start#>" ; endStr= "<R#读出内容" + block + " end#>";
            if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
                s= strCut(code, startStr, endStr, 1);
                code= replace(code, s, ""); //移除
            }
        }else{
            break;
        }
    }

    //删除翻页配置20160309
    startStr= "<!--#list start#-->";
    endStr= "<!--#list end#-->";
    if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
        s= strCut(code, startStr, endStr, 2);
        code= replace(code, s, "");
    }

    if( cStr(Request["gl"])== "yun" ){
        content= getFText("/Jquery/dragsort/Config.html");
        content= getFText("/Jquery/dragsort/模块拖拽.html");
        //Css样式
        startStr= "<style>";
        endStr= "</style>";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            dragSortCssStr= strCut(content, startStr, endStr, 1);
        }
        //开始部分
        startStr= "<!--#top start#-->";
        endStr= "<!--#top end#-->";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            dragSortStart= strCut(content, startStr, endStr, 2);
        }
        //结束部分
        startStr= "<!--#foot start#-->";
        endStr= "<!--#foot end#-->";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            dragSortEnd= strCut(content, startStr, endStr, 2);
        }
        //显示块内容
        startStr= "<!--#value start#-->";
        endStr= "<!--#value end#-->";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            dragSortValue= strCut(content, startStr, endStr, 2);
        }



        //控制处理
        startStr= "<dIv datid='";
        endStr= "</dIv>";
        content= getArray(code, startStr, endStr, false, false);
        splStr= aspSplit(content, "$Array$");
        foreach(var eachs in splStr){
            s=eachs;
            startStr= "【DatId】'";
            id= mid(s, 1, inStr(s, startStr) - 1);
            s= mid(s, inStr(s, startStr) + len(startStr),-1);
            //C=C & "<li><div title='"& Id &"'>" & vbcrlf & "<div " & S & "</div>"& vbcrlf &"<div class='clear'></div></div><div class='clear'></div></li>"
            s= "<div" + s + "</div>";
            //Call Die(S)
            c= c + replace(replace(dragSortValue, "{$value$}", s), "{$id$", id);
        }
        c= replace(c, "【换行】", vbCrlf());
        c= dragSortStart + c + dragSortEnd;
        code= mid(code, 1, inStr(code, "<body>") - 1);
        code= replace(code, "</head>", dragSortCssStr + "</head></body>" + c + "</body></html>");
    }

    //删除VB软件生成的垃圾代码
    startStr= "<dIv datid='" ; endStr= "【DatId】'";
    for( i= 1 ; i<= nHandleNumb; i++){
        if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
            id= strCut(code, startStr, endStr, 2);
            code= replace2(code, startStr + id + endStr, "<div ");
        }else{
            break;
        }
    }
    code= replace(code, "</dIv>", "</div>"); //替换成这个结束div

    //最外围清除
    startStr= "<!--#dialogteststart#-->" ; endStr= "<!--#dialogtestend#-->";
    code= replace(code, "<!--#dialogtest start#-->", startStr);
    code= replace(code, "<!--#dialogtest end#-->", endStr);
    for( i= 1 ; i<= nHandleNumb; i++){
        if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
            s= strCut(code, startStr, endStr, 1);
            code= replace2(code, s, "");
        }else{
            break;
        }
    }
    //内转清除
    startStr= "<!--#teststart#-->" ; endStr= "<!--#testend#-->";
    code= replace(code, "<!--#del start#-->", startStr); //与下面一样
    code= replace(code, "<!--#del end#-->", endStr); //与下面一样 多样式
    code= replace(code, "<!--#test start#-->", startStr);
    code= replace(code, "<!--#test end#-->", endStr);

    for( i= 1 ; i<= nHandleNumb; i++){
        if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
            s= strCut(code, startStr, endStr, 1);
            code= replace2(code, s, "");
        }else{
            break;
        }
    }
    //删除注释的span
    code= replace(code, "<sPAn class=\"testspan\">", ""); //测试Span
    code= replace(code, "<sPAn class=\"testhidde\">", ""); //隐藏Span
    code= replace(code, "</sPAn>", "");

    //delTemplateMyNote = Code:Exit Function

    startStr= "<!--#" ; endStr= "#-->";
    for( i= 1 ; i<= nHandleNumb; i++){
        if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
            s= strCut(code, startStr, endStr, 1);
            code= replace2(code, s, "");
        }else{
            break;
        }
    }


    return code;
}

//处理替换参数值 20160114
string handleReplaceValueParam(string content, string paramName, string replaceStr){
    if( inStr(content, "[$" + paramName)== 0 ){
        paramName= lCase(paramName);
    }
    return replaceValueParam(content, paramName, replaceStr);
}

//替换参数值 2014  12 01
string replaceValueParam(string content, string paramName, string replaceStr){
    string startStr=""; string endStr=""; string labelStr=""; string tempLabelStr=""; string sLen=""; string sNTimeFormat=""; string delHtmlYes=""; string funStr=""; string trimYes=""; string sIsEscape=""; string s=""; int i=0;
    string ifStr=""; //判断字符
    string elseIfStr=""; //第二判断字符
    string valueStr=""; //显示字符
    string elseStr=""; //否则字符
    string elseIfValue=""; string elseValue=""; //第二判断值
    string instrStr=""; string instr2Str=""; //查找字符
    string tempReplaceStr=""; //暂存
    //ReplaceStr = ReplaceStr & "这里面放上内容在这时碳呀。"
    //ReplaceStr = CStr(ReplaceStr)            '转成字符类型
    if( isNul(replaceStr)== true ){ replaceStr= "" ;}
    tempReplaceStr= replaceStr;

    //最多处理99个  20160225
    for( i= 1 ; i<= 999; i++){
        replaceStr= tempReplaceStr; //恢复
        startStr= "[$" + paramName ; endStr= "$]";
        //字段名称严格判断 20160226
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 &&(inStr(content, startStr + " ") > 0 || inStr(content, startStr + endStr) > 0) ){
            //获得对应字段加强版20151231
            if( inStr(content, startStr + endStr) > 0 ){
                labelStr= startStr + endStr;
            }else if( inStr(content, startStr + " ") > 0 ){
                labelStr= strCut(content, startStr + " ", endStr, 1);
            }else{
                labelStr= strCut(content, startStr, endStr, 1);
            }

            tempLabelStr= labelStr;
            labelStr= handleInModule(labelStr, "start");
            //删除Html
            delHtmlYes= rParam(labelStr, "delHtml"); //是否删除Html
            if( delHtmlYes== "true" ){ replaceStr= replace(delHtml(replaceStr), "<", "&lt;") ;}//HTML处理
            //删除两边空格
            trimYes= rParam(labelStr, "trim"); //是否删除两边空格
            if( trimYes== "true" ){ replaceStr= trimVbCrlf(replaceStr) ;}

            //截取字符处理
            sLen= rParam(labelStr, "len"); //字符长度值
            sLen= handleNumber(sLen);
            //If sLen<>"" Then ReplaceStr = CutStr(ReplaceStr,sLen,"null")' Left(ReplaceStr,sLen)
            if( sLen != "" ){

                replaceStr= cutStr(replaceStr, cInt(sLen), "..."); //Left(ReplaceStr,nLen)

            }
            //时间处理
            sNTimeFormat= rParam(labelStr, "format_time"); //时间处理值
            if( sNTimeFormat != "" ){
                replaceStr= format_Time(replaceStr, cInt(sNTimeFormat));
            }

            //获得栏目名称
            s= rParam(labelStr, "getcolumnname");
            if( s != "" ){
                if( s== "@ME" ){
                    s= replaceStr;
                }
                replaceStr= getColumnName(s);
            }
            //获得栏目URL
            s= rParam(labelStr, "getcolumnurl");
            if( s != "" ){
                if( s== "@ME" ){
                    s= replaceStr;
                }
                replaceStr= getColumnUrl(s, "id");
            }
            //是否为密码类型
            s= rParam(labelStr, "password");
            if( s != "" ){
                if( s != "" ){
                    replaceStr= s;
                }
            }

            ifStr= rParam(labelStr, "if");
            elseIfStr= rParam(labelStr, "elseif");
            valueStr= rParam(labelStr, "value");

            elseIfValue= rParam(labelStr, "elseifvalue");
            elseValue= rParam(labelStr, "elsevalue");
            instrStr= rParam(labelStr, "instr");
            instr2Str= rParam(labelStr, "instr2");

            //call echo("ifStr",ifStr)
            //call echo("valueStr",valueStr)
            //call echo("elseStr",elseStr)
            //call echo("elseIfStr",elseIfStr)
            //call echo("replaceStr",replaceStr)
            if( ifStr != "" || instrStr != "" ){
                if((ifStr== cStr(replaceStr) && ifStr != "") ){
                    replaceStr= valueStr;
                }else if( elseIfStr== cStr(replaceStr) && elseIfStr != "" ){
                    replaceStr= valueStr;
                    if( elseIfValue != "" ){
                        replaceStr= elseIfValue;
                    }
                }else if( inStr(cStr(replaceStr), instrStr) > 0 && instrStr != "" ){
                    replaceStr= valueStr;
                }else if( inStr(cStr(replaceStr), instr2Str) > 0 && instr2Str != "" ){

                    replaceStr= valueStr;
                    if( elseIfValue != "" ){
                        replaceStr= elseIfValue;
                    }
                }else{
                    if( elseValue != "@ME" ){
                        replaceStr= elseValue;
                    }
                }
            }

            //函数处理20151231    [$title  function='left(@ME,40)'$]
            funStr= rParam(labelStr, "function"); //函数
            if( funStr != "" ){
                funStr= replace(funStr, "@ME", replaceStr);
                replaceStr= handleContentCode(funStr, "");
            }

            //默认值
            s= rParam(labelStr, "default");
            if( s != "" && s != "@ME" ){
                if( replaceStr== "" ){
                    replaceStr= s;
                }
            }
            //escape转码
            sIsEscape= lCase(rParam(labelStr, "escape"));
            if( sIsEscape== "1" || sIsEscape== "true" ){
                replaceStr= escape(replaceStr);
            }

            //文本颜色
            s= rParam(labelStr, "fontcolor"); //函数
            if( s != "" ){
                replaceStr= "<font color=\"" + s + "\">" + replaceStr + "</font>";
            }




            //call echo(tempLabelStr,replaceStr)
            content= replace(content, tempLabelStr, replaceStr);
        }else{
            break;
        }
    }
    return content;
}


//显示编辑器20160115
string displayEditor(string action){
    string c="";
    c= c + "<script type=\"text/javascript\" src=\"\\Jquery\\syntaxhighlighter\\scripts/shCore.js\"></" + "script> " + vbCrlf();
    c= c + "<script type=\"text/javascript\" src=\"\\Jquery\\syntaxhighlighter\\scripts/shBrushJScript.js\"></" + "script>" + vbCrlf();
    c= c + "<script type=\"text/javascript\" src=\"\\Jquery\\syntaxhighlighter\\scripts/shBrushPhp.js\"></" + "script> " + vbCrlf();
    c= c + "<script type=\"text/javascript\" src=\"\\Jquery\\syntaxhighlighter\\scripts/shBrushVb.js\"></" + "script> " + vbCrlf();
    c= c + "<link type=\"text/css\" rel=\"stylesheet\" href=\"\\Jquery\\syntaxhighlighter\\styles/shCore.css\"/>" + vbCrlf();
    c= c + "<link type=\"text/css\" rel=\"stylesheet\" href=\"\\Jquery\\syntaxhighlighter\\styles/shThemeDefault.css\"/>" + vbCrlf();
    c= c + "<script type=\"text/javascript\">" + vbCrlf();
    c= c + "    SyntaxHighlighter.config.clipboardSwf = '\\Jquery\\syntaxhighlighter\\scripts/clipboard.swf';" + vbCrlf();
    c= c + "    SyntaxHighlighter.all();" + vbCrlf();
    c= c + "</" + "script>" + vbCrlf();

    return c;
}
//处理网站url20160202
string handleWebUrl(string url){
    if( cStr(Request["gl"]) != "" ){
        url= getUrlAddToParam(url, "&gl=" + cStr(Request["gl"]), "replace");
    }
    if( cStr(Request["templatedir"]) != "" ){
        url= getUrlAddToParam(url, "&templatedir=" + cStr(Request["templatedir"]), "replace");
    }
    return url;
}

//
//处理在线修改
//MainContent = HandleDisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainContent,"style='float:right;padding:0 4px;'")
string handleDisplayOnlineEditDialog(string url, string content, string cssStyle, string replaceStr){
    string controlStr=""; string[] splStr; string s=""; bool isAdd;
    if( cStr(Request["gl"])== "edit" ){
        if( inStr(url, "&") > 0 ){
            url= url + "&vbgl=true";
        }
        isAdd= false; //添加默认为假
        controlStr= getControlStr(url) + "\"" + cssStyle;
        if( replaceStr != "" ){
            splStr= aspSplit(replaceStr, "|");
            foreach(var eachs in splStr){
                s=eachs;
                if( s != "" && inStr(content, s) > 0 ){
                    content= replace2(content, s, s + controlStr);
                    isAdd= true;
                    break;
                }
            }
        }
        if( isAdd== false ){
            //第一种
            //C = "<div "& ControlStr &">" & vbCrlf
            //C=C & Content & vbCrlf
            //C = C & "</div>" & vbCrlf
            //Content = C
            //第二种
            content= htmlAddAction(content, controlStr);

            //Content = "<div "& ControlStr &">" & Content & "</div>"
        }
    }
    return content;
}
//获得控制内容
string getControlStr(string url){
    string getControlStr="";
    if( cStr(Request["gl"])== "edit" ){
        getControlStr= " onMouseMove=\"onColor(this,'#FDFAC6','red')\" onMouseOut=\"offColor(this,'','')\" onDblClick=\"window1('" + url + "','信息修改')\" title='双击或右键选在线修改' oncontextmenu=\"CommonMenu(event,this,'')"; //删除网址为空
    }
    return getControlStr;
}

//html加动作(20151103)  call rw(htmlAddAction("  <a href=""javascript:;"">222222</a>", "onclick=""javascript:alert(111);"" "))
string htmlAddAction(string content, string jsAction){
    string s=""; string startStr=""; string endStr=""; bool isHandle; string lableName="";
    s= content;
    s= phpTrim(s);
    startStr= mid(s, 1, inStr(s, " "));
    endStr= ">";
    isHandle= true;

    lableName= aspTrim(lCase(replace(startStr, "<", "")));
    if( inStr(s, startStr)== 0 || inStr(s, endStr)== 0 || inStr("|a|div|span|font|h1|h2|h3|h4|h5|h6|dt|dd|dl|li|ul|table|tr|td|", "|" + lableName + "|")== 0 ){
        isHandle= false;
    }

    if( isHandle== true ){
        content= startStr + jsAction + right(s, len(s) - len(startStr));
    }
    return content;
}


</script>
