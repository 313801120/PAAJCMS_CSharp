<%
/************************************************************
���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
��Ȩ��Դ������ѹ�����������;����ʹ�á� 
������2016-08-04
��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
*                                    Powered by PAAJCMS 
************************************************************/
%>
<script runat="server" language="c#">
//��վ���� 20160223



//����ģ���滻����


//ȥ��ģ���ﲻ��Ҫ��ʾ���� ɾ��ģ�����ҵ�ע�ʹ���
string delTemplateMyNote(string code){
    string startStr=""; string endStr=""; int i=0; string s=""; int nHandleNumb=0; string[] splStr; string block=""; string id="";
    string content=""; string dragSortCssStr=""; string dragSortStart=""; string dragSortEnd=""; string dragSortValue=""; string c="";
    string lableName=""; string lableStartStr=""; string lableEndStr="";
    nHandleNumb= 99; //���ﶨ�����Ҫ

    //��ǿ��  �����Ҳ����<!--#aaa start#--><!--#aaa end#-->
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



    //���ReadBlockList�������б�����  �����и�����ĵط����������ݿ��Դ��ⲿ�������ݣ�����Ժ���
    //Call Eerr("ReadBlockList",ReadBlockList)
    //д��20141118
    //splStr = Split(ReadBlockList, vbCrLf)                 '�������֣�������
    //�޸���20151230
    for( i= 1 ; i<= nHandleNumb; i++){
        startStr= "<R#��������" ; endStr= " start#>";
        block= strCut(code, startStr, endStr, 2);
        if( block != "" ){
            startStr= "<R#��������" + block + " start#>" ; endStr= "<R#��������" + block + " end#>";
            if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
                s= strCut(code, startStr, endStr, 1);
                code= replace(code, s, ""); //�Ƴ�
            }
        }else{
            break;
        }
    }

    //ɾ����ҳ����20160309
    startStr= "<!--#list start#-->";
    endStr= "<!--#list end#-->";
    if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
        s= strCut(code, startStr, endStr, 2);
        code= replace(code, s, "");
    }

    if( cStr(Request["gl"])== "yun" ){
        content= getFText("/Jquery/dragsort/Config.html");
        content= getFText("/Jquery/dragsort/ģ����ק.html");
        //Css��ʽ
        startStr= "<style>";
        endStr= "</style>";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            dragSortCssStr= strCut(content, startStr, endStr, 1);
        }
        //��ʼ����
        startStr= "<!--#top start#-->";
        endStr= "<!--#top end#-->";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            dragSortStart= strCut(content, startStr, endStr, 2);
        }
        //��������
        startStr= "<!--#foot start#-->";
        endStr= "<!--#foot end#-->";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            dragSortEnd= strCut(content, startStr, endStr, 2);
        }
        //��ʾ������
        startStr= "<!--#value start#-->";
        endStr= "<!--#value end#-->";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            dragSortValue= strCut(content, startStr, endStr, 2);
        }



        //���ƴ���
        startStr= "<dIv datid='";
        endStr= "</dIv>";
        content= getArray(code, startStr, endStr, false, false);
        splStr= aspSplit(content, "$Array$");
        foreach(var eachs in splStr){
            s=eachs;
            startStr= "��DatId��'";
            id= mid(s, 1, inStr(s, startStr) - 1);
            s= mid(s, inStr(s, startStr) + len(startStr),-1);
            //C=C & "<li><div title='"& Id &"'>" & vbcrlf & "<div " & S & "</div>"& vbcrlf &"<div class='clear'></div></div><div class='clear'></div></li>"
            s= "<div" + s + "</div>";
            //Call Die(S)
            c= c + replace(replace(dragSortValue, "{$value$}", s), "{$id$", id);
        }
        c= replace(c, "�����С�", vbCrlf());
        c= dragSortStart + c + dragSortEnd;
        code= mid(code, 1, inStr(code, "<body>") - 1);
        code= replace(code, "</head>", dragSortCssStr + "</head></body>" + c + "</body></html>");
    }

    //ɾ��VB������ɵ���������
    startStr= "<dIv datid='" ; endStr= "��DatId��'";
    for( i= 1 ; i<= nHandleNumb; i++){
        if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
            id= strCut(code, startStr, endStr, 2);
            code= replace2(code, startStr + id + endStr, "<div ");
        }else{
            break;
        }
    }
    code= replace(code, "</dIv>", "</div>"); //�滻���������div

    //����Χ���
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
    //��ת���
    startStr= "<!--#teststart#-->" ; endStr= "<!--#testend#-->";
    code= replace(code, "<!--#del start#-->", startStr); //������һ��
    code= replace(code, "<!--#del end#-->", endStr); //������һ�� ����ʽ
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
    //ɾ��ע�͵�span
    code= replace(code, "<sPAn class=\"testspan\">", ""); //����Span
    code= replace(code, "<sPAn class=\"testhidde\">", ""); //����Span
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

//�����滻����ֵ 20160114
string handleReplaceValueParam(string content, string paramName, string replaceStr){
    if( inStr(content, "[$" + paramName)== 0 ){
        paramName= lCase(paramName);
    }
    return replaceValueParam(content, paramName, replaceStr);
}

//�滻����ֵ 2014  12 01
string replaceValueParam(string content, string paramName, string replaceStr){
    string startStr=""; string endStr=""; string labelStr=""; string tempLabelStr=""; string sLen=""; string sNTimeFormat=""; string delHtmlYes=""; string funStr=""; string trimYes=""; string sIsEscape=""; string s=""; int i=0;
    string ifStr=""; //�ж��ַ�
    string elseIfStr=""; //�ڶ��ж��ַ�
    string valueStr=""; //��ʾ�ַ�
    string elseStr=""; //�����ַ�
    string elseIfValue=""; string elseValue=""; //�ڶ��ж�ֵ
    string instrStr=""; string instr2Str=""; //�����ַ�
    string tempReplaceStr=""; //�ݴ�
    //ReplaceStr = ReplaceStr & "�����������������ʱ̼ѽ��"
    //ReplaceStr = CStr(ReplaceStr)            'ת���ַ�����
    if( isNul(replaceStr)== true ){ replaceStr= "" ;}
    tempReplaceStr= replaceStr;

    //��ദ��99��  20160225
    for( i= 1 ; i<= 999; i++){
        replaceStr= tempReplaceStr; //�ָ�
        startStr= "[$" + paramName ; endStr= "$]";
        //�ֶ������ϸ��ж� 20160226
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 &&(inStr(content, startStr + " ") > 0 || inStr(content, startStr + endStr) > 0) ){
            //��ö�Ӧ�ֶμ�ǿ��20151231
            if( inStr(content, startStr + endStr) > 0 ){
                labelStr= startStr + endStr;
            }else if( inStr(content, startStr + " ") > 0 ){
                labelStr= strCut(content, startStr + " ", endStr, 1);
            }else{
                labelStr= strCut(content, startStr, endStr, 1);
            }

            tempLabelStr= labelStr;
            labelStr= handleInModule(labelStr, "start");
            //ɾ��Html
            delHtmlYes= rParam(labelStr, "delHtml"); //�Ƿ�ɾ��Html
            if( delHtmlYes== "true" ){ replaceStr= replace(delHtml(replaceStr), "<", "&lt;") ;}//HTML����
            //ɾ�����߿ո�
            trimYes= rParam(labelStr, "trim"); //�Ƿ�ɾ�����߿ո�
            if( trimYes== "true" ){ replaceStr= trimVbCrlf(replaceStr) ;}

            //��ȡ�ַ�����
            sLen= rParam(labelStr, "len"); //�ַ�����ֵ
            sLen= handleNumber(sLen);
            //If sLen<>"" Then ReplaceStr = CutStr(ReplaceStr,sLen,"null")' Left(ReplaceStr,sLen)
            if( sLen != "" ){

                replaceStr= cutStr(replaceStr, cInt(sLen), "..."); //Left(ReplaceStr,nLen)

            }
            //ʱ�䴦��
            sNTimeFormat= rParam(labelStr, "format_time"); //ʱ�䴦��ֵ
            if( sNTimeFormat != "" ){
                replaceStr= format_Time(replaceStr, cInt(sNTimeFormat));
            }

            //�����Ŀ����
            s= rParam(labelStr, "getcolumnname");
            if( s != "" ){
                if( s== "@ME" ){
                    s= replaceStr;
                }
                replaceStr= getColumnName(s);
            }
            //�����ĿURL
            s= rParam(labelStr, "getcolumnurl");
            if( s != "" ){
                if( s== "@ME" ){
                    s= replaceStr;
                }
                replaceStr= getColumnUrl(s, "id");
            }
            //�Ƿ�Ϊ��������
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

            //��������20151231    [$title  function='left(@ME,40)'$]
            funStr= rParam(labelStr, "function"); //����
            if( funStr != "" ){
                funStr= replace(funStr, "@ME", replaceStr);
                replaceStr= handleContentCode(funStr, "");
            }

            //Ĭ��ֵ
            s= rParam(labelStr, "default");
            if( s != "" && s != "@ME" ){
                if( replaceStr== "" ){
                    replaceStr= s;
                }
            }
            //escapeת��
            sIsEscape= lCase(rParam(labelStr, "escape"));
            if( sIsEscape== "1" || sIsEscape== "true" ){
                replaceStr= escape(replaceStr);
            }

            //�ı���ɫ
            s= rParam(labelStr, "fontcolor"); //����
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


//��ʾ�༭��20160115
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
//������վurl20160202
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
//���������޸�
//MainContent = HandleDisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainContent,"style='float:right;padding:0 4px;'")
string handleDisplayOnlineEditDialog(string url, string content, string cssStyle, string replaceStr){
    string controlStr=""; string[] splStr; string s=""; bool isAdd;
    if( cStr(Request["gl"])== "edit" ){
        if( inStr(url, "&") > 0 ){
            url= url + "&vbgl=true";
        }
        isAdd= false; //���Ĭ��Ϊ��
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
            //��һ��
            //C = "<div "& ControlStr &">" & vbCrlf
            //C=C & Content & vbCrlf
            //C = C & "</div>" & vbCrlf
            //Content = C
            //�ڶ���
            content= htmlAddAction(content, controlStr);

            //Content = "<div "& ControlStr &">" & Content & "</div>"
        }
    }
    return content;
}
//��ÿ�������
string getControlStr(string url){
    string getControlStr="";
    if( cStr(Request["gl"])== "edit" ){
        getControlStr= " onMouseMove=\"onColor(this,'#FDFAC6','red')\" onMouseOut=\"offColor(this,'','')\" onDblClick=\"window1('" + url + "','��Ϣ�޸�')\" title='˫�����Ҽ�ѡ�����޸�' oncontextmenu=\"CommonMenu(event,this,'')"; //ɾ����ַΪ��
    }
    return getControlStr;
}

//html�Ӷ���(20151103)  call rw(htmlAddAction("  <a href=""javascript:;"">222222</a>", "onclick=""javascript:alert(111);"" "))
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
