<%
/************************************************************
���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
��Ȩ��Դ������ѹ�����������;����ʹ�á� 
������2016-08-05
��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
*                                    Powered by PAAJCMS 
************************************************************/
%>
<script runat="server" language="c#">
//Html ����HTML���� (2014,1,3)
//��ʾHTML�ṹ        call rw(displayHTmL("<br>aasdfds<br>"))
//�ر���ʾHTML�ṹ   call rwend(unDisplayHtml("&lt;br&gt;aasdfds&lt;br&gt;"))

//��ʾHTML�ṹ
string displayHtml(string str){
    str= replace(str, "<", "&lt;");
    str= replace(str, ">", "&gt;");
    return str;
}
//�ر���ʾHTML�ṹ
string unDisplayHtml(string str){
    str= replace(str, "&lt;", "<");
    str= replace(str, "&gt;", ">");
    return str;
}

//����պ�HTML��ǩ(20150902)  ������ĸ����� �ڶ���
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
                endStr= right(endStr, len(endStr) - len(s) - 2); //����ַ���ȥ��ǰ��ǩ  -2����Ϊ����<>�����ַ�
                //ע��֮ǰ����labelName����
                labelName= mid(s, 1, inStr(s + " ", " ") - 1);
                labelName= lCase(labelName);
                //call eerr("s",s)

                if( inStr(action, "|����A����|") > 0 ){
                    s= handleHtmlAHref(s, labelName, "http://127.0.0.1/TestWeb/Web/", "����A����"); //����ɾ�html��ǩ
                }else if( inStr(action, "|����A���ӵڶ���|") > 0 ){
                    s= handleHtmlAHref(s, labelName, "http://127.0.0.1/debugRemoteWeb.asp?url=", "����A����"); //����ɾ�html��ǩ
                }
                //call echo(s,labelName)   param��embed��Flash�õ�������embed�н�����ǩ��
                if( inStr("|meta|link|embed|param|input|img|br|hr|rect|line|area|script|div|span|a|", "|" + labelName + "|") > 0 ){
                    s= replace(replace(replace(replace(s, " class=\"\"", ""), " alt=\"\"", ""), " title=\"\"", ""), " name=\"\"", ""); //��ʱ��ô��һ�£��Ժ�Ҫ����ϵͳ����
                    s= replace(replace(replace(replace(s, " class=''", ""), " alt=''", ""), " title=''", ""), " name=''", "");

                    //��vb.net����õ� Ҫ��Ȼ���ᱨ����
                    if( labelName== "img" && isImgAddAlt== true ){
                        if( inStr(s, " alt")== 0 ){
                            s= s + " alt=\"\"";
                        }
                        s= aspTrim(s);
                        s= s + " /";
                        //����<script>20160106  ��ʱ������������ȸĽ�
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
                //����javascript script����
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
//����htmlA��ǩ��Href����  ������溯��
string handleHtmlAHref( string content, string labelName, string addToHttpUrl, string action){
    int i=0; string s=""; string c=""; string temp="";
    bool isValue; //�Ƿ�Ϊ����ֵ
    string valueStr=""; //�洢����ֵ
    string yinghaoLabel=""; //����������'"
    string parentName=""; //��������
    string behindStr=""; //����ȫ���ַ�
    string sNotDanYinShuangYinStr=""; //���ǵ����ź�˫�����ַ�
    action= "|" + action + "|";
    content= replace(content + " ", vbTab(), " "); //�˸��滻�ɿո�����һ���ո񣬷������
    content= replace(replace(content, " =", "="), " =", "=");
    isValue= false; //Ĭ������Ϊ�٣���Ϊ���ǻ�ñ�ǩ����
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1); //��õ�ǰһ���ַ�
        behindStr= mid(content, i,-1); //�����ַ�
        if( s== "=" && isValue== false ){ //��������ֵ����Ϊ=��
            isValue= true;
            valueStr= "";
            yinghaoLabel= "";
            if( c != "" && right(c, 1) != " " ){ c= c + " " ;}
            parentName= lCase(temp); //��������תСд
            c= c + parentName + s;
            temp= "";
            //���ֵ��һ���ַ�����Ϊ������������
        }else if( isValue== true && yinghaoLabel== "" ){
            if( s != " " ){
                if( s != "'" && s != "\"" ){
                    sNotDanYinShuangYinStr= s; //���ǵ����ź�˫�����ַ�
                    s= " ";
                }
                yinghaoLabel= s;
                //call echo("yinghaoLabel",yinghaoLabel)
            }
        }else if( isValue== true && yinghaoLabel != "" ){
            //Ϊ���Ž���
            if( yinghaoLabel== s ){
                isValue= false;
                if( labelName== "a" && parentName== "href" && inStr(action, "|����A����|") > 0 ){
                    //����
                    if( inStr(valueStr, "?") > 0 ){
                        valueStr= replace(valueStr, "?", "WenHao") + ".html";
                    }
                    if( inStr("|asp|php|aspx|jsp|", "|" + lCase(mid(valueStr, inStrRev(valueStr, ".") + 1,-1)) + "|") > 0 ){
                        valueStr= valueStr + ".html";
                    }
                    valueStr= addToOrAddHttpUrl(addToHttpUrl, valueStr, "�滻");

                }
                //call echo("labelName",labelName)
                if( yinghaoLabel== " " ){
                    c= c + "\"" + sNotDanYinShuangYinStr + valueStr + "\" "; //׷�� ���ǵ����ź�˫�����ַ�            ��ȫ
                }else{
                    c= c + yinghaoLabel + valueStr + yinghaoLabel; //׷�� ���ǵ����ź�˫�����ַ�
                }
                yinghaoLabel= "";
                sNotDanYinShuangYinStr= ""; //���ǵ����ź�˫�����ַ� ���
            }else{
                valueStr= valueStr + s;
            }
            //Ϊ �ָ�
        }else if( s== " " ){
            //�ݴ����ݲ�Ϊ��
            if( temp != "" ){
                if( left(aspTrim(behindStr) + " ", 1)== "=" ){
                    //����һ���ַ�����=������
                }else{
                    //Ϊ��ǩ
                    if( isValue== false ){
                        temp= lCase(temp) + " "; //��ǩ��������תСд
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
//׷�ӻ��滻��ַ(20150922) �������   addToOrAddHttpUrl("http://127.0.0.1/aa/","http://127.0.0.1/4.asp","�滻") = http://127.0.0.1/aa/4.asp
string addToOrAddHttpUrl(string httpurl, string url, string action){
    string s="";
    action= "|" + action + "|";
    if( inStr(action, "|�滻|") > 0 ){
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

//���HTML��ǩ�� call rwend(getHtmlLableName("<img src><a href=>",0))    ���  img
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
                endStr= right(endStr, len(endStr) - len(s) - 2); //����ַ���ȥ��ǰ��ǩ  -2����Ϊ����<>�����ַ�
                //ע��֮ǰ����labelName����
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

//ɾ��html����� ��ķ��� ɾ������
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
//���html����ָ��ֵ20160520 call echo("webtitle",getHtmlValue(content,"webtitle"))
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
                endStr= right(endStr, len(endStr) - len(s) - 2); //����ַ���ȥ��ǰ��ǩ  -2����Ϊ����<>�����ַ�
                //ע��֮ǰ����labelName����
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

//��ò���ֵ20160520  call rwend(getParamValue("meta name=""keywords"" content=""{$web_keywords$}""","name"))
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
