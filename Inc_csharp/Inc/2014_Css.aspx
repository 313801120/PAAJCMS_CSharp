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
//Css



//ȥ���ַ���ͷβ�������Ļس��Ϳո�


//php��Trim����


//ȥ���ַ�����ͷ�������Ļس��Ϳո�


//ȥ���ַ���ĩβ�������Ļس��Ϳո�



//--------------- ���� ��ʱ�����ļ��� ------------------
//ȥ���ַ���ͷβ��������Tab�˸�Ϳո�



//ȥ���ַ�����ͷ��������Tab�˸�Ϳո�
string lTrimVbTab(string str){
    int nPos=0; bool isBlankChar;
    nPos= 1;
    isBlankChar= true;
    while( isBlankChar){
        if( mid(str, nPos, 1)== " " ){
            nPos= nPos + 1;
        }else if( mid(str, nPos, 1)== vbTab() ){
            nPos= nPos + 1;
        }else{
            isBlankChar= false;
        }
    }
    return right(str, len(str) - nPos + 1);
}

//ȥ���ַ���ĩβ��������Tab�˸�Ϳո�
string rTrimVBTab(string str){
    int nPos=0; bool isBlankChar;
    nPos= len(str);
    isBlankChar= true;
    while( isBlankChar && nPos >= 2){
        if( mid(str, nPos, 1)== " " ){
            nPos= nPos - 1;
        }else if( mid(str, nPos - 1, 1)== vbTab() ){
            nPos= nPos - 1;
        }else{
            isBlankChar= false;
        }
    }
    return aspRTrim(left(str, nPos));
}


//��Htmlҳ��Css Style <style></style> ��css����
string getHtmlCssStyle( string content){
    return handleHtmlStyleCss(content, "", "");
}
//����html���css �滻·��
string handleHtmlStyleCss(string content, string sType, string imgPath){
    string[] splStr; string s=""; bool isStyle; string styleStartStr=""; string styleEndStr=""; string styleStr=""; string cssStr=""; string sNewCssStr="";
    splStr= aspSplit(content, vbCrlf()); //�ָ���
    isStyle= false; //Css��ʽĬ��Ϊ��
    string handleHtmlStyleCss="";
    //ѭ������
    foreach(var eachs in splStr){
        s=eachs;
        if( isStyle== false ){
            if( inStr(lCase(s), "<style") > 0 ){
                styleStartStr= mid(s, inStr(lCase(s), "<style"),-1);
                styleStartStr= mid(styleStartStr, 1, inStr(styleStartStr, ">"));
                styleEndStr= mid(s, inStr(lCase(s), styleStartStr) + len(styleStartStr),-1);
                //HTML�ж����Css��һ��
                if( inStr(styleEndStr, "</style>") > 0 ){
                    styleStr= mid(styleEndStr, 1, inStr(styleEndStr, "</style>") - 1);
                    cssStr= cssStr + styleStr + vbCrlf();
                }else{
                    cssStr= cssStr + styleEndStr + vbCrlf();
                    isStyle= true; //�ռ�CssStyle��ʽ��ʼ
                }
                //Call Echo("StyleStartStr",ShowHtml(StyleStartStr))
                //Call Echo("StyleEndStr",ShowHtml(StyleEndStr))
                //Call Echo("StyleStr",ShowHtml(StyleStr))
                //Call Echo("CssStr",ShowHtml(CssStr))
                //Call RwEnd("")
            }
        }else if( isStyle== true ){
            if( inStr(lCase(s), "</style>") > 0 ){
                styleStr= mid(s, 1, inStr(lCase(s), "</style>") - 1);
                cssStr= cssStr + styleStr + vbCrlf();

                if( sType== "�滻·��" ){
                    sNewCssStr= replaceCssImgPath(cssStr, imgPath);
                    content= replace(content, phpTrim(cssStr), sNewCssStr);
                }

                isStyle= false; //�ռ�CssStyle��ʽ����
            }else{
                cssStr= cssStr + s + vbCrlf();
            }
        }
    }
    if( sType== "�滻·��" ){
        handleHtmlStyleCss= content;
    }else{
        handleHtmlStyleCss= cssStr;
    }
    return handleHtmlStyleCss;
}

//�滻css������ͼƬ·�� 20160718home
string replaceCssImgPath( string cssStr, string imgPath){
    string content=""; string[] splStr; string s=""; string c=""; string fileName=""; string toImgPath=""; string lCaseS="";
    content= getArray(cssStr, "\\(", "\\)", false, false);
    splStr= aspSplit(content, "$Array$");
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            lCaseS= lCase(s);
            fileName= aspTrim(getFileAttr(s, "2"));
            //����ͼƬ �����жϷ������д��Ľ����Ľ�Ҳ�ǳɳ���һ����20160719
            if( inStr(lCaseS, "data:image")== 0 &&(inStr(lCaseS, ".jpg") > 0 || inStr(lCaseS, ".gif") > 0 || inStr(lCaseS, ".png") > 0 || inStr(lCaseS, ".bmp") > 0) ){
                //�޸���20160719
                if( inStr("\",", left(fileName, 1)) > 0 ){
                    fileName= mid(fileName, 2,-1);
                }
                if( inStr("\",", right(fileName, 1)) > 0 ){
                    fileName= mid(fileName, 1, len(fileName) - 1);
                }
                //call echo(s,imgPath  & " and " & fileName)
                cssStr= replace(cssStr, "(" + s + ")", "(" + imgPath + fileName + ")");
                //fileName="data:image/" & fileName
            }

        }
    }
    return cssStr;
}

//����ɸɾ���Css����  CSS��ʽ��
string handleCleanCss( string content){
    string[] splStr; string s=""; string c=""; bool isAddStr; string customS="";
    content= replace(content, "{", vbCrlf() + "{" + vbCrlf());
    content= replace(content, "}", vbCrlf() + "}" + vbCrlf());
    content= replace(content, ";", ";" + vbCrlf());

    splStr= aspSplit(content, vbCrlf());
    isAddStr= false; //׷���ַ�Ĭ��Ϊ��
    foreach(var eachs in splStr){
        s=eachs;
        s= trimVbCrlf(s);
        customS= ""; //�Զ���Sֵ
        if( s != "" ){
            if( inStr(s, "{") > 0 && inStr(s, "}")== 0 ){
                isAddStr= true;
                customS= s;
            }else if( inStr(s, "}") > 0 ){
                isAddStr= false;
            }
            if( left(s, 1) != "{" ){ c= c + vbCrlf() ;}
            if( isAddStr== true ){ s= "    " + s ;}
            if( customS != "" ){ s= customS ;}//�Զ���ֵ��Ϊ�������Զ�������
            c= c + s;

        }
    }
    c= trimVbCrlf(c);
    c= replace(c, "    ;" + vbCrlf(), ""); //������ڵķֺ�
    c= replace(c, ";" + vbCrlf() + "}", vbCrlf() + "}"); //���һ��������Ҫ�ֺ�
    return c;
}



//�Ƴ������ж����
string removeExcessRow(string content){
    string[] splStr; string s=""; string c=""; string tempS="";
    splStr= aspSplit(content, vbCrlf()); //�ָ���
    foreach(var eachs in splStr){
        s=eachs;
        tempS= replace(replace(s, " ", ""), vbTab(), "");
        if( tempS != "" ){
            c= c + s + vbCrlf();
        }
    }
    if( c != "" ){ c= left(c, len(c) - 2); }
    return c;
}


//2014 11 30
//��Css��׷����ʽ  a=CssAddToStyle(GetFText("1.html")," .test {color:#FF0f000; font-size:10px; float:left}")
string cssAddToStyle(string content, string addToStyle){
    string styleName=""; string yunStyleStr=""; string replaceStyleStr=""; string c="";
    if( inStr(addToStyle, "{") > 0 ){
        styleName= aspTrim(mid(addToStyle, 1, inStr(addToStyle, "{") - 1));
    }
    yunStyleStr= findCssStyle(content, styleName);
    replaceStyleStr= cssStyleAddToParam(yunStyleStr, addToStyle); //Css��ʽ�ۼӲ���
    content= replace(content, yunStyleStr, replaceStyleStr);
    //C = C & "<hr>Content=" & Content
    string cssAddToStyle= content;
    //CssAddToStyle = YunStyleStr
    //CssAddToStyle = "StyleName=" & StyleName & "<hr>YunStyleStr=" & YunStyleStr & "<hr>ReplaceStyleStr=" & ReplaceStyleStr
    return cssAddToStyle;
}

//���Css�������Ƿ���ָ����ʽ
bool checkCssStyle(string content, string styleStr){
    string styleName="";
    bool checkCssStyle= true;
    if( inStr(styleStr, "{") > 0 ){
        styleName= aspTrim(mid(styleStr, 1, inStr(styleStr, "{") - 1));
    }
    if( styleName== "" ){
        checkCssStyle= false;
    }else if( findCssStyle(content, styleName)== "" ){
        checkCssStyle= false;
    }
    return checkCssStyle;
}


//Css��ʽ�ۼӲ���
string cssStyleAddToParam( string cssStyleStr, string cssStyleStrTwo){
    string[] splStr; string cssStr=""; string s=""; string paramList=""; string paramName=""; string cssStyleName="";
    cssStyleName= mid(cssStyleStr, 1, inStr(cssStyleStr, "{"));
    if( inStr(cssStyleStr, "{") > 0 ){
        cssStyleStr= mid(cssStyleStr, inStr(cssStyleStr, "{") + 1,-1);
    }
    if( inStr(cssStyleStr, "}") > 0 ){
        cssStyleStr= mid(cssStyleStr, 1, inStr(cssStyleStr, "}") - 1);
    }
    if( inStr(cssStyleStrTwo, "{") > 0 ){
        cssStyleStrTwo= mid(cssStyleStrTwo, inStr(cssStyleStrTwo, "{") + 1,-1);
    }
    if( inStr(cssStyleStrTwo, "}") > 0 ){
        cssStyleStrTwo= mid(cssStyleStrTwo, 1, inStr(cssStyleStrTwo, "}") - 1);
    }
    splStr= aspSplit(replace(cssStyleStr + ";" + cssStyleStrTwo, vbCrlf(), ""), ";");
    foreach(var eachs in splStr){
        s=eachs;
        s= aspTrim(s);
        if( inStr(s, ":") > 0 && s != "" ){
            paramName= aspTrim(mid(s, 1, inStr(s, ":") - 1));
            if( inStr("|" + paramList + "|", "|" + paramName + "|")== 0 ){
                paramList= paramList + paramName + "|";
                //Call Echo("ParamName",ParamName)
                cssStr= cssStr + "    " + s + ";" + vbCrlf();
            }
        }
    }
    if( cssStyleName != "" ){
        cssStr= cssStyleName + vbCrlf() + cssStr + "}";
    }
    string cssStyleAddToParam= cssStr;
    //Call Echo(CssStyleStr,CssStyleStrTwo)
    return cssStyleAddToParam;
}

//����Css�����ҵ���ӦCss��
string findCssStyle( string content, string styleName){
    string[] splStr; string s=""; string tempS=""; string findStyleName="";
    //CAll Echo("StyleName",StyleName)
    //CAll Echo("Content",Content)
    string findCssStyle="";
    styleName= aspTrim(styleName);
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, styleName) > 0 ){
            findStyleName= aspTrim(s);
            if( inStr(findStyleName, "{") > 0 ){
                findStyleName= aspTrim(mid(findStyleName, 1, inStr(findStyleName, "{") - 1));
            }
            if( findStyleName== styleName ){
                //Call Eerr( FindStyleName , StyleName)
                if( inStr(s, "}") > 0 ){
                    findCssStyle= mid(s, 1, inStr(s, "}") + 1);
                    //Call EErr(s,FindCssStyle)
                    return findCssStyle;
                }else{
                    tempS= mid(content, inStr(content, s + vbCrlf()) + 1,-1);
                    tempS= mid(tempS, 1, inStr(tempS, "}") + 1);
                    findCssStyle= tempS;
                    return findCssStyle;
                }
                //Call Eerr("temps",Temps)
            }
            //Call Echo(FindStyleName,StyleName)
        }
    }
    return findCssStyle;
}

//================================================
//������վǰ��Ҫ�õ�
//================================================
//�����ȡ����Css
string handleCutCssCode(string dirPath, string cssStr){
    string content=""; string startStr=""; string endStr=""; string[] splStr; string imageFile=""; string fileName=""; string listStr="";
    startStr= "url\\(" ; endStr= "\\)";
    content= getArray(cssStr, startStr, endStr, false, false);
    splStr= aspSplit(content, "$Array$");
    foreach(var eachimageFile in splStr){
        imageFile=eachimageFile;
        if( imageFile != "" && inStr(imageFile, ".") > 0 && inStr(vbCrlf() + listStr + vbCrlf(), vbCrlf() + imageFile + vbCrlf())== 0 ){//���ظ�ʹ�õ�ͼƬ����
            listStr= listStr + imageFile + vbCrlf();
            fileName= replace(replace(replace(imageFile, "\"", ""), "'", ""), "\\", "/");
            if( inStr(fileName, "/") > 0 ){
                fileName= mid(fileName, inStrRev(fileName, "/") + 1,-1);
            }
            cssStr= replace(cssStr, imageFile, dirPath + fileName);
        }
    }
    return cssStr;
}

//�����ȡ����HtmlDiv
string handleCutDivCode(string dirPath, string divStr){
    string content=""; string startStr=""; string endStr=""; string[] splStr; string imageFile=""; string toImageFile=""; string fileName=""; bool isHandle;
    startStr= "url\\(" ; endStr= "\\)";
    content= getArray(divStr, startStr, endStr, false, false);
    splStr= aspSplit(content, "$Array$");
    foreach(var eachimageFile in splStr){
        imageFile=eachimageFile;

        if( imageFile != "" && inStr(imageFile, ".") > 0 && inStr(imageFile, "{$#")== 0 ){
            //�ж��Ƿ������� 20150202
            if( getWebSite(imageFile)== "" ){
                fileName= replace(replace(replace(imageFile, "\"", ""), "'", ""), "\\", "/");
                if( inStr(fileName, "/") > 0 ){
                    fileName= mid(fileName, inStrRev(fileName, "/") + 1,-1);
                }
                divStr= replace(divStr, imageFile, dirPath + imageFile);
            }
        }
    }
    //ͼƬ����
    //Content = GetIMG(DivStr) & vbCrlf & GetHtmlBackGroundImgList(DivStr)        '�ټӸ�Html����ͼƬ
    content= getImgJsUrl(divStr, "���ظ�") + vbCrlf() + getHtmlBackGroundImgList(divStr); //�ټӸ�Html����ͼƬ  ��ǿ��20150126
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachimageFile in splStr){
        imageFile=eachimageFile;
        if( imageFile != "" ){ //�����ӵ�ַ��ǰΪHTTP:ʱ�򲻴���20150313
            isHandle= false;

            if( left(imageFile, 1)== "\\" ){
                //�ȴ���20150817
            }else if( inStr(imageFile, ".") > 0 && left(imageFile, 5) != "HTTP:" && inStr(imageFile, "{$#")== 0 ){
                isHandle= true;
            }
            if( isHandle== true ){
                toImageFile= dirPath + removeFileDir(imageFile); //�Ƴ��ļ�·��Ŀ¼
                //html��ͼƬ·���滻
                divStr= replace(divStr, "\"" + imageFile + "\"", "\"" + toImageFile + "\"");
                divStr= replace(divStr, "'" + imageFile + "'", "\"" + toImageFile + "\"");
                divStr= replace(divStr, "=" + imageFile + " ", "\"" + toImageFile + "\"");
                divStr= replace(divStr, "=" + imageFile + ">", "\"" + toImageFile + "\"");
            }
        }
    }
    return divStr;
}

//���HTMl�ﱳ��ͼƬ 20150116  �磺 <td width="980" height="169" background="kslx3bg.jpg">
string getHtmlBackGroundImgList( string content){
    content= getArray(content, " background=\"", "\"", false, false);
    content= replace(content, "$Array$", vbCrlf());
    return content;
}


//����html��link css���� Content = getHandleWebHtmlLink("/aa/bb/",Content)  �ⲿ����
string getHandleWebHtmlLink(string rootPath, string content){
    string startStr=""; string endStr=""; string imgList=""; string[] splStr; string c=""; string cssUrl=""; string sNewCssUrl=""; string cssStr="";
    startStr= "<link ";
    cssStr= "";
    endStr= ">";
    imgList= getArray(content, startStr, endStr, false, false);
    //Call RwEnd(ImgList)
    splStr= aspSplit(imgList, "$Array$");
    foreach(var eachcssUrl in splStr){
        cssUrl=eachcssUrl;
        if( cssUrl != "" && inStr(lCase(cssUrl), "stylesheet") > 0 ){
            //���Css��ǿ�棬����20141125
            cssUrl= lCase(replace(replace(replace(cssUrl, "\"", ""), "'", ""), ">", " ")) + " ";
            startStr= "href=" ; endStr= " ";
            if( inStr(cssUrl, startStr) > 0 && inStr(cssUrl, endStr) > 0 ){
                cssUrl= strCut(cssUrl, startStr, endStr, 2);
            }
            sNewCssUrl= handleHttpUrl(cssUrl);
            if( inStr(sNewCssUrl, "/") > 0 ){
                sNewCssUrl= mid(sNewCssUrl, inStrRev(sNewCssUrl, "/") + 1,-1);
            }
            if( lCase(sNewCssUrl) != "common.css" && lCase(sNewCssUrl) != "public.css" ){
                sNewCssUrl= rootPath + sNewCssUrl;
                cssStr= cssStr + "<link href=\"" + sNewCssUrl + "\" rel=\"stylesheet\" type=\"text/css\" />" + vbCrlf();
            }
        }
    }
    if( cssStr != "" ){ cssStr= left(cssStr, len(cssStr) - 2); }
    return cssStr;
}


//���css���ӵ�ַ�б�(20150824)  ��<link>
string getCssListUrlList(string content){
    string startStr=""; string endStr=""; string imgList=""; string[] splStr; string c=""; string cssUrl=""; string cssStr=""; string urlList="";
    startStr= "<link ";
    cssStr= "";
    endStr= ">";
    imgList= getArray(content, startStr, endStr, false, false);
    //Call RwEnd(ImgList)
    splStr= aspSplit(imgList, "$Array$");
    foreach(var eachcssUrl in splStr){
        cssUrl=eachcssUrl;
        if( cssUrl != "" && inStr(lCase(cssUrl), "stylesheet") > 0 ){
            //���Css��ǿ�棬����20141125
            cssUrl= lCase(replace(replace(replace(cssUrl, "\"", ""), "'", ""), ">", " ")) + " ";
            startStr= "href=" ; endStr= " ";
            if( inStr(cssUrl, startStr) > 0 && inStr(cssUrl, endStr) > 0 ){
                cssUrl= strCut(cssUrl, startStr, endStr, 2);
            }
            if( inStr(vbCrlf() + urlList + vbCrlf(), vbCrlf() + cssUrl + vbCrlf())== 0 ){
                if( urlList != "" ){ urlList= urlList + vbCrlf() ;}
                urlList= urlList + cssUrl + vbCrlf();
            }
        }
    }
    return urlList;
}
//call rw(handleHtmlStyle(getftext("1.html")))
//ѹ��html���style��ʽ (20151008)
string handleHtmlStyle(string content){
    string serchS=""; string replaceS=""; int nLength=0;
    serchS= content;
    nLength= inStr(lCase(serchS), "</style>") + 7;
    serchS= mid(serchS, 1, nLength);

    nLength= inStrRev(lCase(serchS), "<style");
    if( nLength > 0 ){
        serchS= mid(serchS, nLength,-1);
    }
    replaceS= serchS;
    replaceS= cssCompression(replaceS, 0) + vbCrlf(); //��ʽ��CSS
    replaceS= removeBlankLines(replaceS);

    content= replace(content, serchS, replaceS);
    return content;
}



//��Css�ļ����ݲ�����(20150824) �������¼��ǩ   ������
//��call rwend(handleReadCssContent("E:\E��\WEB��վ\��ǰ��վ\DataDir\VBģ��\������\Template\ģ�鹦���б�\Bվҳ�����\home\home.css","aa",true))
string handleReadCssContent(string cssFilePath, string labelName, bool isHandleCss){
    string c=""; string startStr=""; string endStr="";
    c= getFText(cssFilePath);
    //��ȡCSS
    startStr= "/*CssCodeStart*/";
    endStr= "/*CssCodeEnd*/";
    if( inStr(c, startStr) > 0 && inStr(c, endStr) > 0 ){
        c= strCut(c, startStr, endStr, 2);
    }
    //����CSS
    if( isHandleCss== true ){
        c= cssCompression(c, 0);
    }
    if( labelName != "" ){
        c= "/*" + labelName + " start*/" + c + "/*" + labelName + " end*/";
    }
    return c;
}

//����Css��ʽ��PX��%
string handleCssPX( string sValue){
    sValue= lCase(aspTrim(sValue));
    if( right(sValue, 1) != "%" && right(sValue, 2) != "px" ){
        sValue= sValue + "px";
    }
    return sValue;
}

</script>
