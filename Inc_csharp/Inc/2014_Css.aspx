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
//Css



//去掉字符串头尾的连续的回车和空格


//php里Trim方法


//去掉字符串开头的连续的回车和空格


//去掉字符串末尾的连续的回车和空格



//--------------- 有用 暂时用这文件里 ------------------
//去掉字符串头尾的连续的Tab退格和空格



//去掉字符串开头的连续的Tab退格和空格
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

//去掉字符串末尾的连续的Tab退格和空格
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


//找Html页中Css Style <style></style> 里css内容
string getHtmlCssStyle( string content){
    return handleHtmlStyleCss(content, "", "");
}
//处理html里的css 替换路径
string handleHtmlStyleCss(string content, string sType, string imgPath){
    string[] splStr; string s=""; bool isStyle; string styleStartStr=""; string styleEndStr=""; string styleStr=""; string cssStr=""; string sNewCssStr="";
    splStr= aspSplit(content, vbCrlf()); //分割行
    isStyle= false; //Css样式默认为假
    string handleHtmlStyleCss="";
    //循环分行
    foreach(var eachs in splStr){
        s=eachs;
        if( isStyle== false ){
            if( inStr(lCase(s), "<style") > 0 ){
                styleStartStr= mid(s, inStr(lCase(s), "<style"),-1);
                styleStartStr= mid(styleStartStr, 1, inStr(styleStartStr, ">"));
                styleEndStr= mid(s, inStr(lCase(s), styleStartStr) + len(styleStartStr),-1);
                //HTML中定义的Css在一行
                if( inStr(styleEndStr, "</style>") > 0 ){
                    styleStr= mid(styleEndStr, 1, inStr(styleEndStr, "</style>") - 1);
                    cssStr= cssStr + styleStr + vbCrlf();
                }else{
                    cssStr= cssStr + styleEndStr + vbCrlf();
                    isStyle= true; //收集CssStyle样式开始
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

                if( sType== "替换路径" ){
                    sNewCssStr= replaceCssImgPath(cssStr, imgPath);
                    content= replace(content, phpTrim(cssStr), sNewCssStr);
                }

                isStyle= false; //收集CssStyle样式结束
            }else{
                cssStr= cssStr + s + vbCrlf();
            }
        }
    }
    if( sType== "替换路径" ){
        handleHtmlStyleCss= content;
    }else{
        handleHtmlStyleCss= cssStr;
    }
    return handleHtmlStyleCss;
}

//替换css内容里图片路径 20160718home
string replaceCssImgPath( string cssStr, string imgPath){
    string content=""; string[] splStr; string s=""; string c=""; string fileName=""; string toImgPath=""; string lCaseS="";
    content= getArray(cssStr, "\\(", "\\)", false, false);
    splStr= aspSplit(content, "$Array$");
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            lCaseS= lCase(s);
            fileName= aspTrim(getFileAttr(s, "2"));
            //引用图片 基础判断方法，有待改进，改进也是成长的一部分20160719
            if( inStr(lCaseS, "data:image")== 0 &&(inStr(lCaseS, ".jpg") > 0 || inStr(lCaseS, ".gif") > 0 || inStr(lCaseS, ".png") > 0 || inStr(lCaseS, ".bmp") > 0) ){
                //修复于20160719
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

//处理成干净的Css内容  CSS格式化
string handleCleanCss( string content){
    string[] splStr; string s=""; string c=""; bool isAddStr; string customS="";
    content= replace(content, "{", vbCrlf() + "{" + vbCrlf());
    content= replace(content, "}", vbCrlf() + "}" + vbCrlf());
    content= replace(content, ";", ";" + vbCrlf());

    splStr= aspSplit(content, vbCrlf());
    isAddStr= false; //追加字符默认为假
    foreach(var eachs in splStr){
        s=eachs;
        s= trimVbCrlf(s);
        customS= ""; //自定义S值
        if( s != "" ){
            if( inStr(s, "{") > 0 && inStr(s, "}")== 0 ){
                isAddStr= true;
                customS= s;
            }else if( inStr(s, "}") > 0 ){
                isAddStr= false;
            }
            if( left(s, 1) != "{" ){ c= c + vbCrlf() ;}
            if( isAddStr== true ){ s= "    " + s ;}
            if( customS != "" ){ s= customS ;}//自定义值不为空则用自定义内容
            c= c + s;

        }
    }
    c= trimVbCrlf(c);
    c= replace(c, "    ;" + vbCrlf(), ""); //清除多于的分号
    c= replace(c, ";" + vbCrlf() + "}", vbCrlf() + "}"); //最后一个参数不要分号
    return c;
}



//移除内容中多除行
string removeExcessRow(string content){
    string[] splStr; string s=""; string c=""; string tempS="";
    splStr= aspSplit(content, vbCrlf()); //分割行
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
//向Css里追加样式  a=CssAddToStyle(GetFText("1.html")," .test {color:#FF0f000; font-size:10px; float:left}")
string cssAddToStyle(string content, string addToStyle){
    string styleName=""; string yunStyleStr=""; string replaceStyleStr=""; string c="";
    if( inStr(addToStyle, "{") > 0 ){
        styleName= aspTrim(mid(addToStyle, 1, inStr(addToStyle, "{") - 1));
    }
    yunStyleStr= findCssStyle(content, styleName);
    replaceStyleStr= cssStyleAddToParam(yunStyleStr, addToStyle); //Css样式累加参数
    content= replace(content, yunStyleStr, replaceStyleStr);
    //C = C & "<hr>Content=" & Content
    string cssAddToStyle= content;
    //CssAddToStyle = YunStyleStr
    //CssAddToStyle = "StyleName=" & StyleName & "<hr>YunStyleStr=" & YunStyleStr & "<hr>ReplaceStyleStr=" & ReplaceStyleStr
    return cssAddToStyle;
}

//检测Css内容中是否有指定样式
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


//Css样式累加参数
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

//根据Css名称找到对应Css块
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
//导入网站前端要用到
//================================================
//处理截取到的Css
string handleCutCssCode(string dirPath, string cssStr){
    string content=""; string startStr=""; string endStr=""; string[] splStr; string imageFile=""; string fileName=""; string listStr="";
    startStr= "url\\(" ; endStr= "\\)";
    content= getArray(cssStr, startStr, endStr, false, false);
    splStr= aspSplit(content, "$Array$");
    foreach(var eachimageFile in splStr){
        imageFile=eachimageFile;
        if( imageFile != "" && inStr(imageFile, ".") > 0 && inStr(vbCrlf() + listStr + vbCrlf(), vbCrlf() + imageFile + vbCrlf())== 0 ){//对重复使用的图片处理
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

//处理截取到的HtmlDiv
string handleCutDivCode(string dirPath, string divStr){
    string content=""; string startStr=""; string endStr=""; string[] splStr; string imageFile=""; string toImageFile=""; string fileName=""; bool isHandle;
    startStr= "url\\(" ; endStr= "\\)";
    content= getArray(divStr, startStr, endStr, false, false);
    splStr= aspSplit(content, "$Array$");
    foreach(var eachimageFile in splStr){
        imageFile=eachimageFile;

        if( imageFile != "" && inStr(imageFile, ".") > 0 && inStr(imageFile, "{$#")== 0 ){
            //判断是否有域名 20150202
            if( getWebSite(imageFile)== "" ){
                fileName= replace(replace(replace(imageFile, "\"", ""), "'", ""), "\\", "/");
                if( inStr(fileName, "/") > 0 ){
                    fileName= mid(fileName, inStrRev(fileName, "/") + 1,-1);
                }
                divStr= replace(divStr, imageFile, dirPath + imageFile);
            }
        }
    }
    //图片处理
    //Content = GetIMG(DivStr) & vbCrlf & GetHtmlBackGroundImgList(DivStr)        '再加个Html背景图片
    content= getImgJsUrl(divStr, "不重复") + vbCrlf() + getHtmlBackGroundImgList(divStr); //再加个Html背景图片  加强版20150126
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachimageFile in splStr){
        imageFile=eachimageFile;
        if( imageFile != "" ){ //当链接地址当前为HTTP:时则不处理20150313
            isHandle= false;

            if( left(imageFile, 1)== "\\" ){
                //等处理20150817
            }else if( inStr(imageFile, ".") > 0 && left(imageFile, 5) != "HTTP:" && inStr(imageFile, "{$#")== 0 ){
                isHandle= true;
            }
            if( isHandle== true ){
                toImageFile= dirPath + removeFileDir(imageFile); //移除文件路径目录
                //html中图片路径替换
                divStr= replace(divStr, "\"" + imageFile + "\"", "\"" + toImageFile + "\"");
                divStr= replace(divStr, "'" + imageFile + "'", "\"" + toImageFile + "\"");
                divStr= replace(divStr, "=" + imageFile + " ", "\"" + toImageFile + "\"");
                divStr= replace(divStr, "=" + imageFile + ">", "\"" + toImageFile + "\"");
            }
        }
    }
    return divStr;
}

//获得HTMl里背景图片 20150116  如： <td width="980" height="169" background="kslx3bg.jpg">
string getHtmlBackGroundImgList( string content){
    content= getArray(content, " background=\"", "\"", false, false);
    content= replace(content, "$Array$", vbCrlf());
    return content;
}


//完善html里link css链接 Content = getHandleWebHtmlLink("/aa/bb/",Content)  外部调用
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
            //获得Css加强版，改于20141125
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


//获得css链接地址列表(20150824)  找<link>
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
            //获得Css加强版，改于20141125
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
//压缩html里的style样式 (20151008)
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
    replaceS= cssCompression(replaceS, 0) + vbCrlf(); //格式化CSS
    replaceS= removeBlankLines(replaceS);

    content= replace(content, serchS, replaceS);
    return content;
}



//读Css文件内容并处理(20150824) 加特殊记录标签   不常用
//如call rwend(handleReadCssContent("E:\E盘\WEB网站\至前网站\DataDir\VB模块\服务器\Template\模块功能列表\B站页面设计\home\home.css","aa",true))
string handleReadCssContent(string cssFilePath, string labelName, bool isHandleCss){
    string c=""; string startStr=""; string endStr="";
    c= getFText(cssFilePath);
    //截取CSS
    startStr= "/*CssCodeStart*/";
    endStr= "/*CssCodeEnd*/";
    if( inStr(c, startStr) > 0 && inStr(c, endStr) > 0 ){
        c= strCut(c, startStr, endStr, 2);
    }
    //处理CSS
    if( isHandleCss== true ){
        c= cssCompression(c, 0);
    }
    if( labelName != "" ){
        c= "/*" + labelName + " start*/" + c + "/*" + labelName + " end*/";
    }
    return c;
}

//处理Css样式里PX或%
string handleCssPX( string sValue){
    sValue= lCase(aspTrim(sValue));
    if( right(sValue, 1) != "%" && right(sValue, 2) != "px" ){
        sValue= sValue + "px";
    }
    return sValue;
}

</script>
