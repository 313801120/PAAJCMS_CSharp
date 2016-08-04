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
//动态生静态(2013,12,17)

//================ 快速获得网站操作 ==================
//在线修改 修改单文本
//MainStr = DisplayOnlineED2(WEB_ADMINURL &"MainInfo.Asp?act=ShowEdit&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainStr, "<li|<a ")
//在线修改 产品大类
//DidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditBigClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), DidStr, "<li|<a ")
//在线修改 产品小类
//SidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), SidStr, "<li|<a ")
//在线修改 产品子类
//S = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditThreeClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), S, "<li|<a ")
//在线修改  文章
//ProStr = DisplayOnlineED2(WEB_ADMINURL &"Product.Asp?act=ShowEditProduct&Id=" & TempRs("Id") & "&n=" & GetRnd(11), ProStr, "<li|<a ")
//在线修改 导航大类
//NavDidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavDidStr, "<li|<a ")
//在线修改 导航小类
//NavSidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavSmall&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavSidStr, "<li|<a ")

//-------------------------------- 下面为网站后台常用快捷标签代码区 -------------------------------------------

//符加文字颜色
string infoColor(string str, string color){
    if( color != "" ){ str= "<font color=" + color + ">" + str + "</font>" ;}
    return str;
}
//图片加载失败显示默认图片
string imgError(){
    return " onerror=\"this.src='/UploadFiles/NoImg.jpg'\"";
}
//获得target样式
string handleTargetStr( string sType){
    string handleTargetStr="";
    if( sType != "" ){
        handleTargetStr= " target='" + sType + "'";
    }
    return handleTargetStr;
}
//打开方式  (辅助)
string aTarget(string sType){
    return handleTargetStr(sType);
}
//获得链接Title样式
string aTitle( string title){
    string aTitle="";
    if( title != "" ){
        aTitle= " Title='" + title + "'";
    }
    return aTitle;
}
//获得链接Title
string imgAlt( string alt){
    string imgAlt="";
    if( alt != "" ){
        imgAlt= " alt='" + alt + "'";
    }
    return imgAlt;
}
//图片标题与Alt
string imgTitleAlt( string str){
    string imgTitleAlt="";
    if( str != "" ){
        imgTitleAlt= " alt='" + str + "' title='" + str + "'";
    }
    return imgTitleAlt;
}
//获得A Rel值
string aRel( bool isType){
    string aRel="";
    if( isType== true ){
        aRel= " rel='nofollow'";
    }
    return aRel;
}
//获得target样式
string styleClass( string className){
    string styleClass="";
    if( className != "" ){
        styleClass= " class='" + className + "'";
    }
    return styleClass;
}
//文本加粗
string textFontB( string text, bool isFontB){
    if( isFontB== true ){
        text= "<strong>" + text + "</strong>";
    }
    return text;
}
//文本加颜色
string textFontColor( string text, string color){
    if( color != "" ){
        text= "<font color='" + color + "'>" + text + "</font>";
    }
    return text;
}
//处理文本颜色与加粗
string fontColorFontB(string title, bool isFontB, string fontColor){
    return textFontColor(textFontB(title, isFontB), fontColor);
}
//获得默认文章信息文件名称
string getDefaultFileName(){
    return format_Time(now(), 6);
}
//获得链接  例：'"<a " & AHref(Url, TempRs("BigClassName"), TempRs("Target")) & ">" & TempRs("BigClassName") & "</a>"
string aHref(string url, string title, string target){
    url= handleHttpUrl(url); //处理一下URL 让之完整
    return "href='" + url + "'" + aTitle(title) + aTarget(target);
}
//获得图片路径
string imgSrc(string url, string title, string target){
    url= handleHttpUrl(url); //处理一下URL 让之完整
    return "src='" + url + "'" + aTitle(title) + imgAlt(title) + aTarget(target);
}

//============== 网站后台使用 ==================

//选择Target打开方式
string selectTarget(string target){
    string c=""; string sel="";
    c= c + "<select name=\"Target\" id=\"Target\">" + vbCrlf();
    c= c + "  <option value=''>链接打开方式</option>" + vbCrlf();
    if( target== "" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "  <option" + sel + " value=''>本页打开</option>" + vbCrlf();
    if( target== "_blank" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "  <option value=\"_blank\"" + sel + ">新页打开</option>" + vbCrlf();
    if( target== "Index" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "  <option value=\"Index\"" + sel + ">Index页打开</option>" + vbCrlf();
    if( target== "Main" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "  <option value=\"Main\"" + sel + ">Main页打开</option>" + vbCrlf();
    c= c + "</select>" + vbCrlf();
    return c;
}
//选择文本颜色
string selectFontColor(string fontColor){
    string c=""; string sel="";
    c= c + "  <select name=\"FontColor\" id=\"FontColor\">" + vbCrlf();
    c= c + "    <option value=''>文本颜色</option>" + vbCrlf();
    if( fontColor== "Red" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"Red\" class=\"FontColor_Red\"" + sel + ">红色</option>" + vbCrlf();
    if( fontColor== "Blue" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"Blue\" class=\"FontColor_Blue\"" + sel + ">蓝色</option>" + vbCrlf();
    if( fontColor== "Green" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"Green\" class=\"FontColor_Green\"" + sel + ">绿色</option>" + vbCrlf();
    if( fontColor== "Black" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"Black\" class=\"FontColor_Black\"" + sel + ">黑色</option>" + vbCrlf();
    if( fontColor== "White" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"White\" class=\"FontColor_White\"" + sel + ">白色</option>" + vbCrlf();
    c= c + "  </select>" + vbCrlf();
    return c;
}
//选择男女
string selectSex(string sex){
    string c=""; string sel="";
    c= c + "  <select name=\"FontColor\" id=\"FontColor\">" + vbCrlf();
    c= c + "    <option value=\"男\">男</option>" + vbCrlf();
    sel= IIF(sex== "女", " selected", "");
    c= c + "    <option value=\"女\"" + sel + ">女</option>" + vbCrlf();
    c= c + "  </select>" + vbCrlf();
    return c;
}
//选择Session或Cookies验证
string selectSessionCookies(string verificationMode){
    string c=""; string sel="";
    c= c + "  <select name=\"VerificationMode\" id=\"VerificationMode\">" + vbCrlf();
    c= c + "    <option value=\"1\">Session验证</option>" + vbCrlf();
    sel= IIF(verificationMode== "0", " selected", "");
    c= c + "    <option value=\"0\"" + sel + ">Cookies验证</option>" + vbCrlf();
    c= c + "  </select>" + vbCrlf();
    return c;
}
//显示选择分割内容  showSelectList("aa","aa|bb|cc","|","bb")
string showSelectList(string IDName, string content, string sSplType, string thisValue){
    string c=""; string sel=""; string[] splStr; string s="";
    IDName= aspTrim(IDName);
    if( sSplType== "" ){ sSplType= "|_-|" ;}
    if( IDName != "" ){ c= c + "  <select name=\"" + IDName + "\" id=\"" + IDName + "\">" + vbCrlf() ;}

    splStr= aspSplit(content, sSplType);
    foreach(var eachs in splStr){
        s=eachs;
        sel= "";
        if( s== thisValue ){ sel= " selected" ;}
        c= c + "    <option value=\"" + s + "\"" + sel + ">" + s + "</option>" + vbCrlf();
    }
    if( IDName != "" ){ c= c + "  </select>" + vbCrlf() ;}
    return c;
}

//显示文章展示列表样式 20150114   例 Call Rw(ShowArticleListStyle("下载列表二.html"))
string showArticleListStyle( string thisValue){
    return handleArticleListStyleOrInfoStyle("文章展示样式", "ArticleListStyle", thisValue);
}
//显示文章信息展示样式 20150114   例 Call Rw(ShowArticleInfoStyle("下载列表二.html"))
string showArticleInfoStyle( string thisValue){
    return handleArticleListStyleOrInfoStyle("文章信息展示样式", "ArticleInfoStyle", thisValue);
}
//处理文章展示列表样式和文章信息样式
string handleArticleListStyleOrInfoStyle(string folderName, string inputName, string thisValue){
    string resourceDir=""; string content=""; string c=""; string[] splStr; string fileName=""; string sel="";

    resourceDir= cfg_webImages + "\\" + folderName + "\\";

    content= getDirHtmlNameList(resourceDir);

    thisValue= lCase(thisValue); //转成小写 好对比

    c= c + "  <select name=\"" + inputName + "\" id=\"" + inputName + "\">" + vbCrlf();
    c= c + "    <option value=\"\"></option>" + vbCrlf();
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachfileName in splStr){
        fileName=eachfileName;
        if( fileName != "" ){
            sel= IIF(lCase(fileName)== thisValue, " selected", "");
            c= c + "    <option value=\"" + fileName + "\"" + sel + ">" + fileName + "</option>" + vbCrlf();
        }
    }
    c= c + "  </select>" + vbCrlf();

    return c;
}

//获得模块皮肤 ShowWebModuleSkins("ModuleSkins", ModuleSkins)
string showWebModuleSkins(string inputName, string thisValue){
    string resourceDir=""; string content=""; string c=""; string[] splStr; string fileName=""; string sel="";
    resourceDir= cfg_webImages + "\\Index\\column";
    //Call Echo("ResourceDir",ResourceDir)
    content= getDirFolderNameList(resourceDir);
    //Call Echo("Content",Content)

    thisValue= lCase(thisValue); //转成小写 好对比

    c= c + "  <select name=\"" + inputName + "\" id=\"" + inputName + "\">" + vbCrlf();
    c= c + "    <option value=\"\"></option>" + vbCrlf();
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachfileName in splStr){
        fileName=eachfileName;
        if( fileName != "" ){
            sel= IIF(lCase(fileName)== thisValue, " selected", "");
            c= c + "    <option value=\"" + fileName + "\"" + sel + ">" + fileName + "</option>" + vbCrlf();
        }
    }
    c= c + "  </select>" + vbCrlf();

    return c;
}

//显示单选项列表
string showRadioList(string IDName, string content, string sSplType, string thisValue){
    string c=""; string sel=""; string[] splStr; string s=""; int i=0;
    IDName= aspTrim(IDName);
    if( sSplType== "" ){ sSplType= "|_-|" ;}
    i= 0;
    splStr= aspSplit(content, sSplType);
    foreach(var eachs in splStr){
        s=eachs;
        sel= "" ; i= i + 1;
        if( s== thisValue ){ sel= " checked" ;}
        c= c + "<input type=\"radio\" name=\"" + IDName + "\" id=\"" + IDName + i + "\" value=\"radio\" " + sel + "><label for=\"" + IDName + i + "\">" + s + "</label>" + vbCrlf();
    }

    return c;
}
//显示Input复选 InputCheckBox("Id",ID,"")
string inputCheckBox(string textName, bool isChecked, string helpStr){
    //Dim sel
    //If CStr(valueStr) = "True" Or CStr(isChecked) = "1" Then sel = " isChecked" Else sel = ""
    //inputCheckBox = "<input type='checkbox' name='" & textName & "' id='" & textName & "'" & sel & " value='1'>"
    //If helpStr <> "" Then inputCheckBox = "<label for='" & textName & "'>" & inputCheckBox & helpStr & "</label> "
    return handleInputCheckBox(textName, isChecked, "1", helpStr, "");
}
//显示Input复选 InputCheckBox("Id",ID,"")
string inputCheckBox3(string textName, bool isChecked, string valueStr, string helpStr){
    return handleInputCheckBox(textName, isChecked, valueStr, helpStr, "newidname");
}
string handleInputCheckBox(string textName, bool isChecked, string valueStr, string helpStr, string sType){
    string s=""; string sel=""; string idName="";
    if( cStr(valueStr)== "True" || isChecked== true ){ sel= " checked" ;}else{ sel= "" ;}
    idName= textName; //id名等于文件名称
    sType= "|" + sType + "|";
    if( inStr(sType, "|newidname|") > 0 ){
        idName= textName + phpRand(1, 9999);
    }
    s= "<input type='checkbox' name='" + textName + "' id='" + idName + "'" + sel + " value='" + valueStr + "'>";
    if( helpStr != "" ){ s= "<label for='" + idName + "'>" + s + helpStr + "</label> " ;}
    return s;
}

//显示Input文本  InputText("FolderName", FolderName, "40px", "帮助文字")
string inputText(string textName, string valueStr, string width, string helpStr){
    string css="";

    width= aspTrim(lCase(width));
    if( width != "" ){
        if( right(width, 1) != "%" && right(width, 2) != "px" ){
            width= width + "px";
        }
        css= " style='width:" + width + ";'";
    }
    return "<input name=\"" + textName + "\" type=\"text\" id=\"" + textName + "\" value=\"" + valueStr + "\"" + css + " />" + helpStr;
}
//显示Input文本  InputText("FolderName", FolderName, "40px", "帮助文字")
string inputText2(string textName, string valueStr, string width, string className, string helpStr){
    string css="";
    if( className != "" ){
        className= " class=\"" + className + "\"";
    }
    width= aspTrim(lCase(width));
    if( width != "" ){
        if( right(width, 1) != "%" && right(width, 2) != "px" ){
            width= width + "px";
        }
        css= " style='width:" + width + ";'";
    }
    return "<input name=\"" + textName + "\" type=\"text\" id=\"" + textName + "\" value=\"" + valueStr + "\"" + css + className + " />" + helpStr;
}
//显示Input文本在左边  InputLeftText(TextName, ValueStr, "98%", "")
string inputLeftText(string textName, string valueStr, string width, string helpStr){
    string css="";
    width= aspTrim(lCase(width));
    if( width != "" ){
        if( right(width, 1) != "%" && right(width, 2) != "px" ){
            width= width + "px";
        }
        css= " style='width:" + width + ";'";
    }
    return helpStr + "<input name=\"" + textName + "\" type=\"text\" id=\"" + textName + "\" value=\"" + valueStr + "\"" + css + " />" + vbCrlf();
}
//显示Input文本在左边 帮助文字在右边
string inputLeftTextHelpTextRight(string textName, string valueStr, string width, string helpStr){
    string css="";
    width= aspTrim(lCase(width));
    if( width != "" ){
        if( right(width, 1) != "%" && right(width, 2) != "px" ){
            width= width + "px";
        }
        css= " style='width:" + width + ";'";
    }
    return "<input name=\"" + textName + "\" type=\"text\" id=\"" + textName + "\" value=\"" + valueStr + "\"" + css + " />" + helpStr;
}
//显示Input文本在中边 提示文本在左边
string inputLeftTextContent(string textName, string valueStr, string width, string helpStr){
    return handleInputLeftRightTextContent("左边", textName, valueStr, width, helpStr);
}
//显示Input文本在中边 提示文本在右边
string inputRightTextContent(string textName, string valueStr, string width, string helpStr){
    return handleInputLeftRightTextContent("右边", textName, valueStr, width, helpStr);
}
//显示Input文本在中边 提示文本在左边 或 提示文本在右边 20150114
string handleInputLeftRightTextContent(string sType, string textName, string valueStr, string width, string helpStr){
    string css="";
    width= aspTrim(lCase(width));
    if( width != "" ){
        if( right(width, 1) != "%" && right(width, 2) != "px" ){
            width= width + "px";
        }
        css= " style='width:" + width + ";'";
    }
    if( css== "" ){
        css= " style='text-align:center;'";
    }else{
        css= replace(css, ";'", ";text-align:center;'");
    }
    string handleInputLeftRightTextContent= "<input name=\"" + textName + "\" type=\"text\" id=\"" + textName + "\" value=\"" + valueStr + "\"" + css + " />";

    if( sType== "左边" ){
        handleInputLeftRightTextContent= helpStr + handleInputLeftRightTextContent + vbCrlf();
    }else{
        handleInputLeftRightTextContent= handleInputLeftRightTextContent + helpStr;
    }

    return handleInputLeftRightTextContent;
}

//显示Input文本在左边密码
string inputLeftPassText(string textName, string valueStr, string width, string helpStr){
    string css="";
    width= aspTrim(lCase(width));
    if( width != "" ){
        if( right(width, 1) != "%" && right(width, 2) != "px" ){
            width= width + "px";
        }
        css= " style='width:" + width + ";'";
    }
    return helpStr + "<input name=\"" + textName + "\" type=\"password\" id=\"" + textName + "\" value=\"" + valueStr + "\"" + css + " />" + vbCrlf();
}
//显示Input文本在左边密码类型
string inputLeftPassTextContent(string textName, string valueStr, string width, string helpStr){
    string css="";
    width= aspTrim(lCase(width));
    if( width != "" ){
        if( right(width, 1) != "%" && right(width, 2) != "px" ){
            width= width + "px";
        }
        css= " style='width:" + width + ";'";
    }
    if( css== "" ){
        css= " style='text-align:center;'";
    }else{
        css= replace(css, ";'", ";text-align:center;'");
    }
    return helpStr + "<input name=\"" + textName + "\" type=\"password\" id=\"" + textName + "\" value=\"" + valueStr + "\"" + css + " />" + vbCrlf();
}
//显示Input隐藏文本
string inputHiddenText(string textName, string valueStr){
    return "<input name=\"" + textName + "\" type=\"hidden\" id=\"" + textName + "\" value=\"" + valueStr + "\" />" + vbCrlf();
}
//显示Input文本域 InputTextArea("FindTpl", FindTpl, "60%" , "120px", "")
string inputTextArea(string textName, string valueStr, string width, string height, string helpStr){
    string css=""; string heightStr="";
    width= aspTrim(lCase(width));
    if( width != "" ){
        if( right(width, 1) != "%" && right(width, 2) != "px" ){
            width= width + "px";
        }
        css= " style='width:" + width + ";'";
    }
    if( height != "" ){
        if( checkNumber(height) ){ //自动加个px像素
            height= height + "px";
        }
        heightStr= "height:" + height + ";";
        if( css != "" ){
            css= replace(css, ";'", ";" + heightStr + ";'");
        }else{
            css= " style='height:" + height + ";'";
        }
    }
    css= replace(css, ";;", ";"); //去掉多余的值
    return "<textarea name=\"" + textName + "\" type=\"text\" id=\"" + textName + "\"" + css + ">" + valueStr + "</textarea>" + helpStr;
}
//显示隐藏Input文本域 InputTextArea("WebDescription", WebDescription, "99%", "100px", "")
string inputHiddenTextArea(string textName, string valueStr, string width, string height, string helpStr){
    return handleInputHiddenTextArea(textName, valueStr, width, height, "", helpStr);
}
//显示隐藏Input文本域 InputTextArea("WebDescription", WebDescription, "99%", "100px", "")
string handleInputHiddenTextArea(string textName, string valueStr, string width, string height, string className, string helpStr){
    string css=""; string heightStr="";
    if( className != "" ){
        className= " class=\"" + className + "\"";
    }
    if( width != "" ){ css= " style='width:" + width + ";'" ;}
    if( height != "" ){
        heightStr= "height:" + height + ";";
        if( css != "" ){
            css= replace(css, ";'", ";" + heightStr + ";'");
        }else{
            css= " style='height:" + height + ";display:none;'";
        }
    }
    return "<textarea name=\"" + textName + "\" type=\"text\" id=\"" + textName + "\"" + css + className + ">" + valueStr + "</textarea>" + helpStr;
}
//显示目录列表 以Select方式显示
string showSelectDirList(string folderPath, string valueStr){
    string[] splStr; string c=""; string fileName=""; string sel="";
    splStr= aspSplit(getDirFileSort(folderPath), vbCrlf());
    foreach(var eachfileName in splStr){
        fileName=eachfileName;
        if( fileName != "" ){
            sel= IIF(valueStr== fileName, " selected", "");
            c= c + "<option value=\"" + folderPath + fileName + "\" " + sel + ">" + fileName + "</option>" + vbCrlf();
        }
    }
    return c;
}
//给Input加个Disabled不可操作
string inputDisabled( string content){
    return replace(content, "<input ", "<input disabled=\"disabled\" ");
}

//给Input加个rel关系内容
string inputAddAlt( string content, string altStr){
    string searchStr=""; string replaceStr="";
    searchStr= "<input ";
    replaceStr= searchStr + "alt=\"" + altStr + "\" ";
    if( inStr(content, searchStr) > 0 ){
        content= replace(content, searchStr, replaceStr);
    }else{
        searchStr= "<textarea ";
        replaceStr= searchStr + "alt=\"" + altStr + "\" ";
        if( inStr(content, searchStr) > 0 ){
            content= replace(content, searchStr, replaceStr);
        }
    }
    return content;
}



//快速调用设置====================================================

//网站描述
string webTitle_InputTextArea(string webTitle){
    return inputText("WebTitle", webTitle, "70%", "  多个关键词用-隔开"); //不填为网站默认标题
}
//网站关键词
string webKeywords_InputText(string webKeywords){
    return inputText("WebKeywords", webKeywords, "70%", " 请以，隔开(中文逗号)");
}
//网站描述
string webDescription_InputTextArea(string webDescription){
    return inputTextArea("WebDescription", webDescription, "99%", "100px", "");
}
//静态文件夹名
string folderName_InputText(string folderName){
    return inputText("FolderName", folderName, "40%", "");
}
//静态文件名
string fileName_InputText(string fileName){
    return inputText("FileName", fileName, "40%", ".html 也可以是网络上的链接地址");
}
//模板文件名

string templatePath_InputText(string templatePath){
    return inputText("TemplatePath", templatePath, "40%", " 不填为默认");
}
//获得拼音按钮内容
string clickPinYinHTMLStr(string did){
    return "<a href=\"javascript:GetPinYin('FolderName','" + did + "','AjAx.Asp?act=GetPinYin')\" >获得拼音</a>";
}
//选择文本颜色与文本加粗
string showFontColorFontB(string fontColor, bool isFontB){
    return selectFontColor(fontColor) + inputCheckBox("FontB", isFontB, "加粗");
}
//显示文本TEXT排序
string showSort(string sort){
    string showSort= inputText("Sort", sort, "30px", "");
    return replace(showSort, ";'", ";text-align:center;'");
}
//网站导航类型顶部底部等
string showWebNavType(bool isNavTop, bool isNavButtom, bool isNavLeft, bool isNavContent, bool isNavRight, bool isNavOthre){
    string c="";
    c= c + inputCheckBox("NavTop", isNavTop, "顶部导航");
    c= c + inputCheckBox("NavButtom", isNavButtom, "底部导航");
    c= c + inputCheckBox("NavLeft", isNavLeft, "左边导航");
    c= c + inputCheckBox("NavContent", isNavContent, "中间导航");
    c= c + inputCheckBox("NavRight", isNavRight, "右边导航");
    c= c + inputCheckBox("NavOthre", isNavOthre, "其它导航");
    return c;
}
string showOnHtml(bool isOnHtml){
    return inputCheckBox("OnHtml", isOnHtml, "生成HTML");
}
string showThrough(bool isThrough){
    return inputCheckBox("Through", isThrough, "审核");
}
string showRecommend(bool isRecommend){
    return inputCheckBox("Recommend", isRecommend, "推荐");
}
//显示开户与关闭图片
string showOnOffImg(string id, string table, string fieldName, bool isRecommend, string url){
    string temp=""; string img=""; string aUrl="";
    if( rq("page") != "" ){ temp= "&page=" + rq("page") ;}else{ temp= "" ;}
    if( isRecommend== true ){
        img= "<img src=\"" + adminDir + "Images/yes.gif\">";
    }else{
        img= "<img src=\"" + adminDir + "Images/webno.gif\">";
    }
    //Call Echo(GetUrl(),""& adminDir &"HandleDatabase.Asp?act=SetTrueFalse&Table=" & Table & "&FieldName=" & FieldName & "&Url=" & Url & "&Id=" & Id)
    aUrl= getUrlAddToParam(getUrl(), "" + adminDir + "HandleDatabase.Asp?act=SetTrueFalse&Table=" + table + "&FieldName=" + fieldName + "&Url=" + url + "&Id=" + id, "replace");
    string showOnOffImg= "<a href=\"" + aUrl + "\">" + img + "</a>";
    //旧版
    //ShowOnOffImg = "<a href="& adminDir &"HandleDatabase.Asp?act=SetTrueFalse&Table=" & Table & "&FieldName=" & FieldName & "&Url=" & Url & "&Id=" & Id & Temp & ">" & Img & "</a>"
    return showOnOffImg;
}
//显示开户与关闭图片
string newShowOnOffImg(string id, string table, string fieldName, bool isRecommend, string url){
    string temp=""; string img="";
    if( rq("page") != "" ){ temp= "&page=" + rq("page") ;}else{ temp= "" ;}
    if( isRecommend== true ){
        img= "<img src=\"/Images/yes.gif\">";
    }else{
        img= "<img src=\"/Images/webno.gif\">";
    }
    return "<a href=/WebAdmin/ZAction.Asp?act=Through&Table=" + table + "&FieldName=" + fieldName + "&Url=" + url + "&Id=" + id + temp + ">" + img + "</a>";
}


//获得控制Css样式 20150128  暂时不用
string controlDialogCss(){
    string c="";
    c= "<style>" + vbCrlf();
    c= c + "/*控制Css20150128*/" + vbCrlf();
    c= c + ".controlDialog{" + vbCrlf();
    c= c + "    position:relative;" + vbCrlf();
    c= c + "    height:50px;" + vbCrlf();
    c= c + "    width:auto;" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + ".controlDialog .menu{" + vbCrlf();
    c= c + "    position:absolute;" + vbCrlf();
    c= c + "    right:0px;" + vbCrlf();
    c= c + "    top:0px;" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + ".controlDialog .menu a{" + vbCrlf();
    c= c + "    color:#FF0000;" + vbCrlf();
    c= c + "    font-size:14px;" + vbCrlf();
    c= c + "    text-decoration:none;" + vbCrlf();
    c= c + "    background-color:#FFFFFF;" + vbCrlf();
    c= c + "    border:1px solid #003300;" + vbCrlf();
    c= c + "    padding:4px;" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + ".controlDialog .menu a:hover{" + vbCrlf();
    c= c + "    color:#C60000;" + vbCrlf();
    c= c + "    text-decoration:underline;" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "</style>" + vbCrlf();
    return c;
}


//删除里暂存代码
string batchDeleteTempStr(string content, string startStr, string endStr){
    int i=0; string s="";
    for( i= 1 ; i<= 9; i++){
        if( inStr(content, startStr)== 0 ){
            break;
        }
        s= getStrCut(content, startStr, endStr, 1);
        content= replace(content, s, "");
    }
    return content;
}
</script>
