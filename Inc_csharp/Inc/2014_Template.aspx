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
//删除这种多余标签 <R#读出内容BlockName网站公告 start#>  <R#读出内容BlockName网站公告 end#>
//对模板处理



//读模块内容
string XY_ReadTemplateModule(string action){
    string moduleId=""; string filePath=""; string c=""; int i=0;
    string sourceList=""; //源内容列表 20150109
    string replaceList=""; //替换内容列表
    string[] splSource; string[] splReplace; string sourceStr=""; string replaceStr="";
    filePath= rParam(action, "File");
    moduleId= rParam(action, "ModuleId");
    sourceList= rParam(action, "SourceList");
    replaceList= rParam(action, "ReplaceList");
    //Call Echo(SourceList,ReplaceList)

    if( moduleId== "" ){ moduleId= rParam(action, "ModuleName") ;}//用块名称
    filePath= filePath + ".html";
    //Call Echo("FilePath",FilePath)
    //Call Echo("ModuleId",ModuleId)
    c= readTemplateModuleStr(filePath, "", moduleId);
    //加替换于20160331
    if( sourceList != "" && replaceList != "" ){
        splSource= aspSplit(sourceList, "[Array]");
        splReplace= aspSplit(replaceList, "[Array]");
        for( i= 0 ; i<= uBound(splSource); i++){
            sourceStr= splSource[i];
            replaceStr= splReplace[i];
            c= replace(c, sourceStr, replaceStr);
        }
    }
    return c;
}

//读模块内容
string readTemplateModuleStr(string filePath, string defaultContent, string moduleId){
    string startStr=""; string endStr=""; string content="";
    string readTemplateModuleStr="";
    startStr= "<!--#Module " + moduleId + " start#-->";
    endStr= "<!--#Module " + moduleId + " end#-->";
    //FilePath = ReplaceGlobleLable(FilePath)                '替换全部标签        '添加于2014 12 11

    //文件不存在，则追加模板路径 20150616 给VB软件里用
    if( checkFile(filePath)== false ){
        filePath= cfg_webTemplate + "/" + filePath;
    }

    //没有handleRGV这个函数了
    //filePath = handleRGV(filePath, "[$模块目录$]", "Module/")                              'Module

    if( defaultContent != "" ){
        content= defaultContent;
    }else if( checkFile(filePath)== true ){
        content= getFText(filePath);
    }else{
        content= code; //默认用内容指定内容
    }

    if( inStr(content, startStr)== 0 ){
        startStr= "<!--#Module " + moduleId + " Start#-->";
    }
    if( inStr(content, endStr)== 0 ){
        endStr= "<!--#Module " + moduleId + " End#-->";
    }
    if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
        readTemplateModuleStr= strCut(content, startStr, endStr, 2);
    }else{
        readTemplateModuleStr= "模块[" + moduleId + "]不存在,路径=" + filePath;
        die(readTemplateModuleStr + content);
    }
    return readTemplateModuleStr;
}
//找模块对应内容
string findModuleStr(string content, string valueStr){
    string startStr=""; string endStr=""; string yuanStr=""; string replaceStr=""; int i=0; string block=""; string blockFile=""; string action="";
    for( i= 1 ; i<= 9; i++){
        startStr= "[$读出内容 " ; endStr= "$]";
        if( inStr(valueStr, startStr) > 0 && inStr(valueStr, endStr) > 0 ){
            action= strCut(valueStr, startStr, endStr, 2);
            block= rParam(action, "Block");
            blockFile= rParam(action, "File");
            if( inStr(vbCrlf() + readBlockList + vbCrlf(), vbCrlf() + block + vbCrlf())== 0 ){
                readBlockList= readBlockList + block + vbCrlf();
            }
            //块文件存在 则读出内容
            if( blockFile != "" ){
                content= getFText(blockFile);
            }
            yuanStr= startStr + action + endStr;
            replaceStr= "";

            startStr= "<R#读出内容" + block + " start#>" ; endStr= "<R#读出内容" + block + " end#>";
            if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
                replaceStr= strCut(content, startStr, endStr, 2);
            }else{
                startStr= "<!--#读出内容" + block ; endStr= "#-->";
                if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
                    replaceStr= strCut(content, startStr, endStr, 2);
                }
            }
            //Call Eerr(YuanStr,ReplaceStr)
            valueStr= replace(valueStr, yuanStr, replaceStr);
            //Call Echo("ValueStr",ValueStr)
        }else{
            //没有模块要处理了 则退出
            break;
        }
    }
    return valueStr;
}

//读出Left模板样式        这里面字符 ' 来回复制会出错，所以我们用 \|*|\ 代替处理
string XY_ReadColumeSetTitle(string action){
    string startStr=""; string endStr=""; string style=""; string title=""; string valueStr=""; string moreClass=""; string moreUrl=""; string moreStr=""; string aStr=""; string c="";
    action= handleInModule(action, "start");
    style= rParam(action, "style");
    title= rParam(action, "Title");
    //Call Echo("ContentHeight",ContentHeight)
    //ValueStr = RParam(Action,"value")
    //根据模块找内容
    valueStr= moduleFindContent(action, "value");
    //Call Eerr("ValueStr",ValueStr)
    valueStr= findModuleStr(code, valueStr); //找模块对应内容

    moreClass= rParam(action, "MoreClass");
    moreUrl= phpTrim(rParam(action, "MoreUrl"));
    moreStr= rParam(action, "MoreStr");
    valueStr= handleInModule(valueStr, "end");
    c= readColumeSetTitle(action, style, title, valueStr);

    if( moreClass== "" ){ moreClass= "more" ;}//More链接为空 则用默认代替
    //If MoreUrl="" Then MoreUrl="#"                    'More链接网址为空 则用默认#代替
    //More链接样式不能为空，因为没有样式它就不能让More在最近边
    if( moreUrl != "" && moreStr != "" ){
        //AStr = "<a href='"& MoreUrl &"' class='"& MoreClass &"'>"& MoreStr &"</a>"
        aStr= "<a " + aHref(moreUrl, title, "") + " class='" + moreClass + "'>" + moreStr + "</a>";
        c= replace(c, "<!--#AMore#-->", aStr);
    }
    return c;
}

//读栏目并赋标题与内容值
string readColumeSetTitle(string action, string styleID, string columeTitle, string columeContent){
    string titleWidth=""; //标题宽度
    string titleHeight=""; //标题高度
    string contentHeight=""; //内容高度
    string contentWidth=""; //内容宽度
    string contentCss="";

    titleWidth= rParam(action, "TitleWidth"); //获得标题高度    待应用20150715
    titleHeight= rParam(action, "TitleHeight"); //获得标题宽度
    contentWidth= rParam(action, "ContentWidth"); //获得内容宽度
    contentHeight= rParam(action, "ContentHeight"); //获得内容高度

    //标题宽
    titleWidth= aspTrim(titleWidth);
    //自动加px单位，不加会无效果 20150115
    if( right(titleHeight, 1) != "%" && right(titleHeight, 2) != "px" && titleHeight != "" && titleHeight != "auto" ){
        titleHeight= titleHeight + "px";
    }
    if( right(titleWidth, 1) != "%" && right(titleWidth, 2) != "px" && titleWidth != "" && titleWidth != "auto" ){
        titleWidth= titleWidth + "px";
    }
    //内容高
    contentHeight= aspTrim(contentHeight);
    //自动加px单位，不加会无效果 20150115
    if( right(contentHeight, 1) != "%" && right(contentHeight, 2) != "px" && contentHeight != "" && contentHeight != "auto" ){
        contentHeight= contentHeight + "px";
    }
    //内容宽
    contentWidth= aspTrim(contentWidth);
    //自动加px单位，不加会无效果 20150115
    if( right(contentWidth, 1) != "%" && right(contentWidth, 2) != "px" && contentWidth != "" && contentWidth != "auto" ){
        contentWidth= contentWidth + "px";
    }

    if( contentHeight != "" ){
        contentCss= "height:" + contentHeight + ";";
    }
    if( contentWidth != "" ){
        contentCss= contentCss + "width:" + contentWidth + ";";
    }

    string content="";
    content= readColumn(styleID);
    //标题宽
    if( titleWidth != "" ){
        content= replace(content, "<div class=\"tvalue\">", "<div class=\"tvalue\" style='width:" + titleWidth + ";'>");
    }
    //内容高
    if( contentCss != "" ){
        content= replace(content, "<div class=\"ccontent\">", "<div class=\"ccontent\" style='" + contentCss + "'>");
    }
    //call echo(ContentWidth,ContentCss)

    content= replace(content, "栏目标题", columeTitle);
    content= replace(content, "栏目内容", columeContent);
    return content;
}

//读栏目模块
string readColumn(string sID){
    string templateFilePath=""; string startStr=""; string endStr=""; string s="";
    templateFilePath= cfg_webTemplate + "\\Template_Left.html";
    startStr= "/*columnlist" + sID + "Start*/";
    endStr= "/*columnlist" + sID + "End*/";
    s= readTemplateFileModular(templateFilePath, startStr, endStr);
    if( s== "[$NO$]" ){
        s= "Left样式ID[" + sID + "]不存在";
    }
    return s;
}


//读模板素材
string readTemplateSource(string sID){
    string templateFilePath=""; string startStr=""; string endStr=""; string s="";
    templateFilePath= cfg_webTemplate + "\\TemplateSource.html";
    startStr= "<!--#sourceHtml" + sID + "Start#-->";
    endStr= "<!--#sourceHtml" + sID + "End#-->";
    s= readTemplateFileModular(templateFilePath, startStr, endStr);
    if( s== "[$NO$]" ){
        s= "模板资源ID[" + sID + "]不存在";
    }
    return s;
}



//读模板文件中某模块
string readTemplateFileModular(string templateFilePath, string startStr, string endStr){
    string content="";
    string readTemplateFileModular= "";
    content= getFText(templateFilePath);
    if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
        readTemplateFileModular= strCut(content, startStr, endStr, 2);
    }else{
        readTemplateFileModular= "[$NO$]";
    }
    return readTemplateFileModular;
}

//读文件模板素材
string readTemplateFileSource(string templateFilePath, string sID){
    string startStr=""; string endStr=""; string s=""; string c="";
    string readTemplateFileSource="";
    startStr= "<!--#sourceHtml" + replace(sID, ".html", "") + "Start#-->";
    endStr= "<!--#sourceHtml" + replace(sID, ".html", "") + "End#-->";
    s= readTemplateFileModular(templateFilePath, startStr, endStr);
    if( s== "[$NO$]" ){
        //加一个读取本文本里配置列表(20150815)
        c= getStrCut(pubCode, startStr, endStr, 2);
        if( c != "" ){
            readTemplateFileSource= c;
            //call rwend(c)
            return readTemplateFileSource;
        }
        c= getFText(templateFilePath);
        //存在 <!--#TemplateSplitStart#-->  就返回当前全部内容
        if( inStr(c, "<!--#DialogStart#-->") > 0 ){
            readTemplateFileSource= c;
            return readTemplateFileSource;
        }

        s= "模板资源ID[" + sID + "]不存在,路径TemplateFilePath=" + handlePath(templateFilePath);
    }
    return s;
}
//读出文件展示列表资源
string readArticleListStyleSource(string sID){
    string filePath="";
    filePath= cfg_webImages + "\\文章展示样式\\" + sID;
    if( checkFile(filePath)== false ){
        filePath= cfg_webTemplate + "\\Resources\\" + sID;
    }
    //call echo(checkfile(filePath),filePath)
    string readArticleListStyleSource= readTemplateFileSource(filePath, sID);

    return readArticleListStyleSource;
}
//读出文件信息列表资源
string readArticleInfoStyleSource(string sID){
    string filePath="";
    filePath= cfg_webImages + "\\文章信息展示样式\\" + sID;
    if( checkFile(filePath)== false ){
        filePath= cfg_webTemplate + "\\Resources\\" + sID;
    }
    return readTemplateFileSource(filePath, sID);
}


</script>
