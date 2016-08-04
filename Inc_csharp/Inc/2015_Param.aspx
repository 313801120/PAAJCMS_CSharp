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


//http://127.0.0.1/函数/ClassAspCode.Asp?act=GetFileFunctionStrList    从这里面获得最新设置



//**************************************** 给php用 通用 ****************************************

//替换参数值 2014  12 01
string newReplaceValueParam(string content, string paramName, string replaceStr){
    string startStr=""; string endStr=""; string labelStr=""; string sLen=""; string sNTimeFormat=""; string delHtmlYes=""; string trimYes="";
    //ReplaceStr = ReplaceStr & "这里面放上内容在这时碳呀。"
    //ReplaceStr = CStr(ReplaceStr)            '转成字符类型
    if( isNul(replaceStr)== true ){ replaceStr= "" ;}

    startStr= "[$" + paramName ; endStr= "$]";
    if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
        labelStr= strCut(content, startStr, endStr, 1);
        //删除Html
        delHtmlYes= rParam(labelStr, "DelHtml"); //是否删除Html
        if( delHtmlYes== "true" ){ replaceStr= replace(delHtml(replaceStr), "<", "&lt;") ;}//HTML处理
        //删除两边空格
        trimYes= rParam(labelStr, "Trim"); //是否删除两边空格
        if( trimYes== "true" ){ replaceStr= trimVbCrlf(replaceStr) ;}

        //截取字符处理
        sLen= handleNumber(rParam(labelStr, "Len"));
        //If nLen<>"" Then ReplaceStr = CutStr(ReplaceStr,nLen,"null")' Left(ReplaceStr,nLen)
        if( sLen != "" ){ replaceStr= cutStr(replaceStr, cInt(sLen), "...") ;}//Left(ReplaceStr,nLen)


        //时间处理
        sNTimeFormat= rParam(labelStr, "Format_Time"); //时间处理值
        if( sNTimeFormat != "" ){
            replaceStr= format_Time(replaceStr, cInt(sNTimeFormat));
        }
        content= replace(content, labelStr, replaceStr);

    }
    return content;
}

//根据标签找到对应内容
string newRParam(string dataCode, string action, string moduleName){
    string defaultStr=""; string startStr=""; string endStr="";
    defaultStr= rParam(action, moduleName);
    startStr= "<!--#" + defaultStr + " start#-->";
    endStr= "<!--#" + defaultStr + " end#-->";

    if( defaultStr != "" ){
        //判断是否存在
        if( inStr(dataCode, startStr) > 0 && inStr(dataCode, endStr) > 0 ){
            defaultStr= strCut(dataCode, startStr, endStr, 2);
        }else{
            startStr= "<!--#" + defaultStr;
            endStr= "#-->";
            if( inStr(dataCode, startStr) > 0 && inStr(dataCode, endStr) > 0 ){
                defaultStr= strCut(dataCode, startStr, endStr, 2);

                //Call Echo("有","StartStr=" & StartStr & ",EndStr=" & EndStr  & ",Default=" & Default)
            }
        }
    }
    return defaultStr;
}




//**************************************** 给php用 通用 **************************************** end


//运行全部动作(20150827)
string getContentAllRunStr(string content){
    string[] splStr; string s=""; string c=""; string tempS="";
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            tempS= s;
            s= replace(replace(s, chr(10).ToString(), ""), chr(13).ToString(), ""); //奇怪为什么 s里会有 chr(10)与chr(13) 呢？
            c= c + tempS + "=" + handleContentCode(s, "") + vbCrlf();
        }
    }
    return c;
}

//获得内容运行后字符  用法 len(aaaa)  没双引单引
string getContentRunStr( string content){
    return handleContentCode(content, "");
}
//处理内容里有""，则给它删除掉20150329
//检测内容运行后字符
string checkContentRunStr( string content){
    return handleContentCode(content, "check");
}
//处理双引号
string handleDoubleQuotation( string s){
    string sNew="";
    sNew= phpTrim(s);
    if( left(sNew, 1)== "\"" && right(sNew, 1)== "\"" ){
        s= mid(sNew, 2, len(sNew) - 2);
    }
    return s;
}
//辅助上面
string strDQ( string s){
    return handleDoubleQuotation(s);
}
//处理成数据 20150330
string[] handleToArray(string content){
    string[] splStr; int i=0;
    content= strCut(content, "(", ")", 2);
    //Call Rw(Content)
    splStr= aspSplit(content, ",");
    //Call Rw("<hr>")
    for( i= 0 ; i<= uBound(splStr); i++){
        splStr[i]= strDQ(splStr[i]);
        //Call Echo(I,SplStr(I))
    }
    return splStr;
}
//处理内容里代码 引擎20150324   http://127.0.0.1/函数/ClassAspCode.Asp?act=GetFileFunctionStrList  获得最新    version 1.0
string handleContentCode(string content,string sType){return "";return "";//留空函数
    return "";
}


//内部模块处理 HandleInModule(Content,"start") HandleInModule(Content,"end")
string handleInModule(string content, string sType){
    sType= lCase(cStr(sType));
    if( sType== "1" || sType== "start" ){
        content= replace(content, "\\'", "\\|*|\\");
        content= replace(content, "\\=", "\\|&|\\"); //后加20141024
    }else if( sType== "2" || sType== "end" ){
        content= replace(content, "\\|*|\\", "'");
        content= replace(content, "\\$", "$");
        content= replace(content, "\\}", "}");

        content= replace(content, "\\|&|\\", "="); //后加20141024
    }
    return content;
}
//清除特殊样式后获得标签值
string clearRParam( string action, string lableStr){
    string s="";
    //Action=Replace(Action,"\'","【|\‘|】")
    action= replace(action, "\\'", ""); //把这种清掉
    s= rParam(action, lableStr);
    //s=replace(s,"【|\‘|】", "\'")
    return s;
}
//获得参数内容后 放到动作里处理一下（20151023）
string atRParam( string action, string lableStr){
    string atRParam= rParam(action, lableStr);
    if( inStr(atRParam, "{$") > 0 && inStr(atRParam, "$}") > 0 ){
        atRParam= handleTemplateAction(atRParam, false); //处理动作  注意这是之前旧版本
        atRParam= handleAction(atRParam); //这是新版本
    }
    return atRParam;
}
//读单个参数值  Title = RParam(Action,"Title")     起强版获取参数值20150723
string rParam( string action, string lableStr){
    string s="";

    //原始 单引号
    s= handleRParam(action, lableStr, "'");
    //原始 双引号
    if( s== "" ){
        s= handleRParam(action, lableStr, "\"");
    }
    //原始 空
    if( s== "" ){
        s= handleRParam(action, lableStr, "");
    }

    //小写 单引号
    if( s== "" ){
        s= handleRParam(action, lCase(lableStr), "'");
    }
    //小写 双引号
    if( s== "" ){
        s= handleRParam(action, lCase(lableStr), "\"");
    }
    //小写 空
    if( s== "" ){
        s= handleRParam(action, lCase(lableStr), "");
    }

    //大写 单引号
    if( s== "" ){
        s= handleRParam(action, uCase(lableStr), "'");
    }
    //大写 双引号
    if( s== "" ){
        s= handleRParam(action, uCase(lableStr), "\"");
    }
    //大写 空
    if( s== "" ){
        s= handleRParam(action, uCase(lableStr), "");
    }
    //不要这个，要不不稳定(20151022)
    //if s=false then s=""
    if( s== "[#空*值_#]" ){ s= "" ;}
    return s;
}
//处理 读单个参数值
string handleRParam( string action, string lableStr, string typeStr){
    string lalbeName=""; string endTypeStr=""; bool isTrue; string s="";
    isTrue= false; //是否为真
    endTypeStr= IIF(typeStr != "", typeStr, " ");
    action= vbCrlf() + " " + action; //给它也加个空格，要不然在没有函数，前面就没有空格
    //默认前面加空格
    lalbeName= " " + lableStr; //加个空格是为了精准
    //不存在  前面加点
    if( inStr(action, lalbeName + "=" + typeStr)== 0 && isTrue== false ){
        lalbeName= "'" + lableStr;
    }else{
        isTrue= true;
    }
    //不存在 前面加双引号
    if( inStr(action, lalbeName + "=" + typeStr)== 0 && isTrue== false ){
        lalbeName= "\"" + lableStr;
    }else{
        isTrue= true;
    }
    //不存在    前面加TAB
    if( inStr(action, lalbeName + "=" + typeStr)== 0 && isTrue== false ){
        lalbeName= vbTab() + lableStr;
    }else{
        isTrue= true;
    }
    //不存在    前面加换行
    if( inStr(action, lalbeName + "=" + typeStr)== 0 && isTrue== false ){
        lalbeName= vbCrlf() + lableStr;
    }else{
        isTrue= true;
    }
    if( inStr(action, lalbeName + "=" + typeStr) > 0 && inStr(action, endTypeStr) > 0 ){
        s= strCut(action, lalbeName + "=" + typeStr, endTypeStr, 2);
        s= handleInModule(s, "end"); //处理里面参数 追加于20141031            还原内容值

        if( s== "" ){
            s= "[#空*值_#]";
        }

        //判断是否对参数进行动作制作
        if( inStr(s, "{$") > 0 && inStr(s, "$}") > 0 ){

            //handleRParam=HandleTemplateAction(handleRParam,true)        '处理动作
            //handleRParam = handleModuleReplaceArray(handleRParam)'给AddSqL处理一下动作 这是处理替换，不需要，因为在 HandleTemplateAction有替换了(20151021)

        }
        //不要这个，要不不稳定(20151022)
        //if handleRParam="" then
        //handleRParam=false
        //end if
    }
    return s;
}


//获得配置块 20150105 GetConfigBlock(ConfigContent, BlockName)
string getConfigBlock(string configContent, string blockName){
    return getCutConfigBlock(configContent, blockName, "#", "#");
}
//获得配置块 20150105
string getConfigBlock2(string configContent, string blockName){
    return getCutConfigBlock(configContent, blockName, "[#", "#]");
}
//获得配置块 20150105
string getConfigBlock3(string configContent, string blockName){
    return getCutConfigBlock(configContent, blockName, "[$", "$]");
}
//截取配置内容中块 20150105
string getCutConfigBlock(string configContent, string blockName, string startLable, string endLable){
    string startStr=""; string endStr="";
    startStr= startLable + blockName + endLable;
    endStr= startLable + blockName + endLable;
    string getCutConfigBlock= "";
    //开始标签处理
    if( inStr(configContent, startStr + " start") > 0 ){
        startStr= startStr + " start";
    }else{
        startStr= startStr + " Start";
    }
    //结束标签处理
    if( inStr(configContent, endStr + " end") > 0 ){
        endStr= endStr + " end";
    }else{
        endStr= endStr + " End";
    }

    if( inStr(configContent, startStr) > 0 && inStr(configContent, endStr) > 0 ){
        getCutConfigBlock= strCut(configContent, startStr, endStr, 2);
    }
    return getCutConfigBlock;
}
//获得配置内容块20150401
string getConfigContentBlock( string configContent, string blockName){
    return getCutConfigBlock(configContent, blockName, "", "");
}
//获得配置文件里块20150401  getConfigFileBlock(ConfigPath, "#txtRunCode#")  测试标签块时则自动创建
string getConfigFileBlock( string configFile, string blockName){
    string content=""; string findStr=""; string replaceStr=""; string startStr=""; string endStr="";
    content= getFText(configFile);
    string getConfigFileBlock= "";
    //MsgBox ("ConfigFile=" & ConfigFile & "(" & checkFile(ConfigFile) & "，" & GetFSize(ConfigFile) & ")" & vbCrLf & "Content=" & Content)
    startStr= blockName + " start";
    endStr= blockName + " end";
    replaceStr= startStr + "" + endStr;
    if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
        findStr= strCut(content, startStr, endStr, 2);
        getConfigFileBlock= findStr;
    }else{
        createFile(configFile, content + replaceStr);
    }
    return getConfigFileBlock;
}
//设置配置文件里块 20150401 call setConfigFileBlock(ConfigFile, "aaabbc", "#上传目录列表#")  存在则更新
string setConfigFileBlock( string configFile, string writeContent, string blockName){
    string content=""; string findStr=""; string replaceStr=""; string startStr=""; string endStr="";
    content= getFText(configFile);
    startStr= blockName + " start";
    endStr= blockName + " end";
    replaceStr= startStr + writeContent + endStr;
    if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
        findStr= strCut(content, startStr, endStr, 1);
        content= replace(content, findStr, replaceStr);
    }else{
        content= content + replaceStr;
    }
    createFile(configFile, content);
    return content;
}

//删除配置块 20150322
string delConfigBlock(string config, string blockName){
    return delCutConfigBlock(config, blockName, "#", "#");
}
//删除配置块 20150322
string delConfigBlock2(string config, string blockName){
    return delCutConfigBlock(config, blockName, "[#", "#]");
}

//删除配置块 20150322
string delConfigBlock3(string config, string blockName){
    return delCutConfigBlock(config, blockName, "[$", "$]");
}
//删除配置内容 20150322
string delCutConfigBlock(string config, string blockName, string startLable, string endLable){
    string startStr=""; string endStr=""; string s="";
    startStr= startLable + blockName + endLable + " start";
    endStr= startLable + blockName + endLable + " end";
    if( inStr(config, startStr) > 0 && inStr(config, endStr) > 0 ){
        s= strCut(config, startStr, endStr, 1);
        config= replace(config, s, "");
    }
    return config;
}




//获得文件里设置参数　20150315
string getFileParamValue(string configPath, string paramName){
    return handleGetSetFileParameValue(configPath, paramName, "", "获得");
}
//设置文件里设置参数　20150315
string setFileParamValue(string configPath, string paramName, string valueStr){
    return handleGetSetFileParameValue(configPath, paramName, valueStr, "设置");
}
//处理获得设置文件参数值　20150315
string handleGetSetFileParameValue(string configPath, string paramName, string valueStr, string sType){
    string content=""; string startStr=""; string endStr=""; string yunStr=""; string replaceStr="";
    //文件为假时，创建一个空文件看看，如果不能创建这个文件则说明这个文件地址有问题，则退出20150324
    if( checkFile(configPath)== false ){
        createFile(configPath, "");
    }
    string handleGetSetFileParameValue= "";
    if( checkFile(configPath)== false ){ return handleGetSetFileParameValue; }//文件不存在则退出

    content= trimVbCrlf(getFText(configPath));
    startStr= vbCrlf() + paramName + "=" ; endStr= vbCrlf();
    replaceStr= vbCrlf() + paramName + "=" + valueStr + vbCrlf();
    if( inStr(vbCrlf() + content, startStr) > 0 && inStr(content + vbCrlf(), endStr) > 0 ){
        yunStr= strCut(vbCrlf() + content + vbCrlf(), startStr, endStr, 2);
        if( sType== "获得" ){
            handleGetSetFileParameValue= yunStr;
            return handleGetSetFileParameValue;
        }
        yunStr= startStr + yunStr + endStr;
        content= replace(vbCrlf() + content + vbCrlf(), yunStr, replaceStr);
        createFile(configPath, content);
    }else{
        createFile(configPath, content + vbCrlf() + trimVbCrlf(replaceStr));
    }
    return handleGetSetFileParameValue;
}

//设置内容里参数 20150611
string setRParam(string configPath, string paramName, string paramValue, string sIsNoAdd){
    string content=""; string startStr=""; string endStr=""; string s="";
    content= phpTrim(getFText(configPath));
    startStr= paramName + "='" ; endStr= "'";
    if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
        s= strCut(content, startStr, endStr, 2);
        content= replace(content, startStr + s + endStr, startStr + paramValue + endStr);
        createFile(configPath, content);

    }else if( aspTrim(sIsNoAdd)== "1" ){
        createAddFile(configPath, startStr + paramValue + endStr);
    }
    return "";
}

//追加或替换参数值 20150615
string addReplaceRParam( string content, string startStr, string endStr, string valueStr){
    string s="";
    valueStr= startStr + valueStr + endStr;
    if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
        s= strCut(content, startStr, endStr, 1);
        content= replace(content, s, valueStr);
    }else{
        content= content + vbCrlf() + valueStr;
    }
    return content;
}
//删除指定字符N次
string deleteStrCut( string content, string startStr, string endStr, string cutType, int nDelCount){
    int i=0; string s="";
    if( nDelCount== 0 ){
        nDelCount= 99;
    }
    for( i= 0 ; i<= nDelCount; i++){
        s= getStrCut(content, startStr, endStr, 1);
        if( s != "" ){
            content= replace(content, s, "");
        }else{
            break;
        }
    }
    return content;
}



//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容",""))
//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容","追加在前"))
//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容","追加"))
//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容","追加在前"))
//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容","外部追加在前"))
//call Rwend(replaceContentModule(getftext("1.html"), "<div>", "</div>", "替换内容", "外部追加"))
//替换内容里模块   ReplaceType(空为替换，追加在前，追加在后(追加)，外部追加在前，外部追加在后(外部追加))
string replaceContentModule( string content, string startStr, string endStr, string replaceValue, string replaceType){
    string[] splStr; string[] splxx; string s=""; int i=0; string sSplType=""; string valueList=""; string sNewStartStr=""; string sNewEndStr=""; string sourceValueList=""; string sourceValue=""; string tempS=""; string sNewReplaceValue="";
    if( inStr(content, startStr)== 0 && inStr(content, endStr)== 0 ){
        string replaceContentModule= content;
        return replaceContentModule;
    }
    sSplType= "$Array$";
    for( i= 1 ; i<= 99; i++){
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            s= strCut(content, startStr, endStr, 1) ; tempS= s;
            s= mid(s, len(startStr) + 1, len(s) - len(startStr) - len(endStr));
            sNewStartStr= getEachStrAddValue(startStr, "|*|");
            if( inStr(sSplType + valueList + sSplType, sSplType + sNewStartStr + sSplType)== 0 ){
                if( valueList != "" ){ valueList= valueList + sSplType ;}
                valueList= valueList + sNewStartStr;
                if( sourceValueList != "" ){ sourceValueList= sourceValueList + sSplType ;}
                sourceValueList= sourceValueList + startStr;
            }
            sNewEndStr= getEachStrAddValue(endStr, "|*|");
            if( inStr(sSplType + valueList + sSplType, sSplType + sNewEndStr + sSplType)== 0 ){
                if( valueList != "" ){ valueList= valueList + sSplType ;}
                valueList= valueList + sNewEndStr;
                if( sourceValueList != "" ){ sourceValueList= sourceValueList + sSplType ;}
                sourceValueList= sourceValueList + endStr;
            }

            if( replaceType== "追加在前" ){
                sNewReplaceValue= sNewStartStr + replaceValue + s + sNewEndStr;
            }else if( replaceType== "追加在后" || replaceType== "追加" ){
                sNewReplaceValue= sNewStartStr + s + replaceValue + sNewEndStr;
            }else if( replaceType== "外部追加在前" ){
                sNewReplaceValue= replaceValue + sNewStartStr + s + sNewEndStr;

            }else if( replaceType== "外部追加在后" || replaceType== "外部追加" ){
                sNewReplaceValue= sNewStartStr + s + sNewEndStr + replaceValue;
            }else{
                sNewReplaceValue= replaceValue;
            }

            content= replace(content, tempS, sNewReplaceValue);
        }else{
            break;
        }
    }
    //call rwend(content)
    splStr= aspSplit(valueList, sSplType);
    splxx= aspSplit(sourceValueList, sSplType);
    for( i= 0 ; i<= uBound(splStr); i++){
        sourceValue= splStr[i];
        replaceValue= splxx[i];
        content= replace(content, sourceValue, replaceValue);
    }
    return content;
}

//call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "替换内容", "追加在前"))
//call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "替换内容", "追加"))
//call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "替换内容", ""))
//替换内容里一行模块   ReplaceType(空为替换，追加在前，追加在后(追加))
string replaceContentRowModule(string content, string searchValue, string replaceValue, string replaceType){
    string[] splStr; string[] splxx; int i=0; string sSplType=""; string valueList=""; string sourceValueList=""; string sourceValue=""; string sNewReplaceValue=""; string sNewSearchValue="";
    sSplType= "$Array$";
    for( i= 1 ; i<= 99; i++){
        if( inStr(content, searchValue) > 0 ){
            sNewSearchValue= getEachStrAddValue(searchValue, "|*|");
            if( inStr(sSplType + valueList + sSplType, sSplType + sNewSearchValue + sSplType)== 0 ){
                if( valueList != "" ){ valueList= valueList + sSplType ;}
                valueList= valueList + sNewSearchValue;
                if( sourceValueList != "" ){ sourceValueList= sourceValueList + sSplType ;}
                sourceValueList= sourceValueList + searchValue;
            }
            if( replaceType== "追加在前" ){
                sNewReplaceValue= replaceValue + sNewSearchValue;
            }else if( replaceType== "追加在后" || replaceType== "追加" ){
                sNewReplaceValue= sNewSearchValue + replaceValue;
            }else{
                sNewReplaceValue= replaceValue;
            }
            content= replace(content, searchValue, sNewReplaceValue);
        }else{
            break;
        }
    }

    //call rwend(content)
    splStr= aspSplit(valueList, sSplType);
    splxx= aspSplit(sourceValueList, sSplType);
    for( i= 0 ; i<= uBound(splStr); i++){
        sourceValue= splStr[i];
        replaceValue= splxx[i];
        content= replace(content, sourceValue, replaceValue);
    }
    string replaceContentRowModule= content;

    return replaceContentRowModule;
}
//处理配置文档(20150804)
void handleConfigFile(string configPath){
    string c="";
    if( checkFile(configPath)== false ){
        c= "#Help帮助# start" + vbCrlf() + "默认帮助内容" + vbCrlf() + "#Help帮助# end";
        createFile(configPath, c);
    }
}

//获得内容里指定类型值   RParam加强版(20161025)
string getRParam( string content, string lableStr){
    string contentLCase=""; string endS=""; int i=0; string s=""; string c=""; bool isStart; string startStr=""; bool isValue;
    content= " " + content + " "; //避免更精准获得值
    contentLCase= lCase(content);
    lableStr= lCase(lableStr);
    endS= mid(content, inStr(contentLCase, lableStr) + len(lableStr),-1);
    //call echo("ends",ends)
    isStart= false; //是否有开始类型值
    isValue= false; //是否有值
    for( i= 1 ; i<= len(endS); i++){
        s= mid(endS, i, 1);
        if( isStart== true ){
            if( s != "" ){
                if( startStr== "" ){
                    startStr= s;
                }else{
                    if( startStr== "\"" || startStr== "'" ){
                        if( s== startStr ){
                            isValue= true;
                            break;
                        }
                    }else if( s== " " && c== "" ){

                    }else if( s== " " || s== "/" || s== ">" ){
                        isValue= true;
                        break;
                    }
                    if( s != " " ){
                        c= c + s;
                    }
                }
            }
        }

        if( s== "=" ){
            isStart= true;
        }
    }
    if( isValue== false ){
        c= "";
    }
    string getRParam= c;
    //call echo("c",c)
    return getRParam;
}

//获得模板某标签默认内容 代码进行了二次查找 会在HTML模板里二次查找默认值
string getDefaultValue(string action){
    return moduleFindContent(action, "default");
}

//添加模块替换数组


//根据标签找到对应内容
string moduleFindContent(string action, string moduleName){
    string defaultStr=""; string startStr=""; string endStr="";
    defaultStr= rParam(action, moduleName); //把转小写LCase去掉 （20151008）

    startStr= "<!--#" + defaultStr + " start#-->";
    endStr= "<!--#" + defaultStr + " end#-->";
    //[_18年独家一次性祛斑第一品牌2014年10月21日 10时59分]
    //Call Echo("Default",Default)
    //判断是否存在
    if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
        defaultStr= getStrCut(code, startStr, endStr, 2);
    }else if( defaultStr != "" ){
        startStr= "<!--#" + defaultStr;
        endStr= "#-->";
        if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
            defaultStr= getStrCut(code, startStr, endStr, 2);
        }
    }

    //删除默认值20150712
    string deletedefault="";
    deletedefault= rParam(action, "deletedefault");
    if( deletedefault== "true" ){
        addModuleReplaceArray("【删除】", startStr + defaultStr + endStr);
    }
    return defaultStr;
}
</script>
