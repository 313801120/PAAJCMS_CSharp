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


//http://127.0.0.1/����/ClassAspCode.Asp?act=GetFileFunctionStrList    ������������������



//**************************************** ��php�� ͨ�� ****************************************

//�滻����ֵ 2014  12 01
string newReplaceValueParam(string content, string paramName, string replaceStr){
    string startStr=""; string endStr=""; string labelStr=""; string sLen=""; string sNTimeFormat=""; string delHtmlYes=""; string trimYes="";
    //ReplaceStr = ReplaceStr & "�����������������ʱ̼ѽ��"
    //ReplaceStr = CStr(ReplaceStr)            'ת���ַ�����
    if( isNul(replaceStr)== true ){ replaceStr= "" ;}

    startStr= "[$" + paramName ; endStr= "$]";
    if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
        labelStr= strCut(content, startStr, endStr, 1);
        //ɾ��Html
        delHtmlYes= rParam(labelStr, "DelHtml"); //�Ƿ�ɾ��Html
        if( delHtmlYes== "true" ){ replaceStr= replace(delHtml(replaceStr), "<", "&lt;") ;}//HTML����
        //ɾ�����߿ո�
        trimYes= rParam(labelStr, "Trim"); //�Ƿ�ɾ�����߿ո�
        if( trimYes== "true" ){ replaceStr= trimVbCrlf(replaceStr) ;}

        //��ȡ�ַ�����
        sLen= handleNumber(rParam(labelStr, "Len"));
        //If nLen<>"" Then ReplaceStr = CutStr(ReplaceStr,nLen,"null")' Left(ReplaceStr,nLen)
        if( sLen != "" ){ replaceStr= cutStr(replaceStr, cInt(sLen), "...") ;}//Left(ReplaceStr,nLen)


        //ʱ�䴦��
        sNTimeFormat= rParam(labelStr, "Format_Time"); //ʱ�䴦��ֵ
        if( sNTimeFormat != "" ){
            replaceStr= format_Time(replaceStr, cInt(sNTimeFormat));
        }
        content= replace(content, labelStr, replaceStr);

    }
    return content;
}

//���ݱ�ǩ�ҵ���Ӧ����
string newRParam(string dataCode, string action, string moduleName){
    string defaultStr=""; string startStr=""; string endStr="";
    defaultStr= rParam(action, moduleName);
    startStr= "<!--#" + defaultStr + " start#-->";
    endStr= "<!--#" + defaultStr + " end#-->";

    if( defaultStr != "" ){
        //�ж��Ƿ����
        if( inStr(dataCode, startStr) > 0 && inStr(dataCode, endStr) > 0 ){
            defaultStr= strCut(dataCode, startStr, endStr, 2);
        }else{
            startStr= "<!--#" + defaultStr;
            endStr= "#-->";
            if( inStr(dataCode, startStr) > 0 && inStr(dataCode, endStr) > 0 ){
                defaultStr= strCut(dataCode, startStr, endStr, 2);

                //Call Echo("��","StartStr=" & StartStr & ",EndStr=" & EndStr  & ",Default=" & Default)
            }
        }
    }
    return defaultStr;
}




//**************************************** ��php�� ͨ�� **************************************** end


//����ȫ������(20150827)
string getContentAllRunStr(string content){
    string[] splStr; string s=""; string c=""; string tempS="";
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            tempS= s;
            s= replace(replace(s, chr(10).ToString(), ""), chr(13).ToString(), ""); //���Ϊʲô s����� chr(10)��chr(13) �أ�
            c= c + tempS + "=" + handleContentCode(s, "") + vbCrlf();
        }
    }
    return c;
}

//����������к��ַ�  �÷� len(aaaa)  û˫������
string getContentRunStr( string content){
    return handleContentCode(content, "");
}
//������������""�������ɾ����20150329
//����������к��ַ�
string checkContentRunStr( string content){
    return handleContentCode(content, "check");
}
//����˫����
string handleDoubleQuotation( string s){
    string sNew="";
    sNew= phpTrim(s);
    if( left(sNew, 1)== "\"" && right(sNew, 1)== "\"" ){
        s= mid(sNew, 2, len(sNew) - 2);
    }
    return s;
}
//��������
string strDQ( string s){
    return handleDoubleQuotation(s);
}
//��������� 20150330
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
//������������� ����20150324   http://127.0.0.1/����/ClassAspCode.Asp?act=GetFileFunctionStrList  �������    version 1.0
string handleContentCode(string content,string sType){return "";return "";//���պ���
    return "";
}


//�ڲ�ģ�鴦�� HandleInModule(Content,"start") HandleInModule(Content,"end")
string handleInModule(string content, string sType){
    sType= lCase(cStr(sType));
    if( sType== "1" || sType== "start" ){
        content= replace(content, "\\'", "\\|*|\\");
        content= replace(content, "\\=", "\\|&|\\"); //���20141024
    }else if( sType== "2" || sType== "end" ){
        content= replace(content, "\\|*|\\", "'");
        content= replace(content, "\\$", "$");
        content= replace(content, "\\}", "}");

        content= replace(content, "\\|&|\\", "="); //���20141024
    }
    return content;
}
//���������ʽ���ñ�ǩֵ
string clearRParam( string action, string lableStr){
    string s="";
    //Action=Replace(Action,"\'","��|\��|��")
    action= replace(action, "\\'", ""); //���������
    s= rParam(action, lableStr);
    //s=replace(s,"��|\��|��", "\'")
    return s;
}
//��ò������ݺ� �ŵ������ﴦ��һ�£�20151023��
string atRParam( string action, string lableStr){
    string atRParam= rParam(action, lableStr);
    if( inStr(atRParam, "{$") > 0 && inStr(atRParam, "$}") > 0 ){
        atRParam= handleTemplateAction(atRParam, false); //������  ע������֮ǰ�ɰ汾
        atRParam= handleAction(atRParam); //�����°汾
    }
    return atRParam;
}
//����������ֵ  Title = RParam(Action,"Title")     ��ǿ���ȡ����ֵ20150723
string rParam( string action, string lableStr){
    string s="";

    //ԭʼ ������
    s= handleRParam(action, lableStr, "'");
    //ԭʼ ˫����
    if( s== "" ){
        s= handleRParam(action, lableStr, "\"");
    }
    //ԭʼ ��
    if( s== "" ){
        s= handleRParam(action, lableStr, "");
    }

    //Сд ������
    if( s== "" ){
        s= handleRParam(action, lCase(lableStr), "'");
    }
    //Сд ˫����
    if( s== "" ){
        s= handleRParam(action, lCase(lableStr), "\"");
    }
    //Сд ��
    if( s== "" ){
        s= handleRParam(action, lCase(lableStr), "");
    }

    //��д ������
    if( s== "" ){
        s= handleRParam(action, uCase(lableStr), "'");
    }
    //��д ˫����
    if( s== "" ){
        s= handleRParam(action, uCase(lableStr), "\"");
    }
    //��д ��
    if( s== "" ){
        s= handleRParam(action, uCase(lableStr), "");
    }
    //��Ҫ�����Ҫ�����ȶ�(20151022)
    //if s=false then s=""
    if( s== "[#��*ֵ_#]" ){ s= "" ;}
    return s;
}
//���� ����������ֵ
string handleRParam( string action, string lableStr, string typeStr){
    string lalbeName=""; string endTypeStr=""; bool isTrue; string s="";
    isTrue= false; //�Ƿ�Ϊ��
    endTypeStr= IIF(typeStr != "", typeStr, " ");
    action= vbCrlf() + " " + action; //����Ҳ�Ӹ��ո�Ҫ��Ȼ��û�к�����ǰ���û�пո�
    //Ĭ��ǰ��ӿո�
    lalbeName= " " + lableStr; //�Ӹ��ո���Ϊ�˾�׼
    //������  ǰ��ӵ�
    if( inStr(action, lalbeName + "=" + typeStr)== 0 && isTrue== false ){
        lalbeName= "'" + lableStr;
    }else{
        isTrue= true;
    }
    //������ ǰ���˫����
    if( inStr(action, lalbeName + "=" + typeStr)== 0 && isTrue== false ){
        lalbeName= "\"" + lableStr;
    }else{
        isTrue= true;
    }
    //������    ǰ���TAB
    if( inStr(action, lalbeName + "=" + typeStr)== 0 && isTrue== false ){
        lalbeName= vbTab() + lableStr;
    }else{
        isTrue= true;
    }
    //������    ǰ��ӻ���
    if( inStr(action, lalbeName + "=" + typeStr)== 0 && isTrue== false ){
        lalbeName= vbCrlf() + lableStr;
    }else{
        isTrue= true;
    }
    if( inStr(action, lalbeName + "=" + typeStr) > 0 && inStr(action, endTypeStr) > 0 ){
        s= strCut(action, lalbeName + "=" + typeStr, endTypeStr, 2);
        s= handleInModule(s, "end"); //����������� ׷����20141031            ��ԭ����ֵ

        if( s== "" ){
            s= "[#��*ֵ_#]";
        }

        //�ж��Ƿ�Բ������ж�������
        if( inStr(s, "{$") > 0 && inStr(s, "$}") > 0 ){

            //handleRParam=HandleTemplateAction(handleRParam,true)        '������
            //handleRParam = handleModuleReplaceArray(handleRParam)'��AddSqL����һ�¶��� ���Ǵ����滻������Ҫ����Ϊ�� HandleTemplateAction���滻��(20151021)

        }
        //��Ҫ�����Ҫ�����ȶ�(20151022)
        //if handleRParam="" then
        //handleRParam=false
        //end if
    }
    return s;
}


//������ÿ� 20150105 GetConfigBlock(ConfigContent, BlockName)
string getConfigBlock(string configContent, string blockName){
    return getCutConfigBlock(configContent, blockName, "#", "#");
}
//������ÿ� 20150105
string getConfigBlock2(string configContent, string blockName){
    return getCutConfigBlock(configContent, blockName, "[#", "#]");
}
//������ÿ� 20150105
string getConfigBlock3(string configContent, string blockName){
    return getCutConfigBlock(configContent, blockName, "[$", "$]");
}
//��ȡ���������п� 20150105
string getCutConfigBlock(string configContent, string blockName, string startLable, string endLable){
    string startStr=""; string endStr="";
    startStr= startLable + blockName + endLable;
    endStr= startLable + blockName + endLable;
    string getCutConfigBlock= "";
    //��ʼ��ǩ����
    if( inStr(configContent, startStr + " start") > 0 ){
        startStr= startStr + " start";
    }else{
        startStr= startStr + " Start";
    }
    //������ǩ����
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
//����������ݿ�20150401
string getConfigContentBlock( string configContent, string blockName){
    return getCutConfigBlock(configContent, blockName, "", "");
}
//��������ļ����20150401  getConfigFileBlock(ConfigPath, "#txtRunCode#")  ���Ա�ǩ��ʱ���Զ�����
string getConfigFileBlock( string configFile, string blockName){
    string content=""; string findStr=""; string replaceStr=""; string startStr=""; string endStr="";
    content= getFText(configFile);
    string getConfigFileBlock= "";
    //MsgBox ("ConfigFile=" & ConfigFile & "(" & checkFile(ConfigFile) & "��" & GetFSize(ConfigFile) & ")" & vbCrLf & "Content=" & Content)
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
//���������ļ���� 20150401 call setConfigFileBlock(ConfigFile, "aaabbc", "#�ϴ�Ŀ¼�б�#")  ���������
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

//ɾ�����ÿ� 20150322
string delConfigBlock(string config, string blockName){
    return delCutConfigBlock(config, blockName, "#", "#");
}
//ɾ�����ÿ� 20150322
string delConfigBlock2(string config, string blockName){
    return delCutConfigBlock(config, blockName, "[#", "#]");
}

//ɾ�����ÿ� 20150322
string delConfigBlock3(string config, string blockName){
    return delCutConfigBlock(config, blockName, "[$", "$]");
}
//ɾ���������� 20150322
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




//����ļ������ò�����20150315
string getFileParamValue(string configPath, string paramName){
    return handleGetSetFileParameValue(configPath, paramName, "", "���");
}
//�����ļ������ò�����20150315
string setFileParamValue(string configPath, string paramName, string valueStr){
    return handleGetSetFileParameValue(configPath, paramName, valueStr, "����");
}
//�����������ļ�����ֵ��20150315
string handleGetSetFileParameValue(string configPath, string paramName, string valueStr, string sType){
    string content=""; string startStr=""; string endStr=""; string yunStr=""; string replaceStr="";
    //�ļ�Ϊ��ʱ������һ�����ļ�������������ܴ�������ļ���˵������ļ���ַ�����⣬���˳�20150324
    if( checkFile(configPath)== false ){
        createFile(configPath, "");
    }
    string handleGetSetFileParameValue= "";
    if( checkFile(configPath)== false ){ return handleGetSetFileParameValue; }//�ļ����������˳�

    content= trimVbCrlf(getFText(configPath));
    startStr= vbCrlf() + paramName + "=" ; endStr= vbCrlf();
    replaceStr= vbCrlf() + paramName + "=" + valueStr + vbCrlf();
    if( inStr(vbCrlf() + content, startStr) > 0 && inStr(content + vbCrlf(), endStr) > 0 ){
        yunStr= strCut(vbCrlf() + content + vbCrlf(), startStr, endStr, 2);
        if( sType== "���" ){
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

//������������� 20150611
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

//׷�ӻ��滻����ֵ 20150615
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
//ɾ��ָ���ַ�N��
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



//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","�滻����",""))
//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","�滻����","׷����ǰ"))
//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","�滻����","׷��"))
//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","�滻����","׷����ǰ"))
//call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","�滻����","�ⲿ׷����ǰ"))
//call Rwend(replaceContentModule(getftext("1.html"), "<div>", "</div>", "�滻����", "�ⲿ׷��"))
//�滻������ģ��   ReplaceType(��Ϊ�滻��׷����ǰ��׷���ں�(׷��)���ⲿ׷����ǰ���ⲿ׷���ں�(�ⲿ׷��))
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

            if( replaceType== "׷����ǰ" ){
                sNewReplaceValue= sNewStartStr + replaceValue + s + sNewEndStr;
            }else if( replaceType== "׷���ں�" || replaceType== "׷��" ){
                sNewReplaceValue= sNewStartStr + s + replaceValue + sNewEndStr;
            }else if( replaceType== "�ⲿ׷����ǰ" ){
                sNewReplaceValue= replaceValue + sNewStartStr + s + sNewEndStr;

            }else if( replaceType== "�ⲿ׷���ں�" || replaceType== "�ⲿ׷��" ){
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

//call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "�滻����", "׷����ǰ"))
//call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "�滻����", "׷��"))
//call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "�滻����", ""))
//�滻������һ��ģ��   ReplaceType(��Ϊ�滻��׷����ǰ��׷���ں�(׷��))
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
            if( replaceType== "׷����ǰ" ){
                sNewReplaceValue= replaceValue + sNewSearchValue;
            }else if( replaceType== "׷���ں�" || replaceType== "׷��" ){
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
//���������ĵ�(20150804)
void handleConfigFile(string configPath){
    string c="";
    if( checkFile(configPath)== false ){
        c= "#Help����# start" + vbCrlf() + "Ĭ�ϰ�������" + vbCrlf() + "#Help����# end";
        createFile(configPath, c);
    }
}

//���������ָ������ֵ   RParam��ǿ��(20161025)
string getRParam( string content, string lableStr){
    string contentLCase=""; string endS=""; int i=0; string s=""; string c=""; bool isStart; string startStr=""; bool isValue;
    content= " " + content + " "; //�������׼���ֵ
    contentLCase= lCase(content);
    lableStr= lCase(lableStr);
    endS= mid(content, inStr(contentLCase, lableStr) + len(lableStr),-1);
    //call echo("ends",ends)
    isStart= false; //�Ƿ��п�ʼ����ֵ
    isValue= false; //�Ƿ���ֵ
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

//���ģ��ĳ��ǩĬ������ ��������˶��β��� ����HTMLģ������β���Ĭ��ֵ
string getDefaultValue(string action){
    return moduleFindContent(action, "default");
}

//���ģ���滻����


//���ݱ�ǩ�ҵ���Ӧ����
string moduleFindContent(string action, string moduleName){
    string defaultStr=""; string startStr=""; string endStr="";
    defaultStr= rParam(action, moduleName); //��תСдLCaseȥ�� ��20151008��

    startStr= "<!--#" + defaultStr + " start#-->";
    endStr= "<!--#" + defaultStr + " end#-->";
    //[_18�����һ������ߵ�һƷ��2014��10��21�� 10ʱ59��]
    //Call Echo("Default",Default)
    //�ж��Ƿ����
    if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
        defaultStr= getStrCut(code, startStr, endStr, 2);
    }else if( defaultStr != "" ){
        startStr= "<!--#" + defaultStr;
        endStr= "#-->";
        if( inStr(code, startStr) > 0 && inStr(code, endStr) > 0 ){
            defaultStr= getStrCut(code, startStr, endStr, 2);
        }
    }

    //ɾ��Ĭ��ֵ20150712
    string deletedefault="";
    deletedefault= rParam(action, "deletedefault");
    if( deletedefault== "true" ){
        addModuleReplaceArray("��ɾ����", startStr + defaultStr + endStr);
    }
    return defaultStr;
}
</script>
