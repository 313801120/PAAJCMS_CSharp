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
//ɾ�����ֶ����ǩ <R#��������BlockName��վ���� start#>  <R#��������BlockName��վ���� end#>
//��ģ�崦��



//��ģ������
string XY_ReadTemplateModule(string action){
    string moduleId=""; string filePath=""; string c=""; int i=0;
    string sourceList=""; //Դ�����б� 20150109
    string replaceList=""; //�滻�����б�
    string[] splSource; string[] splReplace; string sourceStr=""; string replaceStr="";
    filePath= rParam(action, "File");
    moduleId= rParam(action, "ModuleId");
    sourceList= rParam(action, "SourceList");
    replaceList= rParam(action, "ReplaceList");
    //Call Echo(SourceList,ReplaceList)

    if( moduleId== "" ){ moduleId= rParam(action, "ModuleName") ;}//�ÿ�����
    filePath= filePath + ".html";
    //Call Echo("FilePath",FilePath)
    //Call Echo("ModuleId",ModuleId)
    c= readTemplateModuleStr(filePath, "", moduleId);
    //���滻��20160331
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

//��ģ������
string readTemplateModuleStr(string filePath, string defaultContent, string moduleId){
    string startStr=""; string endStr=""; string content="";
    string readTemplateModuleStr="";
    startStr= "<!--#Module " + moduleId + " start#-->";
    endStr= "<!--#Module " + moduleId + " end#-->";
    //FilePath = ReplaceGlobleLable(FilePath)                '�滻ȫ����ǩ        '�����2014 12 11

    //�ļ������ڣ���׷��ģ��·�� 20150616 ��VB�������
    if( checkFile(filePath)== false ){
        filePath= cfg_webTemplate + "/" + filePath;
    }

    //û��handleRGV���������
    //filePath = handleRGV(filePath, "[$ģ��Ŀ¼$]", "Module/")                              'Module

    if( defaultContent != "" ){
        content= defaultContent;
    }else if( checkFile(filePath)== true ){
        content= getFText(filePath);
    }else{
        content= code; //Ĭ��������ָ������
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
        readTemplateModuleStr= "ģ��[" + moduleId + "]������,·��=" + filePath;
        die(readTemplateModuleStr + content);
    }
    return readTemplateModuleStr;
}
//��ģ���Ӧ����
string findModuleStr(string content, string valueStr){
    string startStr=""; string endStr=""; string yuanStr=""; string replaceStr=""; int i=0; string block=""; string blockFile=""; string action="";
    for( i= 1 ; i<= 9; i++){
        startStr= "[$�������� " ; endStr= "$]";
        if( inStr(valueStr, startStr) > 0 && inStr(valueStr, endStr) > 0 ){
            action= strCut(valueStr, startStr, endStr, 2);
            block= rParam(action, "Block");
            blockFile= rParam(action, "File");
            if( inStr(vbCrlf() + readBlockList + vbCrlf(), vbCrlf() + block + vbCrlf())== 0 ){
                readBlockList= readBlockList + block + vbCrlf();
            }
            //���ļ����� ���������
            if( blockFile != "" ){
                content= getFText(blockFile);
            }
            yuanStr= startStr + action + endStr;
            replaceStr= "";

            startStr= "<R#��������" + block + " start#>" ; endStr= "<R#��������" + block + " end#>";
            if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
                replaceStr= strCut(content, startStr, endStr, 2);
            }else{
                startStr= "<!--#��������" + block ; endStr= "#-->";
                if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
                    replaceStr= strCut(content, startStr, endStr, 2);
                }
            }
            //Call Eerr(YuanStr,ReplaceStr)
            valueStr= replace(valueStr, yuanStr, replaceStr);
            //Call Echo("ValueStr",ValueStr)
        }else{
            //û��ģ��Ҫ������ ���˳�
            break;
        }
    }
    return valueStr;
}

//����Leftģ����ʽ        �������ַ� ' ���ظ��ƻ�������������� \|*|\ ���洦��
string XY_ReadColumeSetTitle(string action){
    string startStr=""; string endStr=""; string style=""; string title=""; string valueStr=""; string moreClass=""; string moreUrl=""; string moreStr=""; string aStr=""; string c="";
    action= handleInModule(action, "start");
    style= rParam(action, "style");
    title= rParam(action, "Title");
    //Call Echo("ContentHeight",ContentHeight)
    //ValueStr = RParam(Action,"value")
    //����ģ��������
    valueStr= moduleFindContent(action, "value");
    //Call Eerr("ValueStr",ValueStr)
    valueStr= findModuleStr(code, valueStr); //��ģ���Ӧ����

    moreClass= rParam(action, "MoreClass");
    moreUrl= phpTrim(rParam(action, "MoreUrl"));
    moreStr= rParam(action, "MoreStr");
    valueStr= handleInModule(valueStr, "end");
    c= readColumeSetTitle(action, style, title, valueStr);

    if( moreClass== "" ){ moreClass= "more" ;}//More����Ϊ�� ����Ĭ�ϴ���
    //If MoreUrl="" Then MoreUrl="#"                    'More������ַΪ�� ����Ĭ��#����
    //More������ʽ����Ϊ�գ���Ϊû����ʽ���Ͳ�����More�������
    if( moreUrl != "" && moreStr != "" ){
        //AStr = "<a href='"& MoreUrl &"' class='"& MoreClass &"'>"& MoreStr &"</a>"
        aStr= "<a " + aHref(moreUrl, title, "") + " class='" + moreClass + "'>" + moreStr + "</a>";
        c= replace(c, "<!--#AMore#-->", aStr);
    }
    return c;
}

//����Ŀ��������������ֵ
string readColumeSetTitle(string action, string styleID, string columeTitle, string columeContent){
    string titleWidth=""; //������
    string titleHeight=""; //����߶�
    string contentHeight=""; //���ݸ߶�
    string contentWidth=""; //���ݿ��
    string contentCss="";

    titleWidth= rParam(action, "TitleWidth"); //��ñ���߶�    ��Ӧ��20150715
    titleHeight= rParam(action, "TitleHeight"); //��ñ�����
    contentWidth= rParam(action, "ContentWidth"); //������ݿ��
    contentHeight= rParam(action, "ContentHeight"); //������ݸ߶�

    //�����
    titleWidth= aspTrim(titleWidth);
    //�Զ���px��λ�����ӻ���Ч�� 20150115
    if( right(titleHeight, 1) != "%" && right(titleHeight, 2) != "px" && titleHeight != "" && titleHeight != "auto" ){
        titleHeight= titleHeight + "px";
    }
    if( right(titleWidth, 1) != "%" && right(titleWidth, 2) != "px" && titleWidth != "" && titleWidth != "auto" ){
        titleWidth= titleWidth + "px";
    }
    //���ݸ�
    contentHeight= aspTrim(contentHeight);
    //�Զ���px��λ�����ӻ���Ч�� 20150115
    if( right(contentHeight, 1) != "%" && right(contentHeight, 2) != "px" && contentHeight != "" && contentHeight != "auto" ){
        contentHeight= contentHeight + "px";
    }
    //���ݿ�
    contentWidth= aspTrim(contentWidth);
    //�Զ���px��λ�����ӻ���Ч�� 20150115
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
    //�����
    if( titleWidth != "" ){
        content= replace(content, "<div class=\"tvalue\">", "<div class=\"tvalue\" style='width:" + titleWidth + ";'>");
    }
    //���ݸ�
    if( contentCss != "" ){
        content= replace(content, "<div class=\"ccontent\">", "<div class=\"ccontent\" style='" + contentCss + "'>");
    }
    //call echo(ContentWidth,ContentCss)

    content= replace(content, "��Ŀ����", columeTitle);
    content= replace(content, "��Ŀ����", columeContent);
    return content;
}

//����Ŀģ��
string readColumn(string sID){
    string templateFilePath=""; string startStr=""; string endStr=""; string s="";
    templateFilePath= cfg_webTemplate + "\\Template_Left.html";
    startStr= "/*columnlist" + sID + "Start*/";
    endStr= "/*columnlist" + sID + "End*/";
    s= readTemplateFileModular(templateFilePath, startStr, endStr);
    if( s== "[$NO$]" ){
        s= "Left��ʽID[" + sID + "]������";
    }
    return s;
}


//��ģ���ز�
string readTemplateSource(string sID){
    string templateFilePath=""; string startStr=""; string endStr=""; string s="";
    templateFilePath= cfg_webTemplate + "\\TemplateSource.html";
    startStr= "<!--#sourceHtml" + sID + "Start#-->";
    endStr= "<!--#sourceHtml" + sID + "End#-->";
    s= readTemplateFileModular(templateFilePath, startStr, endStr);
    if( s== "[$NO$]" ){
        s= "ģ����ԴID[" + sID + "]������";
    }
    return s;
}



//��ģ���ļ���ĳģ��
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

//���ļ�ģ���ز�
string readTemplateFileSource(string templateFilePath, string sID){
    string startStr=""; string endStr=""; string s=""; string c="";
    string readTemplateFileSource="";
    startStr= "<!--#sourceHtml" + replace(sID, ".html", "") + "Start#-->";
    endStr= "<!--#sourceHtml" + replace(sID, ".html", "") + "End#-->";
    s= readTemplateFileModular(templateFilePath, startStr, endStr);
    if( s== "[$NO$]" ){
        //��һ����ȡ���ı��������б�(20150815)
        c= getStrCut(pubCode, startStr, endStr, 2);
        if( c != "" ){
            readTemplateFileSource= c;
            //call rwend(c)
            return readTemplateFileSource;
        }
        c= getFText(templateFilePath);
        //���� <!--#TemplateSplitStart#-->  �ͷ��ص�ǰȫ������
        if( inStr(c, "<!--#DialogStart#-->") > 0 ){
            readTemplateFileSource= c;
            return readTemplateFileSource;
        }

        s= "ģ����ԴID[" + sID + "]������,·��TemplateFilePath=" + handlePath(templateFilePath);
    }
    return s;
}
//�����ļ�չʾ�б���Դ
string readArticleListStyleSource(string sID){
    string filePath="";
    filePath= cfg_webImages + "\\����չʾ��ʽ\\" + sID;
    if( checkFile(filePath)== false ){
        filePath= cfg_webTemplate + "\\Resources\\" + sID;
    }
    //call echo(checkfile(filePath),filePath)
    string readArticleListStyleSource= readTemplateFileSource(filePath, sID);

    return readArticleListStyleSource;
}
//�����ļ���Ϣ�б���Դ
string readArticleInfoStyleSource(string sID){
    string filePath="";
    filePath= cfg_webImages + "\\������Ϣչʾ��ʽ\\" + sID;
    if( checkFile(filePath)== false ){
        filePath= cfg_webTemplate + "\\Resources\\" + sID;
    }
    return readTemplateFileSource(filePath, sID);
}


</script>
