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
//��̬����̬(2013,12,17)

//================ ���ٻ����վ���� ==================
//�����޸� �޸ĵ��ı�
//MainStr = DisplayOnlineED2(WEB_ADMINURL &"MainInfo.Asp?act=ShowEdit&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainStr, "<li|<a ")
//�����޸� ��Ʒ����
//DidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditBigClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), DidStr, "<li|<a ")
//�����޸� ��ƷС��
//SidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), SidStr, "<li|<a ")
//�����޸� ��Ʒ����
//S = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditThreeClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), S, "<li|<a ")
//�����޸�  ����
//ProStr = DisplayOnlineED2(WEB_ADMINURL &"Product.Asp?act=ShowEditProduct&Id=" & TempRs("Id") & "&n=" & GetRnd(11), ProStr, "<li|<a ")
//�����޸� ��������
//NavDidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavDidStr, "<li|<a ")
//�����޸� ����С��
//NavSidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavSmall&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavSidStr, "<li|<a ")

//-------------------------------- ����Ϊ��վ��̨���ÿ�ݱ�ǩ������ -------------------------------------------

//����������ɫ
string infoColor(string str, string color){
    if( color != "" ){ str= "<font color=" + color + ">" + str + "</font>" ;}
    return str;
}
//ͼƬ����ʧ����ʾĬ��ͼƬ
string imgError(){
    return " onerror=\"this.src='/UploadFiles/NoImg.jpg'\"";
}
//���target��ʽ
string handleTargetStr( string sType){
    string handleTargetStr="";
    if( sType != "" ){
        handleTargetStr= " target='" + sType + "'";
    }
    return handleTargetStr;
}
//�򿪷�ʽ  (����)
string aTarget(string sType){
    return handleTargetStr(sType);
}
//�������Title��ʽ
string aTitle( string title){
    string aTitle="";
    if( title != "" ){
        aTitle= " Title='" + title + "'";
    }
    return aTitle;
}
//�������Title
string imgAlt( string alt){
    string imgAlt="";
    if( alt != "" ){
        imgAlt= " alt='" + alt + "'";
    }
    return imgAlt;
}
//ͼƬ������Alt
string imgTitleAlt( string str){
    string imgTitleAlt="";
    if( str != "" ){
        imgTitleAlt= " alt='" + str + "' title='" + str + "'";
    }
    return imgTitleAlt;
}
//���A Relֵ
string aRel( bool isType){
    string aRel="";
    if( isType== true ){
        aRel= " rel='nofollow'";
    }
    return aRel;
}
//���target��ʽ
string styleClass( string className){
    string styleClass="";
    if( className != "" ){
        styleClass= " class='" + className + "'";
    }
    return styleClass;
}
//�ı��Ӵ�
string textFontB( string text, bool isFontB){
    if( isFontB== true ){
        text= "<strong>" + text + "</strong>";
    }
    return text;
}
//�ı�����ɫ
string textFontColor( string text, string color){
    if( color != "" ){
        text= "<font color='" + color + "'>" + text + "</font>";
    }
    return text;
}
//�����ı���ɫ��Ӵ�
string fontColorFontB(string title, bool isFontB, string fontColor){
    return textFontColor(textFontB(title, isFontB), fontColor);
}
//���Ĭ��������Ϣ�ļ�����
string getDefaultFileName(){
    return format_Time(now(), 6);
}
//�������  ����'"<a " & AHref(Url, TempRs("BigClassName"), TempRs("Target")) & ">" & TempRs("BigClassName") & "</a>"
string aHref(string url, string title, string target){
    url= handleHttpUrl(url); //����һ��URL ��֮����
    return "href='" + url + "'" + aTitle(title) + aTarget(target);
}
//���ͼƬ·��
string imgSrc(string url, string title, string target){
    url= handleHttpUrl(url); //����һ��URL ��֮����
    return "src='" + url + "'" + aTitle(title) + imgAlt(title) + aTarget(target);
}

//============== ��վ��̨ʹ�� ==================

//ѡ��Target�򿪷�ʽ
string selectTarget(string target){
    string c=""; string sel="";
    c= c + "<select name=\"Target\" id=\"Target\">" + vbCrlf();
    c= c + "  <option value=''>���Ӵ򿪷�ʽ</option>" + vbCrlf();
    if( target== "" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "  <option" + sel + " value=''>��ҳ��</option>" + vbCrlf();
    if( target== "_blank" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "  <option value=\"_blank\"" + sel + ">��ҳ��</option>" + vbCrlf();
    if( target== "Index" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "  <option value=\"Index\"" + sel + ">Indexҳ��</option>" + vbCrlf();
    if( target== "Main" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "  <option value=\"Main\"" + sel + ">Mainҳ��</option>" + vbCrlf();
    c= c + "</select>" + vbCrlf();
    return c;
}
//ѡ���ı���ɫ
string selectFontColor(string fontColor){
    string c=""; string sel="";
    c= c + "  <select name=\"FontColor\" id=\"FontColor\">" + vbCrlf();
    c= c + "    <option value=''>�ı���ɫ</option>" + vbCrlf();
    if( fontColor== "Red" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"Red\" class=\"FontColor_Red\"" + sel + ">��ɫ</option>" + vbCrlf();
    if( fontColor== "Blue" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"Blue\" class=\"FontColor_Blue\"" + sel + ">��ɫ</option>" + vbCrlf();
    if( fontColor== "Green" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"Green\" class=\"FontColor_Green\"" + sel + ">��ɫ</option>" + vbCrlf();
    if( fontColor== "Black" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"Black\" class=\"FontColor_Black\"" + sel + ">��ɫ</option>" + vbCrlf();
    if( fontColor== "White" ){ sel= " selected" ;}else{ sel= "" ;}
    c= c + "    <option value=\"White\" class=\"FontColor_White\"" + sel + ">��ɫ</option>" + vbCrlf();
    c= c + "  </select>" + vbCrlf();
    return c;
}
//ѡ����Ů
string selectSex(string sex){
    string c=""; string sel="";
    c= c + "  <select name=\"FontColor\" id=\"FontColor\">" + vbCrlf();
    c= c + "    <option value=\"��\">��</option>" + vbCrlf();
    sel= IIF(sex== "Ů", " selected", "");
    c= c + "    <option value=\"Ů\"" + sel + ">Ů</option>" + vbCrlf();
    c= c + "  </select>" + vbCrlf();
    return c;
}
//ѡ��Session��Cookies��֤
string selectSessionCookies(string verificationMode){
    string c=""; string sel="";
    c= c + "  <select name=\"VerificationMode\" id=\"VerificationMode\">" + vbCrlf();
    c= c + "    <option value=\"1\">Session��֤</option>" + vbCrlf();
    sel= IIF(verificationMode== "0", " selected", "");
    c= c + "    <option value=\"0\"" + sel + ">Cookies��֤</option>" + vbCrlf();
    c= c + "  </select>" + vbCrlf();
    return c;
}
//��ʾѡ��ָ�����  showSelectList("aa","aa|bb|cc","|","bb")
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

//��ʾ����չʾ�б���ʽ 20150114   �� Call Rw(ShowArticleListStyle("�����б��.html"))
string showArticleListStyle( string thisValue){
    return handleArticleListStyleOrInfoStyle("����չʾ��ʽ", "ArticleListStyle", thisValue);
}
//��ʾ������Ϣչʾ��ʽ 20150114   �� Call Rw(ShowArticleInfoStyle("�����б��.html"))
string showArticleInfoStyle( string thisValue){
    return handleArticleListStyleOrInfoStyle("������Ϣչʾ��ʽ", "ArticleInfoStyle", thisValue);
}
//��������չʾ�б���ʽ��������Ϣ��ʽ
string handleArticleListStyleOrInfoStyle(string folderName, string inputName, string thisValue){
    string resourceDir=""; string content=""; string c=""; string[] splStr; string fileName=""; string sel="";

    resourceDir= cfg_webImages + "\\" + folderName + "\\";

    content= getDirHtmlNameList(resourceDir);

    thisValue= lCase(thisValue); //ת��Сд �öԱ�

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

//���ģ��Ƥ�� ShowWebModuleSkins("ModuleSkins", ModuleSkins)
string showWebModuleSkins(string inputName, string thisValue){
    string resourceDir=""; string content=""; string c=""; string[] splStr; string fileName=""; string sel="";
    resourceDir= cfg_webImages + "\\Index\\column";
    //Call Echo("ResourceDir",ResourceDir)
    content= getDirFolderNameList(resourceDir);
    //Call Echo("Content",Content)

    thisValue= lCase(thisValue); //ת��Сд �öԱ�

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

//��ʾ��ѡ���б�
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
//��ʾInput��ѡ InputCheckBox("Id",ID,"")
string inputCheckBox(string textName, bool isChecked, string helpStr){
    //Dim sel
    //If CStr(valueStr) = "True" Or CStr(isChecked) = "1" Then sel = " isChecked" Else sel = ""
    //inputCheckBox = "<input type='checkbox' name='" & textName & "' id='" & textName & "'" & sel & " value='1'>"
    //If helpStr <> "" Then inputCheckBox = "<label for='" & textName & "'>" & inputCheckBox & helpStr & "</label> "
    return handleInputCheckBox(textName, isChecked, "1", helpStr, "");
}
//��ʾInput��ѡ InputCheckBox("Id",ID,"")
string inputCheckBox3(string textName, bool isChecked, string valueStr, string helpStr){
    return handleInputCheckBox(textName, isChecked, valueStr, helpStr, "newidname");
}
string handleInputCheckBox(string textName, bool isChecked, string valueStr, string helpStr, string sType){
    string s=""; string sel=""; string idName="";
    if( cStr(valueStr)== "True" || isChecked== true ){ sel= " checked" ;}else{ sel= "" ;}
    idName= textName; //id�������ļ�����
    sType= "|" + sType + "|";
    if( inStr(sType, "|newidname|") > 0 ){
        idName= textName + phpRand(1, 9999);
    }
    s= "<input type='checkbox' name='" + textName + "' id='" + idName + "'" + sel + " value='" + valueStr + "'>";
    if( helpStr != "" ){ s= "<label for='" + idName + "'>" + s + helpStr + "</label> " ;}
    return s;
}

//��ʾInput�ı�  InputText("FolderName", FolderName, "40px", "��������")
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
//��ʾInput�ı�  InputText("FolderName", FolderName, "40px", "��������")
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
//��ʾInput�ı������  InputLeftText(TextName, ValueStr, "98%", "")
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
//��ʾInput�ı������ �����������ұ�
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
//��ʾInput�ı����б� ��ʾ�ı������
string inputLeftTextContent(string textName, string valueStr, string width, string helpStr){
    return handleInputLeftRightTextContent("���", textName, valueStr, width, helpStr);
}
//��ʾInput�ı����б� ��ʾ�ı����ұ�
string inputRightTextContent(string textName, string valueStr, string width, string helpStr){
    return handleInputLeftRightTextContent("�ұ�", textName, valueStr, width, helpStr);
}
//��ʾInput�ı����б� ��ʾ�ı������ �� ��ʾ�ı����ұ� 20150114
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

    if( sType== "���" ){
        handleInputLeftRightTextContent= helpStr + handleInputLeftRightTextContent + vbCrlf();
    }else{
        handleInputLeftRightTextContent= handleInputLeftRightTextContent + helpStr;
    }

    return handleInputLeftRightTextContent;
}

//��ʾInput�ı����������
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
//��ʾInput�ı��������������
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
//��ʾInput�����ı�
string inputHiddenText(string textName, string valueStr){
    return "<input name=\"" + textName + "\" type=\"hidden\" id=\"" + textName + "\" value=\"" + valueStr + "\" />" + vbCrlf();
}
//��ʾInput�ı��� InputTextArea("FindTpl", FindTpl, "60%" , "120px", "")
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
        if( checkNumber(height) ){ //�Զ��Ӹ�px����
            height= height + "px";
        }
        heightStr= "height:" + height + ";";
        if( css != "" ){
            css= replace(css, ";'", ";" + heightStr + ";'");
        }else{
            css= " style='height:" + height + ";'";
        }
    }
    css= replace(css, ";;", ";"); //ȥ�������ֵ
    return "<textarea name=\"" + textName + "\" type=\"text\" id=\"" + textName + "\"" + css + ">" + valueStr + "</textarea>" + helpStr;
}
//��ʾ����Input�ı��� InputTextArea("WebDescription", WebDescription, "99%", "100px", "")
string inputHiddenTextArea(string textName, string valueStr, string width, string height, string helpStr){
    return handleInputHiddenTextArea(textName, valueStr, width, height, "", helpStr);
}
//��ʾ����Input�ı��� InputTextArea("WebDescription", WebDescription, "99%", "100px", "")
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
//��ʾĿ¼�б� ��Select��ʽ��ʾ
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
//��Input�Ӹ�Disabled���ɲ���
string inputDisabled( string content){
    return replace(content, "<input ", "<input disabled=\"disabled\" ");
}

//��Input�Ӹ�rel��ϵ����
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



//���ٵ�������====================================================

//��վ����
string webTitle_InputTextArea(string webTitle){
    return inputText("WebTitle", webTitle, "70%", "  ����ؼ�����-����"); //����Ϊ��վĬ�ϱ���
}
//��վ�ؼ���
string webKeywords_InputText(string webKeywords){
    return inputText("WebKeywords", webKeywords, "70%", " ���ԣ�����(���Ķ���)");
}
//��վ����
string webDescription_InputTextArea(string webDescription){
    return inputTextArea("WebDescription", webDescription, "99%", "100px", "");
}
//��̬�ļ�����
string folderName_InputText(string folderName){
    return inputText("FolderName", folderName, "40%", "");
}
//��̬�ļ���
string fileName_InputText(string fileName){
    return inputText("FileName", fileName, "40%", ".html Ҳ�����������ϵ����ӵ�ַ");
}
//ģ���ļ���

string templatePath_InputText(string templatePath){
    return inputText("TemplatePath", templatePath, "40%", " ����ΪĬ��");
}
//���ƴ����ť����
string clickPinYinHTMLStr(string did){
    return "<a href=\"javascript:GetPinYin('FolderName','" + did + "','AjAx.Asp?act=GetPinYin')\" >���ƴ��</a>";
}
//ѡ���ı���ɫ���ı��Ӵ�
string showFontColorFontB(string fontColor, bool isFontB){
    return selectFontColor(fontColor) + inputCheckBox("FontB", isFontB, "�Ӵ�");
}
//��ʾ�ı�TEXT����
string showSort(string sort){
    string showSort= inputText("Sort", sort, "30px", "");
    return replace(showSort, ";'", ";text-align:center;'");
}
//��վ�������Ͷ����ײ���
string showWebNavType(bool isNavTop, bool isNavButtom, bool isNavLeft, bool isNavContent, bool isNavRight, bool isNavOthre){
    string c="";
    c= c + inputCheckBox("NavTop", isNavTop, "��������");
    c= c + inputCheckBox("NavButtom", isNavButtom, "�ײ�����");
    c= c + inputCheckBox("NavLeft", isNavLeft, "��ߵ���");
    c= c + inputCheckBox("NavContent", isNavContent, "�м䵼��");
    c= c + inputCheckBox("NavRight", isNavRight, "�ұߵ���");
    c= c + inputCheckBox("NavOthre", isNavOthre, "��������");
    return c;
}
string showOnHtml(bool isOnHtml){
    return inputCheckBox("OnHtml", isOnHtml, "����HTML");
}
string showThrough(bool isThrough){
    return inputCheckBox("Through", isThrough, "���");
}
string showRecommend(bool isRecommend){
    return inputCheckBox("Recommend", isRecommend, "�Ƽ�");
}
//��ʾ������ر�ͼƬ
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
    //�ɰ�
    //ShowOnOffImg = "<a href="& adminDir &"HandleDatabase.Asp?act=SetTrueFalse&Table=" & Table & "&FieldName=" & FieldName & "&Url=" & Url & "&Id=" & Id & Temp & ">" & Img & "</a>"
    return showOnOffImg;
}
//��ʾ������ر�ͼƬ
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


//��ÿ���Css��ʽ 20150128  ��ʱ����
string controlDialogCss(){
    string c="";
    c= "<style>" + vbCrlf();
    c= c + "/*����Css20150128*/" + vbCrlf();
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


//ɾ�����ݴ����
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
