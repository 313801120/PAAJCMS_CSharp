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
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ģ���ļ�����</title>
</head>
<body>
<style type="text/css">
<!--
body {
    margin-left: 0px;
    margin-top: 0px;
    margin-right: 0px;
    margin-bottom: 0px;
}
a:link,a:visited,a:active {
    color: #000000;
    text-decoration: none;
}
a:hover {
    color: #666666;
    text-decoration: none;
}
.tableline{
    border: 1px solid #999999;
}
body,td,th {
    font-size: 12px;
}
a {
    font-size: 12px;
}
-->
</style>
<script language="javascript">
function checkDel()
    {
    if(confirm("ȷ��Ҫɾ����ɾ���󽫲��ɻָ���"))
    return true;
    else
    return false;
}
</script>
<script runat="server" language="c#">

if( cStr(Session["adminusername"])== "" ){
    eerr("��ʾ", "δ��¼�����ȵ�¼");
}

conn=openConn();
conn.Open();

switch ( cStr(Request["act"]) ){
    case "templateFileList" : displayTemplateDirDialog(cStr(Request["dir"])) ; templateFileList(cStr(Request["dir"]));break;//ģ���б�
    case "delTemplateFile" : delTemplateFile(cStr(Request["dir"]), cStr(Request["fileName"])) ; displayTemplateDirDialog(cStr(Request["dir"])) ; templateFileList(cStr(Request["dir"]))		;break;//ɾ��ģ���ļ�
    case "addEditFile" : displayTemplateDirDialog(cStr(Request["dir"])) ; addEditFile(cStr(Request["dir"]), cStr(Request["fileName"]));break;//��ʾ����޸��ļ�
    default : displayTemplateDirDialog(cStr(Request["dir"])); //��ʾģ��Ŀ¼���
    break;
}

//ģ���ļ��б�
void templateFileList(string dir){
    string content=""; string[] splStr; string fileName=""; string s="";string fileType="" ;string folderName="";string filePath="";

    if( cStr(Session["adminusername"])== "PAAJCMS" ){
        content= getDirFolderNameList(dir,"");
        splStr= aspSplit(content, vbCrlf());
        foreach(var eachfolderName in splStr){
            folderName=eachfolderName;
            s="<a href='?act=templateFileList&dir="+ dir + "/" + folderName +"'>"+ folderName +"</a>";
            echo("<img src='Images/file/folder.gif'>",s);
        }
        content= getDirFileNameList(dir,"");
    }else{
        content= getDirHtmlNameList(dir);
    }
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachfileName in splStr){
        fileName=eachfileName;
        if( fileName!="" ){
            fileType=lCase(getFileAttr(fileName,4));
            filePath=dir + "/" + filename;
            if( inStr("|asa|asp|aspx|bat|bmp|cfm|cmd|com|css|db|default|dll|doc|exe|fla|folder|gif|h|htm|html|inc|ini|jpg|js|jtbc|log|mdb|mid|mp3|php|png|rar|real|rm|swf|txt|wav|xls|xml|zip|","|"+ fileType +"|")==0 ){
                fileType="default";
            }

            s= "<a href=\"../aspxweb.aspx?templatedir=" + escape(dir) + "&templateName=" + fileName + "\" target='_blank'>Ԥ��</a> ";
            echo("<img src='Images/file/"+ fileType +".gif'>" + fileName + "��"+ printSpaceValue(getFSize(filePath)) +"��", s + "| <a href='?act=addEditFile&dir=" + dir + "&fileName=" + fileName + "'>�޸�</a> | <a href='?act=delTemplateFile&dir=" + cStr(Request["dir"]) + "&fileName=" + fileName + "' onclick='return checkDel()'>ɾ��</a>");
        }
    }



}

//ɾ��ģ���ļ�
void delTemplateFile(string dir, string fileName){
    string filePath="";

    handlePower("ɾ��ģ���ļ�");						//����Ȩ�޴���

    filePath= dir + "/" + fileName;
    deleteFile(filePath);
    echo("ɾ���ļ�", filePath);
}

//��ʾ�����ʽ�б�
string displayPanelList(string dir){
    string content="";string[] splstr;string s="";string c="";
    content=getDirFolderNameList(dir);
    splstr=aspSplit(content,vbCrlf());
    c="<select name='selectLeftStyle'>";
    foreach(var eachs in splstr){
        s=eachs;
        s="<option value=''>"+ s +"</option>";
        c=c + s + vbCrlf();
    }
    return c + "</select>";
}


//����޸��ļ�
string addEditFile(string dir, string fileName){
    string filePath="";string promptMsg="";

    if( right(lCase(fileName), 5) != ".html" && cStr(Session["adminusername"]) != "PAAJCMS" ){
        fileName= fileName + ".html";
    }
    filePath= dir + "/" + fileName;

    if( checkFile(filePath)==false ){
        handlePower("���ģ���ļ�");						//����Ȩ�޴���
    }else{
        handlePower("�޸�ģ���ļ�");						//����Ȩ�޴���
    }

    //��������
    if( cStr(Request["issave"])== "true" ){
        createFile(filePath, cStr(Request["content"]));
        promptMsg="����ɹ�";
    }
    </script>
    <form name="form1" method="post" action="?act=addEditFile&issave=true">
    <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline">
    <tr>
    <td height="30">Ŀ¼<%=dir%><br>
    <input name="dir" type="hidden" id="dir" value="<%=dir%>" /></td>
    </tr>
    <tr>
    <td>�ļ�����
    <input name="fileName" type="text" id="fileName" value="<%=fileName%>" size="40">&nbsp;<input type="submit" name="button" id="button" value=" ���� " /><%=promptMsg%>
    <br>
    <textarea name="content" style="width:99%;height:480px;"id="content"><% rw(getFText(filePath))%></textarea></td>
    </tr>
    </table>
    </form>
    <script runat="server" language="c#"> return "";
}
//�ļ�������
string displayTemplateDirDialog(string dir){
    string folderPath="";
    </script>
    <form name="form2" method="post" action="?act=templateFileList">
    <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline">
    <tr>
    <td height="30"><input name="dir" type="text" id="dir" value="<%=dir%>" size="60" />
    <input type="submit" name="button2" id="button2" value=" ���� " /><script runat="server" language="c#">
    folderPath=dir + "/images/column/";
    if( checkFolder(folderPath) ){
        rw("�����ʽ" + displayPanelList(folderPath));
    }
    folderPath=dir + "/images/nav/";
    if( checkFolder(folderPath) ){
        rw("������ʽ" + displayPanelList(folderPath));
    }
    </script></td>
    </tr>
    </table>
    </form>
    <% return "";
}%>



