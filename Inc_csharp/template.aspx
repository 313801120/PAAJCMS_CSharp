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
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>模板文件管理</title>
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
    if(confirm("确认要删除吗？删除后将不可恢复！"))
    return true;
    else
    return false;
}
</script>
<script runat="server" language="c#">

if( cStr(Session["adminusername"])== "" ){
    eerr("提示", "未登录，请先登录");
}

conn=openConn();
conn.Open();

switch ( cStr(Request["act"]) ){
    case "templateFileList" : displayTemplateDirDialog(cStr(Request["dir"])) ; templateFileList(cStr(Request["dir"]));break;//模板列表
    case "delTemplateFile" : delTemplateFile(cStr(Request["dir"]), cStr(Request["fileName"])) ; displayTemplateDirDialog(cStr(Request["dir"])) ; templateFileList(cStr(Request["dir"]))		;break;//删除模板文件
    case "addEditFile" : displayTemplateDirDialog(cStr(Request["dir"])) ; addEditFile(cStr(Request["dir"]), cStr(Request["fileName"]));break;//显示添加修改文件
    default : displayTemplateDirDialog(cStr(Request["dir"])); //显示模板目录面板
    break;
}

//模板文件列表
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

            s= "<a href=\"../aspxweb.aspx?templatedir=" + escape(dir) + "&templateName=" + fileName + "\" target='_blank'>预览</a> ";
            echo("<img src='Images/file/"+ fileType +".gif'>" + fileName + "（"+ printSpaceValue(getFSize(filePath)) +"）", s + "| <a href='?act=addEditFile&dir=" + dir + "&fileName=" + fileName + "'>修改</a> | <a href='?act=delTemplateFile&dir=" + cStr(Request["dir"]) + "&fileName=" + fileName + "' onclick='return checkDel()'>删除</a>");
        }
    }



}

//删除模板文件
void delTemplateFile(string dir, string fileName){
    string filePath="";

    handlePower("删除模板文件");						//管理权限处理

    filePath= dir + "/" + fileName;
    deleteFile(filePath);
    echo("删除文件", filePath);
}

//显示面板样式列表
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


//添加修改文件
string addEditFile(string dir, string fileName){
    string filePath="";string promptMsg="";

    if( right(lCase(fileName), 5) != ".html" && cStr(Session["adminusername"]) != "PAAJCMS" ){
        fileName= fileName + ".html";
    }
    filePath= dir + "/" + fileName;

    if( checkFile(filePath)==false ){
        handlePower("添加模板文件");						//管理权限处理
    }else{
        handlePower("修改模板文件");						//管理权限处理
    }

    //保存内容
    if( cStr(Request["issave"])== "true" ){
        createFile(filePath, cStr(Request["content"]));
        promptMsg="保存成功";
    }
    </script>
    <form name="form1" method="post" action="?act=addEditFile&issave=true">
    <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline">
    <tr>
    <td height="30">目录<%=dir%><br>
    <input name="dir" type="hidden" id="dir" value="<%=dir%>" /></td>
    </tr>
    <tr>
    <td>文件名称
    <input name="fileName" type="text" id="fileName" value="<%=fileName%>" size="40">&nbsp;<input type="submit" name="button" id="button" value=" 保存 " /><%=promptMsg%>
    <br>
    <textarea name="content" style="width:99%;height:480px;"id="content"><% rw(getFText(filePath))%></textarea></td>
    </tr>
    </table>
    </form>
    <script runat="server" language="c#"> return "";
}
//文件夹搜索
string displayTemplateDirDialog(string dir){
    string folderPath="";
    </script>
    <form name="form2" method="post" action="?act=templateFileList">
    <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline">
    <tr>
    <td height="30"><input name="dir" type="text" id="dir" value="<%=dir%>" size="60" />
    <input type="submit" name="button2" id="button2" value=" 进入 " /><script runat="server" language="c#">
    folderPath=dir + "/images/column/";
    if( checkFolder(folderPath) ){
        rw("面板样式" + displayPanelList(folderPath));
    }
    folderPath=dir + "/images/nav/";
    if( checkFolder(folderPath) ){
        rw("导航样式" + displayPanelList(folderPath));
    }
    </script></td>
    </tr>
    </table>
    </form>
    <% return "";
}%>



