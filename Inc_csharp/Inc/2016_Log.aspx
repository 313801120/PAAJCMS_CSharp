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
//日志文件

//错误日志
string errorLog(string content){
    if( openErrorLog== true ){
        rw(content);
    }
    return "";
}

//写入系统操作日志
string writeSystemLog(string tableName, string msgStr){
    string logFile=""; string s=""; string url=""; string ip=""; string addDateTime="";string logDir="";
    logDir= adminDir + "/data/systemLog/";
    createDirFolder(logDir);		//创建文件夹
    logFile=logDir + "/" + format_Time(now(), 2) + ".txt";
    url= ADSql(getThisUrlFileParam());
    addDateTime= format_Time(now(), 1);
    ip= getIP();
    if( inStr(openWriteSystemLog, "|txt|") > 0 ){
        s= s + "姓名：" + cStr(Session["adminusername"]) + vbCrlf();
        s= s + "表：" + tableName + vbCrlf();
        s= s + "信息：" + msgStr + vbCrlf();
        s= s + "网址：" + url + vbCrlf();
        s= s + "时间：" + addDateTime + vbCrlf();
        s= s + "IP：" + ip + vbCrlf();
        s= s + "------------------------" + vbCrlf();
        createAddFile(logFile, s);
        //call echo(logfile,"log")
    }

    if( inStr(openWriteSystemLog, "|txt|") > 0 ){
        conn=openConn();
        conn.Open();

        //判断表存在
        if( inStr(getHandleTableList(),"|"+ db_PREFIX + "systemlog" +"|")>0 ){
            connexecute("insert into " + db_PREFIX + "SystemLog (tablename,msgstr,url,adminname,ip,adddatetime) values('" + tableName + "','" + msgStr + "','" + url + "','" + cStr(Session["adminusername"]) + "','" + ip + "','" + addDateTime + "')");
        }
    }

    return "";
}

</script>
