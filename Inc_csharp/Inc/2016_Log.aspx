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
//��־�ļ�

//������־
string errorLog(string content){
    if( openErrorLog== true ){
        rw(content);
    }
    return "";
}

//д��ϵͳ������־
string writeSystemLog(string tableName, string msgStr){
    string logFile=""; string s=""; string url=""; string ip=""; string addDateTime="";string logDir="";
    logDir= adminDir + "/data/systemLog/";
    createDirFolder(logDir);		//�����ļ���
    logFile=logDir + "/" + format_Time(now(), 2) + ".txt";
    url= ADSql(getThisUrlFileParam());
    addDateTime= format_Time(now(), 1);
    ip= getIP();
    if( inStr(openWriteSystemLog, "|txt|") > 0 ){
        s= s + "������" + cStr(Session["adminusername"]) + vbCrlf();
        s= s + "��" + tableName + vbCrlf();
        s= s + "��Ϣ��" + msgStr + vbCrlf();
        s= s + "��ַ��" + url + vbCrlf();
        s= s + "ʱ�䣺" + addDateTime + vbCrlf();
        s= s + "IP��" + ip + vbCrlf();
        s= s + "------------------------" + vbCrlf();
        createAddFile(logFile, s);
        //call echo(logfile,"log")
    }

    if( inStr(openWriteSystemLog, "|txt|") > 0 ){
        conn=openConn();
        conn.Open();

        //�жϱ����
        if( inStr(getHandleTableList(),"|"+ db_PREFIX + "systemlog" +"|")>0 ){
            connexecute("insert into " + db_PREFIX + "SystemLog (tablename,msgstr,url,adminname,ip,adddatetime) values('" + tableName + "','" + msgStr + "','" + url + "','" + cStr(Session["adminusername"]) + "','" + ip + "','" + addDateTime + "')");
        }
    }

    return "";
}

</script>
