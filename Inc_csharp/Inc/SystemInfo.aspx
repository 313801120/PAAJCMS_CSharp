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
//ϵͳ��Ϣ  (2014,05,27)



//����ϵͳ�汾
string operationSystem(){
    string httpAgent=""; string systemVer="";
    httpAgent= cStr(Request.ServerVariables["HTTP_USER_AGENT"]);
    if( inStr(httpAgent, "NT 5.2") > 0 ){
        systemVer= "Windows Server 2003";
    }else if( inStr(httpAgent, "NT 5.1") > 0 ){
        systemVer= "Windows XP";
    }else if( inStr(httpAgent, "NT 5") > 0 ){
        systemVer= "Windows 2000";
    }else if( inStr(httpAgent, "NT 4") > 0 ){
        systemVer= "Windows NT4";
    }else if( inStr(httpAgent, "4.9") > 0 ){
        systemVer= "Windows ME";
    }else if( inStr(httpAgent, "98") > 0 ){
        systemVer= "Windows 98";
    }else if( inStr(httpAgent, "95") > 0 ){
        systemVer= "Windows 95";
    }else{
        systemVer= httpAgent;
    }
    return httpAgent;
}
//����Ƿ�Ϊ�ֻ�
bool checkMobile(){
    bool checkMobile= false;
    if( cStr(Request.ServerVariables["HTTP_X_WAP_PROFILE"]) != "" ){
        checkMobile= true;
    }
    return checkMobile;
}

//���IIS�汾��
string getIISVersion(){
    return cStr(Request.ServerVariables["SERVER_SOFTWARE"]);
}

</script>
