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
//系统信息  (2014,05,27)



//操作系统版本
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
//检测是否为手机
bool checkMobile(){
    bool checkMobile= false;
    if( cStr(Request.ServerVariables["HTTP_X_WAP_PROFILE"]) != "" ){
        checkMobile= true;
    }
    return checkMobile;
}

//获得IIS版本号
string getIISVersion(){
    return cStr(Request.ServerVariables["SERVER_SOFTWARE"]);
}

</script>
