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
<!--#include file="Include/asp.aspx"-->
<!--#include file="Include/Conn.aspx"--> 
<!--#include file="Include/sys_FSO.aspx"-->
<!--#include file="Include/sys_System.aspx"--> 
<!--#include file="Include/sys_SessionCookie.aspx"--> 
<!--#include file="Include/sys_http.aspx"-->  
<!--#include file="Include/sys_patch.aspx"-->  


<!--#include file="Inc/2014_Array.aspx"--> 
<!--#include file="Inc/2014_Author.aspx"--> 
<!--#include file="Inc/2014_Css.aspx"-->  
<!--#include file="Inc/2014_Js.aspx"--> 
<!--#include file="Inc/2014_Template.aspx"-->
<!--#include file="Inc/2015_APGeneral.aspx"--> 
<!--#include file="Inc/2015_Color.aspx"--> 
<!--#include file="Inc/2015_Editor.aspx"--> 
<!--#include file="Inc/2015_Formatting.aspx"--> 
<!--#include file="Inc/2015_NewWebFunction.aspx"--> 
<!--#include file="Inc/2015_Param.aspx"--> 
<!--#include file="Inc/2015_ToMyPHP.aspx"--> 
<!--#include file="Inc/2016_Log.aspx"--> 
<!--#include file="Inc/2016_SaveData.aspx"--> 
<!--#include file="Inc/2016_WebControl.aspx"--> 
<!--#include file="Inc/ASPPHPAccess.aspx"--> 
<!--#include file="Inc/Cai.aspx"--> 
<!--#include file="Inc/Check.aspx"--> 
<!--#include file="Inc/Common.aspx"--> 
<!--#include file="Inc/EncDec.aspx"--> 
<!--#include file="Inc/function.aspx"--> 
<!--#include file="Inc/FunHTML.aspx"--> 
<!--#include file="Inc/Html.aspx"--> 
<!--#include file="Inc/IE.aspx"--> 
<!--#include file="Inc/Incpage.aspx"--> 
<!--#include file="Inc/PinYin.aspx"--> 
<!--#include file="Inc/Print.aspx"--> 
<!--#include file="Inc/StringNumber.aspx"--> 
<!--#include file="Inc/SystemInfo.aspx"--> 
<!--#include file="Inc/Time.aspx"--> 
<!--#include file="Inc/URL.aspx"--> 
<!--#include file="Inc/2014_GBUTF.aspx"--> 
<!--#include file="Inc/2014_Form.aspx"-->   
<!--#include file="config.aspx"-->

<!--#include file="function.aspx"--> 
<!--#include file="setAccess.aspx"-->
 
<script runat="server" language="c#">
string replaceGlobleVariable(string s){
	return s;
}
string handleAction(string s){
	return s;
}
string loadWebConfig(){
	return "";
}
protected void Page_Load(object sender, EventArgs e){

	loadRun();	
	//�ر����ݿ�
	if(conn!=null){
		if (conn.State != ConnectionState.Closed){
			conn.Close(); 
		}
	}
	
}  
//=========
</script>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��װSQLServer/Access���ݿ�,QQ:313801120:::</title>
<style type="text/css">
a img{border:none}
    .imga{vertical-align:bottom;}
    .imgb{vertical-align:bottom; display:block;}/*���ͼƬ�����п�϶�ļ򵥷���  ��ǰ��img{����ͼƬ�ͻỻ��*/
body,div,p,img,dl,dt,dd,ul,ol,li,h1,h2,h3,h4,h5,h6,pre,form,fieldset,input,textarea,blockquote{padding:0px;margin:0px}
    li{list-style-type:none}
    a{font-size:12px;line-height:18px;color:#000000;text-decoration:none}
    a:hover{text-decoration:none;color:#000099}
    /*PHPCMS����ʽ*/
    .input-text, .measure-input, textarea, input.date, input.endDate, .input-focus {
        border: 1px solid #A7A6AA;
        height: 22px;
        line-height:22px;
        margin: 0 5px 0 0;
        padding: 2px 0 2px 5px;
        border: 1px solid #d0d0d0;
        background: #FFF url(../images/input.png) repeat-x;
        font-family: Verdana, Geneva, sans-serif,"����";
        font-size: 12px;
    }
    input.date, input.endDate {
        background: #fff url(../images/input_date.png) no-repeat right 3px;
        padding-right: 18px;
        font-size: 12px;
    }
    select {
        vertical-align: middle;
        background: none repeat scroll 0 0 #F9F9F9;
        border-color: #666666 #CCCCCC #CCCCCC #666666;
        border-style: solid;
        border-width: 1px;
        color: #333;
        padding: 2px;
    }
    /*�Զ��尴ť ���Ľ�*/
    .btnclick1{
        color: #000;
        font-size: 14px;
        padding:0 20px;
        background-color:#fff;
        border:0px;
        border:1px solid #666666;
        text-decoration:none;
        cursor:pointer;
        line-height: 26px;
        font-weight:bold;
        border-radius:5px;-moz-border-radius:5px;-webkit-border-radius:5px;-o-border-radius:5px;
    }
    .btnclick1:hover{
        background-color:#E6E6E6;
        background-position: 0px -50px;
    }
    .btnclick1:active{
        background-color:#fff;
    }
    .pright {
        width: 720px;
        margin:0 auto;
    }
    .pr-title {
        width: 720px;
        height: 22px;
        margin: 8px auto 0px;
        overflow: hidden;
    }
    .pr-title h3 {
        width: 158px;
        height: 22px;
        line-height: 22px;
        overflow: hidden;
        display: block;
        font-size: 12px;
        padding-top: 1px;
        text-indent: 10px;
        letter-spacing: 2px;
        color: #6D8A4F;
        font-weight: bold;
    }
    input{
        vertical-align:middle;
        margin-right:3px;
        font-size:12px;
    }
    textarea{
        vertical-align:top;
        font-size:12px;
        line-height:156%;
        border:1px solid #AAA;
        padding:3px;
        letter-spacing:1px;
        word-break:break-all;
        overflow-y:auto;
    }
    .input-txt{
        padding:4px 8px 4px 6px;
        border:1px solid #AAA;
        font-size:12px;
        color:#000;
        width:200px;
    }
    .textipt_on{
        border:1px solid #F90;
    }
    hr{
        height:1px;
        line-height:1px;
        overflow:hidden;
        border-width:1px 0px 0px 0px ;
        border-top:1px solid #E6E6E6;/*����Opera*/
    }
    hr:empty {
        margin:8px 0px 7px 0px !important;
        margin:0px;
    }
    small{
        font-size:12px;
    }
    .moncolor td{
        background:#FFC;
    }
    .twbox{
        width:706px;
        border:1px solid #CFDCC9;
        font-size:12px;
        overflow:hidden;
        margin:8px auto;
    }
    .twbox thead tr td{
        height:31px;
        line-height:31px;
        text-indent:10px;
    }
    .twbox thead tr td strong{
        letter-spacing:2px;
        margin-right:14px;
        color:#FFF;
        font-size:14px;
    }
    .twbox thead tr td span{
        color:#CDA;
    }
    .twbox thead tr td p{
        height:31px;
        display:inline;
        float:right;
        margin:-31px 10px 0 0;
        overflow:hidden;
    }
    .twbox thead tr td p *{
        float:right;
    }
    .twbox thead tr td a.thlink{
        color:#FFF;
    }
    .twbox thead tr td a.thlink:hover{
        color:#FFFF00;
        text-decoration:none;
    }
    .twbox tbody {
        overflow:hidden;
        text-align:left;
    }
    .twbox tbody tr th{
        background:#F9FCEF;
        color:#6D8A4F;
        line-height:21px;
        height:21px;
        text-indent:30px;
        font-weight:normal;
        border-bottom:1px solid #EFF7D8;
        letter-spacing:2px;
    }
    .twbox tbody tr td{
        padding:7px;
        border-bottom:1px solid #F2F2F2;
        color:#333;
        vertical-align:top;
    }
    .twbox tbody tr td p{
        line-height:21px;
    }
    .twbox tbody tr td p strong img{
        vertical-align:middle;
    }
    .twbox tbody tr td img{
        vertical-align:top;
        margin:0px 10px 5px 0px;
    }
    .twbox tbody tr td small{
        color:#888;
    }
    .twbox tfoot tr td{
        padding:10px;
        line-height:25px;
        text-align:center;
    }
    .twbox tfoot tr td p{
        line-height:21px;
        margin-bottom:10px;
    }
    input.but{
        height:26px;
        padding-left:6px;
        padding-right:6px;
        line-height:26px;
        font-weight:bold;
        letter-spacing:1px;
        color:#FFF;
        background-color:#FC3;
    }
    .onetd{
        width:120px;
        text-align:right;
        line-height:25px;
    }
    .mytipwrap{
        line-height:30px;
        color:#999999;
    }
    a.mytip {
        line-height: 14px;
        padding: 6px 20px;
        border-style: solid;
        border-width: 1px;
        border-color: #EEE #CCC #CCC #EEE;
        background: #FAFAFA;
        color: #333;
        margin-right:10px;
        text-decoration: none;
    }
    </style>
    </head>
    <body>
    <script runat="server" language="c#">


    string selectDatabase=""; string dbhost=""; string dbuser=""; string dbpwd=""; string dbname=""; string accessDir=""; string accessPath=""; string configFilePath=""; string content=""; string startStr=""; string endStr=""; string findStr=""; string replaceStr="";
    string tempContent=""; string msgStr=""; bool isYes; string connStr="";
    string showLayout="";				//��ʾ����


    void loadRun(){


        if( EDITORTYPE=="asp" ){
            configFilePath= "/inc/Config.Asp";
        }else if( EDITORTYPE=="aspx" ){
            configFilePath= "/inc_csharp/Config.Aspx";
        }


        switch ( cStr(Request["act"]) ){
            case "createAccess" : createAccess() ;break;//�������ݿ�
            case "displayImportTableLayout" : displayImportTableLayout() ;break;//��ʾ��������ҳ��
            case "displayImportData" : displayImportData() ;break;//��ʾ��������
            default : displayDefault();

            break;
        }
    }



    //ע�⣬��������ʱ Default """"  Ĭ��ֵ�����   ��Ҫ������Default ''    20150506
    //��վ���ݿ�
    void webData(string db_PREFIX, string loginname, string loginpwd){
        string tableName=""; string sql=""; bool isAddData; string[] splStr; string[] splxx; int i=0; string s=""; string[] splTrue; string bigClassName="";
        string[] splUrl; string[] splNavName; string sNavClassName=""; string[] splImg; string templateIndex=""; string templateHome=""; string templateMain="";
        //call echo("db_PREFIX",db_PREFIX)
        //��վ��Ŀ��-------------------------------
        tableName= db_PREFIX + "WebColumn";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";

            sql= sql + "columnName VarChar Default \"\","; //��������
            sql= sql + "columnEnName VarChar Default \"\","; //��������(Ӣ��)
            sql= sql + "columnType VarChar Default \"\","; //��������
            sql= sql + "parentId int Default -1,"; //����ĿID
            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "views int Default 0,"; //�������
            sql= sql + "adminId int Default 0,"; //����Աid

            sql= sql + "isDisplay YesNo,"; //�Ƿ���ʾ
            sql= sql + "smallImage VarChar Default \"\","; //СͼƬ��ַ
            sql= sql + "bigImage VarChar Default \"\","; //��ͼƬ��ַ
            sql= sql + "bannerImage VarChar Default \"\","; //��ǰBannerͼƬ

            sql= sql + "flags VarChar Default \"\","; //��
            sql= sql + "displayTitle VarChar Default \"\","; //��ʾ����
            sql= sql + "lableTitle VarChar Default \"\","; //��ǩ����


            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "customAUrl VarChar Default \"\","; //�Զ���������ַ
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ
            sql= sql + "noFollow int Default 0,"; //��Ҫ׷�ٴ���ҳ�ϵ�����
            sql= sql + "fontColor VarChar Default \"\","; //����ɫ
            sql= sql + "fontB YesNo Default No,"; //�ִ�
            sql= sql + "isMakeHtml YesNo Default Yes,"; //��������Html




            sql= sql + "nPageSize int Default 10,"; //ÿҳ��ʾ����
            sql= sql + "sortSql VarChar Default \"\","; //����SQL
            sql= sql + "isHtml YesNo Default No,"; //�Ƿ�Ϊhtml
            sql= sql + "isOnHtml YesNo Default Yes,"; //�Ƿ�������Html

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��


            sql= sql + "noteBody Text Default \"\","; //�ʼ���Ϣ

            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "tempTXT1 Text Default \"\","; //�ݴ��ı�1
            sql= sql + "tempTXT2 Text Default \"\","; //�ݴ��ļ�2
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }

        //������Ϣ��-------------------------------
        tableName= db_PREFIX + "ArticleDetail";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "parentId int Default -1,"; //����ĿID
            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "views int Default 0,"; //�������
            sql= sql + "adminId int Default 0,"; //����Աid

            sql= sql + "smallImage VarChar Default \"\","; //СͼƬ��ַ
            sql= sql + "bigImage VarChar Default \"\","; //��ͼƬ��ַ
            sql= sql + "bannerImage VarChar Default \"\","; //��ǰBannerͼƬ
            sql= sql + "downloadFile VarChar Default \"\","; //�����ļ�
            sql= sql + "smallimageAlt VarChar Default \"\","; //Сͼalt
            sql= sql + "bigimageAlt VarChar Default \"\","; //��ͼalt
            sql= sql + "bannerimageAlt VarChar Default \"\","; //Banneralt


            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "titleColor VarChar Default \"\","; //������ɫ
            sql= sql + "titleAlt VarChar Default \"\","; //����alt
            sql= sql + "lableTitle VarChar Default \"\","; //��ǩ����

            sql= sql + "isThrough YesNo Default No,"; //ͨ�����

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��

            sql= sql + "occasions VarChar Default \"\","; //����
            sql= sql + "hotline VarChar Default \"\","; //����
            sql= sql + "model VarChar Default \"\","; //�ͺ�
            sql= sql + "author VarChar Default \"\","; //����
            sql= sql + "articleSource VarChar Default \"\","; //������Դ

            sql= sql + "price Float Default 0,"; //��Ʒ�۸�
            sql= sql + "newPrice Float Default 0,"; //��Ʒ�¼۸�20150729
            sql= sql + "memberPrice Float Default 0,"; //��Ա�۸�
            sql= sql + "sold Int Default 0,"; //���۳�

            sql= sql + "memberType VarChar Default \"\","; //��Ա����
            sql= sql + "memberUser VarChar Default \"\","; //��Ա��
            sql= sql + "hits Int Default 0,"; //�����



            sql= sql + "productAbout Text Default \"\","; //��Ʒ����
            sql= sql + "articleDescription Text Default \"\","; //ժҪ(���¼�Ҫ����)20141225

            sql= sql + "httpUrl VarChar Default \"\","; //��ַ
            sql= sql + "recordUrl VarChar Default \"\","; //Զ����ַ
            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ
            sql= sql + "customAUrl VarChar Default \"\","; //�Զ���������ַ
            sql= sql + "fontColor VarChar Default \"\","; //����ɫ
            sql= sql + "noFollow int Default 0,"; //��Ҫ׷�ٴ���ҳ�ϵ�����

            sql= sql + "flags VarChar Default \"\","; //��

            sql= sql + "isHtml YesNo Default No,"; //�Ƿ�Ϊhtml
            sql= sql + "isOnHtml YesNo Default Yes,"; //��������Html



            sql= sql + "articleInfoStyle VarChar Default \"\","; //�б�������ʽ
            sql= sql + "articleInfoPhotoWidth VarChar Default \"\","; //����ҳͼƬ��
            sql= sql + "articleInfoPhotoHeight VarChar Default \"\","; //����ҳͼƬ�� 0Ϊ��������� -1Ϊ��


            sql= sql + "relatedTags VarChar Default \"\","; //��ر�ǩ20141219
            sql= sql + "weight Int Default 0,"; //Ȩ�أ�     (ԽСԽ��ǰ)



            sql= sql + "noteBody Text Default \"\","; //�ʼ���Ϣ
            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���

            sql= sql + "tempTXT1 Text Default \"\","; //�ݴ��ı�1
            sql= sql + "tempTXT2 Text Default \"\","; //�ݴ��ļ�2
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }

        //��ҳ��-------------------------------
        tableName= db_PREFIX + "OnePage";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "displayTitle VarChar Default \"\","; //��ʾ����
            sql= sql + "adminId int Default 0,"; //����Աid

            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "customAUrl VarChar Default \"\","; //�Զ���������ַ
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ
            sql= sql + "fontColor VarChar Default \"\","; //����ɫ
            sql= sql + "fontB YesNo Default No,"; //�ִ�
            sql= sql + "noFollow int Default 0,"; //��Ҫ׷�ٴ���ҳ�ϵ�����
            sql= sql + "sortRank Int Default 0,"; //˳��
            sql= sql + "views int Default 0,"; //�������
            sql= sql + "isRecommend YesNo Default Yes,"; //�Ƽ�
            sql= sql + "lableTitle VarChar Default \"\","; //��ǩ����
            sql= sql + "banner VarChar Default \"\","; //��ǰBannerͼƬ

            sql= sql + "isHtml YesNo Default No,"; //�Ƿ�Ϊhtml
            sql= sql + "isOnHtml YesNo Default Yes,"; //��������Html

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��

            sql= sql + "noteBody Text Default \"\","; //�ʼ���Ϣ
            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }

        //�������۱� 20160129-------------------------------
        tableName= db_PREFIX + "TableComment";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "UserID Int Default 0,"; //�û�ID
            sql= sql + "itemID Int Default 0,"; //��ĿID ��ƷID
            sql= sql + "tableName VarChar Default \"\","; //������
            sql= sql + "userName VarChar Default \"\","; //�û�����
            sql= sql + "title VarChar Default \"\","; //���۱���
            sql= sql + "EMail VarChar Default \"\","; //����
            sql= sql + "tel VarChar Default \"\","; //�绰ID
            sql= sql + "ip VarChar Default \"\","; //������IP

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��

            sql= sql + "reply Text Default \"\","; //�ظ�����
            sql= sql + "noteBody Text Default \"\","; //�ʼ���Ϣ

            sql= sql + "isThrough Int Default 0,"; //�Ƿ�ͨ��
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }

        //��վ��Ϣ��-------------------------------
        tableName= db_PREFIX + "WebSite";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "WebSiteUrl VarChar Default \"\","; //��վ��ַ
            sql= sql + "webSiteBottom Text Default \"\","; //�ײ���Ϣ WebSiteBottom
            sql= sql + "WebSiteFlow int Default 999,"; //��վͳ����
            sql= sql + "WebSiteFlowStyle int Default 1,"; //��վͳ����ʽ
            sql= sql + "WebSiteFlowMedian int Default 6,"; //��վͳ�Ƽ�λ��Median��λ��
            sql= sql + "ProductList VarChar Default \"\","; //��Ʒ�б�
            sql= sql + "NewsList VarChar Default \"\","; //�����б�
            sql= sql + "NewsDid VarChar Default \"\","; //���Ŵ���
            //201109׷��
            sql= sql + "TZ51La VarChar Default \"\","; //ͳ����վ
            sql= sql + "UserEmail VarChar Default \"\","; //�û�Email
            sql= sql + "ProductDid VarChar Default \"\","; //��Ʒ����
            sql= sql + "TemplateIndex Text Default '��ҳ',"; //ģ����ҳ
            sql= sql + "TemplateHome Text Default '��ҳ',"; //ģ����ҳ
            sql= sql + "TemplateMain Text Default '��ϸҳ',"; //ģ����ϸҳ
            sql= sql + "TemplateMain2 Text Default '��ϸҳ',"; //ģ����ϸҳ2
            sql= sql + "TemplateMain3 Text Default '��ϸҳ',"; //ģ����ϸҳ3
            sql= sql + "UseNumb Text Default \"\","; //��վʹ�ô���
            //20111118׷��
            sql= sql + "WebRecord VarChar Default \"\","; //��վ��¼
            sql= sql + "ContentWebRecord Text Default \"\","; //ȫ����վ��¼
            //20111128׷��
            sql= sql + "UseHttpUrl VarChar Default \"\","; //��ǰʹ����ַ
            sql= sql + "TempUseHttpUrl Text Default \"\","; //�洢ʹ�ù�����ַ
            sql= sql + "WebDate DateTime Default Date(),"; //��վʱ��
            //2013,12,18
            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����

            sql= sql + "WebTemplate VarChar Default '\\Templates',"; //��վģ��·��
            sql= sql + "WebSkins VarChar Default '\\Skins\\Default\\',"; //��վ��ʽ·��'��ʱ��Ҫ��

            sql= sql + "WebFolderName VarChar Default 'test3',"; //��վ�ļ������� new20150721


            sql= sql + "WebImages VarChar Default '\\Skins\\Default\\Images\\',"; //��վImages·��
            sql= sql + "WebCss VarChar Default '\\Skins\\Default\\Css\\',"; //��վCss·��
            sql= sql + "WebJs VarChar Default '\\Skins\\Default\\Js\\',"; //��վJs·��

            sql= sql + "AddWebSite YesNo Default No,"; //URLǰ�������
            sql= sql + "UpdateHtml YesNo Default No,"; //�޸ĺ����HTML
            sql= sql + "isHtmlFormatting YesNo Default No,"; //�Ƿ�ΪHTML��ʽ��
            sql= sql + "isWebLabelClose YesNo Default No,"; //�Ƿ�Ϊ�պ���ҳ��ǩ
            sql= sql + "isCnToEn YesNo Default No,"; //�Ƿ�Ϊ�ı�������תƴ��
            sql= sql + "flags VarChar Default \"\","; //��


            //20150114
            sql= sql + "ModuleSkins VarChar Default \"\","; //ģ��Ƥ��
            sql= sql + "FindTpl VarChar Default \"\","; //����Ҫ�滻ģ������
            sql= sql + "ReplaceTpl VarChar Default \"\","; //�滻ģ��ָ�����ݣ�����

            sql= sql + "WebCodeFindTpl VarChar Default \"\","; //����Ҫ�滻��ҳ����
            sql= sql + "WebCodeReplaceTpl VarChar Default \"\","; //�滻��ҳָ�����ݣ�����

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��

            //�����
            sql= sql + "WebHtml VarChar Default 0)"; //��վ����״̬����ASP��HTML
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }

        //����Ա��-------------------------------
        tableName= db_PREFIX + "Admin";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "userName VarChar Default \"\","; //�˺�
            sql= sql + "pwd VarChar Default \"\","; //����
            sql= sql + "pseudonym VarChar Default \"\","; //����


            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��

            sql= sql + "regIP VarChar Default \"\","; //ע��IP
            sql= sql + "upIP VarChar Default \"\","; //�޸�IP


            sql= sql + "quanXian Text Default \"\","; //����Ȩ��
            sql= sql + "verificationMode Int Default 1,"; //��֤��ʽ��1ΪSession
            sql= sql + "adminLevel VarChar Default \"\","; //����
            sql= sql + "channel VarChar Default \"\","; //Ƶ��
            sql= sql + "mtest YesNo Default No,"; //�޸ĺ����HTML
            sql= sql + "flags Text Default \"\","; //��        ֻ�������̨���˴��� ��ΪȨ���ر��


            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //����ͳ�Ʊ�-------------------------------
        tableName= db_PREFIX + "SearchStat";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "title VarChar Default \"\","; //�����ؼ���
            sql= sql + "isThrough YesNo Default No,"; //ͨ��
            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "customAUrl VarChar Default \"\","; //�Զ���������ַ
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ
            sql= sql + "isHtml YesNo Default No,"; //�Ƿ�Ϊhtml
            sql= sql + "isOnHtml YesNo Default Yes,"; //�Ƿ�������Html
            sql= sql + "views int Default 0,"; //�������
            sql= sql + "author VarChar Default \"\","; //����

            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "fontColor VarChar Default \"\","; //����ɫ
            sql= sql + "noFollow int Default 0,"; //��Ҫ׷�ٴ���ҳ�ϵ�����
            sql= sql + "flags VarChar Default \"\","; //��

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //�������ӱ�-------------------------------
        tableName= db_PREFIX + "FriendLink";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";

            sql= sql + "adminId int Default 0,"; //����Աid
            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "titleColor VarChar Default \"\","; //������ɫ
            sql= sql + "lableTitle VarChar Default \"\","; //��ǩ����
            sql= sql + "httpUrl VarChar Default \"\","; //��ַ
            sql= sql + "sortRank int Default 0,"; //������

            sql= sql + "titleAlt VarChar Default \"\","; //����alt
            sql= sql + "smallImage VarChar Default \"\","; //СͼƬ��ַ
            sql= sql + "smallimageAlt VarChar Default \"\","; //Сͼalt


            sql= sql + "isThrough YesNo Default No,"; //ͨ�����
            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��


            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ
            sql= sql + "customAUrl VarChar Default \"\","; //�Զ���������ַ
            sql= sql + "fontColor VarChar Default \"\","; //����ɫ
            sql= sql + "noFollow int Default 0,"; //��Ҫ׷�ٴ���ҳ�ϵ�����
            sql= sql + "flags VarChar Default \"\","; //��

            sql= sql + "isHtml YesNo Default No,"; //�Ƿ�Ϊhtml
            sql= sql + "isOnHtml YesNo Default Yes,"; //��������Html

            sql= sql + "weight Int Default 0,"; //Ȩ�أ�     (ԽСԽ��ǰ)

            sql= sql + "noteBody Text Default \"\","; //�ʼ���Ϣ
            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //��Ƹ��-------------------------------
        tableName= db_PREFIX + "Job";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "title VarChar Default \"\","; //ְλ����
            sql= sql + "sex VarChar Default \"\","; //�Ա�
            sql= sql + "age VarChar Default \"\","; //����
            sql= sql + "education VarChar Default \"\","; //����
            sql= sql + "workArea VarChar Default \"\","; //������
            sql= sql + "monthly VarChar Default \"\","; //��н����
            sql= sql + "startDateTime VarChar Default \"\","; //��������
            sql= sql + "endDateTime VarChar Default \"\","; //��ֹ����
            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ

            sql= sql + "fontColor VarChar Default \"\","; //����ɫ
            sql= sql + "noFollow int Default 0,"; //��Ҫ׷�ٴ���ҳ�ϵ�����
            sql= sql + "flags VarChar Default \"\","; //��
            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "titleAlt VarChar Default \"\","; //����alt
            sql= sql + "smallImage VarChar Default \"\","; //СͼƬ��ַ
            sql= sql + "smallimageAlt VarChar Default \"\","; //Сͼalt
            sql= sql + "isThrough Int Default 0,"; //�Ƿ�ͨ��

            sql= sql + "isHtml YesNo Default No,"; //�Ƿ�Ϊhtml
            sql= sql + "isOnHtml YesNo Default Yes,"; //��������Html

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "noteBody Text Default \"\","; //�ʼ���Ϣ
            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //���Ա�-------------------------------
        tableName= db_PREFIX + "GuestBook";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "columnid VarChar Default \"\","; //��ĿID
            sql= sql + "parentId int Default -1,"; //����ĿID
            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "guestName VarChar Default \"\","; //����
            sql= sql + "tel VarChar Default \"\","; //�绰
            sql= sql + "fax VarChar Default \"\","; //����
            sql= sql + "EMail VarChar Default \"\","; //����
            sql= sql + "mobile VarChar Default \"\","; //�ֻ�
            sql= sql + "QQ VarChar Default \"\","; //QQ����
            sql= sql + "MSN VarChar Default \"\","; //MSN����
            sql= sql + "company VarChar Default \"\","; //��˾
            sql= sql + "address VarChar Default \"\","; //��ַ
            sql= sql + "postcode VarChar Default \"\","; //�ʱ�
            sql= sql + "ip VarChar Default \"\","; //IP��ַ
            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "customAUrl VarChar Default \"\","; //�Զ���������ַ
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ
            sql= sql + "fontColor VarChar Default \"\","; //����ɫ

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isThrough Int Default 0,"; //�Ƿ�ͨ��


            sql= sql + "reply Text Default \"\","; //�ظ�����
            sql= sql + "replyip Text Default \"\","; //�ظ�IP
            sql= sql + "replydatetime Text Default \"\","; //�ظ�ʱ��

            sql= sql + "noteBody Text Default \"\","; //�ʼ���Ϣ
            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //������-------------------------------
        tableName= db_PREFIX + "Feedback";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "columnid VarChar Default \"\","; //��ĿID
            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "feedbacktype VarChar Default \"\","; //��������
            sql= sql + "guestName VarChar Default \"\","; //����
            sql= sql + "tel VarChar Default \"\","; //�绰
            sql= sql + "fax VarChar Default \"\","; //����
            sql= sql + "EMail VarChar Default \"\","; //����
            sql= sql + "mobile VarChar Default \"\","; //�ֻ�
            sql= sql + "QQ VarChar Default \"\","; //QQ����
            sql= sql + "MSN VarChar Default \"\","; //MSN����
            sql= sql + "company VarChar Default \"\","; //��˾
            sql= sql + "address VarChar Default \"\","; //��ַ
            sql= sql + "postcode VarChar Default \"\","; //�ʱ�
            sql= sql + "ip VarChar Default \"\","; //IP��ַ
            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "customAUrl VarChar Default \"\","; //�Զ���������ַ
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ
            sql= sql + "fontColor VarChar Default \"\","; //����ɫ

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isThrough Int Default 0,"; //�Ƿ�ͨ��


            sql= sql + "reply Text Default \"\","; //�ظ�����
            sql= sql + "replyip Text Default \"\","; //�ظ�IP
            sql= sql + "replydatetime Text Default \"\","; //�ظ�ʱ��

            sql= sql + "noteBody Text Default \"\","; //�ʼ���Ϣ
            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //��̨�˵���-------------------------------
        tableName= db_PREFIX + "ListMenu";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "parentId int Default -1,"; //����ĿID
            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "lablename VarChar Default \"\","; //��ǩ����
            sql= sql + "customAUrl VarChar Default \"\","; //�Զ���������ַ

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isDisplay YesNo,"; //��ʾ������
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //����QQ��------------------------------
        tableName= db_PREFIX + "LineQQ";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "BigClassName VarChar Default \"\","; //��������
            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "QQ VarChar Default \"\","; //QQ����
            sql= sql + "IsOnlineChat YesNo Default No,"; //�Ƿ�Ϊ��������
            sql= sql + "IsAddFriend YesNo Default Yes,"; //�Ƿ�Ϊ��Ϊ����

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "bodycontent Text)"; //������ϸ����
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //ͶƱ��-------------------------------
        tableName= db_PREFIX + "LineVote";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "title VarChar Default \"\","; //ͶƱ����
            sql= sql + "Option1 VarChar Default \"\","; //ͶƱѡ��һ
            sql= sql + "Option2 VarChar Default \"\","; //ͶƱѡ���
            sql= sql + "Option3 VarChar Default \"\","; //ͶƱѡ����
            sql= sql + "Option4 VarChar Default \"\","; //ͶƱѡ����
            sql= sql + "Option5 VarChar Default \"\","; //ͶƱѡ����
            sql= sql + "Option6 VarChar Default \"\","; //ͶƱѡ����
            sql= sql + "Num1 Int,"; //ͶƱѡ����һ
            sql= sql + "Num2 Int,"; //ͶƱѡ������
            sql= sql + "Num3 Int,"; //ͶƱѡ������
            sql= sql + "Num4 Int,"; //ͶƱѡ������
            sql= sql + "Num5 Int,"; //ͶƱѡ������
            sql= sql + "Num6 Int,"; //ͶƱѡ������
            sql= sql + "isDisplay YesNo Default Yes,"; //ͶƱ��ʾ�ر�
            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����
            sql= sql + "folderName VarChar Default \"\","; //HTMl�ļ�������
            sql= sql + "fileName VarChar Default \"\","; //HTML�ļ�����
            sql= sql + "banner VarChar Default \"\","; //��ǰBannerͼƬ
            sql= sql + "templatePath VarChar Default \"\","; //ģ��·��
            sql= sql + "target VarChar Default \"\","; //���Ӵ򿪷�ʽ
            sql= sql + "fontColor VarChar Default \"\","; //����ɫ
            sql= sql + "fontB YesNo Default No,"; //�ִ�
            sql= sql + "OnHtml YesNo Default Yes,"; //��������Html

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "VoteType YesNo Default Yes)"; //ͶƱ����
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //��Ա��-------------------------------
        tableName= db_PREFIX + "Member";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "userType VarChar Default \"\","; //��Ա����
            sql= sql + "userName VarChar Default \"\","; //��Ա�˺�
            sql= sql + "pwd VarChar Default \"\","; //��Ա����
            sql= sql + "yunPwd VarChar Default \"\","; //ԭ����
            sql= sql + "sex VarChar Default '��',"; //�Ա�
            sql= sql + "age Int Default 18,"; //����
            sql= sql + "tel VarChar Default \"\","; //�绰
            sql= sql + "phone VarChar Default \"\","; //�ֻ�
            sql= sql + "fax VarChar Default \"\","; //����
            sql= sql + "email VarChar Default \"\","; //����
            sql= sql + "postcode VarChar Default \"\","; //�ʱ�
            sql= sql + "address VarChar Default \"\","; //��ַ
            sql= sql + "company VarChar Default \"\","; //��˾

            sql= sql + "regIP VarChar Default \"\","; //ע��IP
            sql= sql + "loginIp VarChar Default \"\","; //IP

            sql= sql + "loginCount Int Default 0,"; //��¼����
            sql= sql + "loginTime DateTime Default Date(),"; //��¼ʱ��
            sql= sql + "lastLoginTime DateTime Default Date(),"; //�ϴε�¼ʱ��



            //QQ��Ϣ
            sql= sql + "openId VarChar Default \"\","; //openId
            sql= sql + "accessToken VarChar Default \"\","; //accessToken
            sql= sql + "nickname VarChar Default \"\","; //nickname
            sql= sql + "qqphoto VarChar Default \"\","; //QQͷ��
            sql= sql + "useryear Int Default 18,"; //��һ�����


            sql= sql + "Province VarChar Default \"\","; //ʡ��
            sql= sql + "City VarChar Default \"\","; //����
            sql= sql + "Area VarChar Default \"\","; //����20150305



            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isThrough Int Default 0,"; //�Ƿ�ͨ��
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //��Ա��־��-------------------------------
        tableName= db_PREFIX + "MemberLog";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "userid Int Default -1,"; //��Աid
            sql= sql + "title VarChar Default \"\","; //����

            sql= sql + "ip VarChar Default \"\","; //IP
            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isThrough Int Default 0,"; //�Ƿ�ͨ��
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //������Ϣ��-------------------------------
        tableName= db_PREFIX + "Payment";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "UserName VarChar Default \"\","; //��Ա�˺�
            sql= sql + "MemberId VarChar Default \"\","; //��ԱID���
            sql= sql + "Sex VarChar Default '��',"; //�Ա�
            sql= sql + "Age Int Default 18,"; //����

            sql= sql + "Tel VarChar Default \"\","; //�绰
            sql= sql + "Phone VarChar Default \"\","; //�ֻ�
            sql= sql + "Fax VarChar Default \"\","; //����
            sql= sql + "Email VarChar Default \"\","; //����
            sql= sql + "postcode VarChar Default \"\","; //�ʱ�
            sql= sql + "Address VarChar Default \"\","; //��ַ
            sql= sql + "Company VarChar Default \"\","; //��˾
            sql= sql + "QQMSN VarChar Default \"\","; //��˾
            sql= sql + "Ip VarChar Default \"\","; //IP

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isThrough Int Default 0)"; //�Ƿ�ͨ��
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //������Ϣ��-------------------------------
        tableName= db_PREFIX + "Previeworder";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "MemberId VarChar Default \"\","; //��Ա����˺�
            sql= sql + "OrderId VarChar Default \"\","; //����ID���
            sql= sql + "ProductId VarChar Default \"\","; //��ƷID���
            sql= sql + "title VarChar Default '��',"; //��Ʒ����
            sql= sql + "Total VarChar Default \"\","; //Щ��Ʒ����
            sql= sql + "Price Int Default 18,"; //�˲�Ʒ����
            sql= sql + "ProductSum VarChar Default \"\","; //�˲�Ʒ��

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isThrough YesNo Default No)"; //�Ƿ񷢻�
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //��Ʒ���۱�-------------------------------
        tableName= db_PREFIX + "ProductComment";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "UserName VarChar Default \"\","; //������
            sql= sql + "title VarChar Default \"\","; //���۱���
            sql= sql + "Pid Int,"; //��ϢID
            sql= sql + "PTitle VarChar Default \"\","; //��Ϣ����
            sql= sql + "bodycontent Text Default \"\","; //��������
            sql= sql + "Sort Int Default 1,"; //��������
            sql= sql + "Ip VarChar Default \"\","; //������IP

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isThrough YesNo Default No)"; //�Ƿ���ʾ����
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //��վͳ�Ʊ� 20160203-------------------------------
        tableName= db_PREFIX + "WebSiteStat";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "visiturl Text Default \"\","; //������ַ
            sql= sql + "viewurl Text Default \"\","; //�����ַ
            sql= sql + "browser VarChar Default \"\","; //������汾
            sql= sql + "operatingSystem VarChar Default \"\","; //����ϵͳ
            sql= sql + "screenWH VarChar Default \"\","; //��Ļ���
            sql= sql + "moreInfo Text Default \"\","; //������Ϣ
            sql= sql + "viewDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "ip VarChar Default \"\","; //����IP

            sql= sql + "dateClass VarChar Default \"\","; //���ڷ���

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��



            sql= sql + "isThrough Int Default 0,"; //�Ƿ�ͨ��
            sql= sql + "noteInfo Text Default \"\","; //�ʼ���Ϣ
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //ϵͳ��־��-------------------------------
        tableName= db_PREFIX + "SystemLog";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "msgStr Text Default \"\","; //��Ϣ����
            sql= sql + "tableName VarChar Default \"\","; //���� Ҳ����˵moduleģ��
            sql= sql + "url Text Default \"\","; //����
            sql= sql + "adminId int Default 0,"; //����Աid
            sql= sql + "adminName VarChar Default 0,"; //����Ա����
            sql= sql + "IP VarChar Default \"\","; //IP��ַ


            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }


        //���ֱ� 20151228-------------------------------
        tableName= db_PREFIX + "webLayout";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "layoutName VarChar Default \"\","; //��������
            sql= sql + "layoutList VarChar Default \"\","; //�����б�

            sql= sql + "sourceStr VarChar Default \"\","; //���滻����
            sql= sql + "replaceStr VarChar Default \"\","; //ͳһ�滻����

            sql= sql + "actionContent Text Default \"\","; //��������

            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "isDisplay YesNo,"; //�Ƿ���ʾ
            sql= sql + "author VarChar Default \"\","; //����
            sql= sql + "views int Default 0,"; //�������
            sql= sql + "adminId int Default 0,"; //����Աid


            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��

            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //ģ��� 20151228-------------------------------
        tableName= db_PREFIX + "webModule";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "moduleType VarChar Default \"\","; //ģ������
            sql= sql + "moduleName VarChar Default \"\","; //ģ������
            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "isDisplay YesNo,"; //�Ƿ���ʾ
            sql= sql + "author VarChar Default \"\","; //����
            sql= sql + "views int Default 0,"; //�������
            sql= sql + "adminId int Default 0,"; //����Աid


            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��

            sql= sql + "aboutcontent Text Default \"\","; //�򵥽���
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //�ɼ���վ��-------------------------------
        tableName= db_PREFIX + "CaiWeb";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "BigClassName VarChar Default \"\","; //��������
            sql= sql + "columnname VarChar Default \"\","; //��Ŀ����
            sql= sql + "httpurl VarChar Default \"\","; //��ַ
            sql= sql + "morePageUrl VarChar Default \"\","; //����ҳ����  {*}        Ϊiҳ
            sql= sql + "charset VarChar Default \"\","; //�ַ�����
            sql= sql + "thisPage int Default 0,"; //��ǰҳ��
            sql= sql + "countPage int Default 0,"; //��ҳ��
            sql= sql + "sType VarChar Default \"\","; //����
            sql= sql + "sortRank int Default 0,"; //������

            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "IP VarChar Default \"\","; //IP��ַ
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //�ɼ����ñ�-------------------------------
        tableName= db_PREFIX + "CaiConfig";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "BigClassName VarChar Default \"\","; //��������
            sql= sql + "sType VarChar Default \"\","; //����
            sql= sql + "sAction VarChar Default \"\","; //����
            sql= sql + "startStr VarChar Default \"\","; //��ʼ��ȡ�ַ�
            sql= sql + "endStr VarChar Default \"\","; //������ȡ�ַ�
            sql= sql + "startAddStr VarChar Default \"\","; //��ʼ��ȡ�ַ�ǰ��׷���ַ�
            sql= sql + "endAddStr VarChar Default \"\","; //������ȡ�ַ���̨׷���ַ�
            sql= sql + "fieldName VarChar Default \"\","; //�ֶ�����
            sql= sql + "fieldCheck int Default 0,"; //�ֶμ��

            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "IP VarChar Default \"\","; //IP��ַ

            sql= sql + "isThrough YesNo Default No,"; //ͨ�����
            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //�ɼ����ݱ�-------------------------------
        tableName= db_PREFIX + "CaiData";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "BigClassName VarChar Default \"\","; //��������
            sql= sql + "columnname VarChar Default \"\","; //��Ŀ����

            sql= sql + "sType VarChar Default \"\","; //����

            sql= sql + "value1 Text Default \"\","; //����1
            sql= sql + "value2 Text Default \"\","; //����2
            sql= sql + "value3 Text Default \"\","; //����3
            sql= sql + "value4 Text Default \"\","; //����4
            sql= sql + "value5 Text Default \"\","; //����5
            sql= sql + "value6 Text Default \"\","; //����6

            sql= sql + "fieldName1 VarChar Default \"\","; //�ֶ�����1
            sql= sql + "fieldName2 VarChar Default \"\","; //�ֶ�����2
            sql= sql + "fieldName3 VarChar Default \"\","; //�ֶ�����3
            sql= sql + "fieldName4 VarChar Default \"\","; //�ֶ�����4
            sql= sql + "fieldName5 VarChar Default \"\","; //�ֶ�����5
            sql= sql + "fieldName6 VarChar Default \"\","; //�ֶ�����6

            sql= sql + "fieldCheck1 VarChar Default \"\","; //�ֶμ��1
            sql= sql + "fieldCheck2 VarChar Default \"\","; //�ֶμ��2
            sql= sql + "fieldCheck3 VarChar Default \"\","; //�ֶμ��3
            sql= sql + "fieldCheck4 VarChar Default \"\","; //�ֶμ��4
            sql= sql + "fieldCheck5 VarChar Default \"\","; //�ֶμ��5
            sql= sql + "fieldCheck6 VarChar Default \"\","; //�ֶμ��6

            sql= sql + "sortRank int Default 0,"; //������
            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "IP VarChar Default \"\","; //IP��ַ
            sql= sql + "isThrough YesNo Default No,"; //ͨ�����

            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //���۱�(�ٶȾ���) 20151228-------------------------------
        tableName= db_PREFIX + "Bidding";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "searchWords VarChar Default \"\","; //������
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "showReason VarChar Default \"\","; //չ������
            sql= sql + "nComputerSearch int Default 0,"; //����������
            sql= sql + "nMoblieSearch int Default 0,"; //�ƶ�������
            sql= sql + "nCountSearch int Default 0,"; //����������

            sql= sql + "nWordPrice int Default 0,"; //�´���ࣨ�Ϸ���ָ����
            sql= sql + "nDegree int Default 0,"; //�������ҳ̶�

            sql= sql + "sortRank Int Default 0,"; //˳��
            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��


            sql= sql + "bodycontent Text Default \"\")"; //��ע
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //��վURLɨ��� 20160428-------------------------------
        tableName= db_PREFIX + "WebUrlScan";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "bigClassName VarChar Default \"\","; //��������
            sql= sql + "linkType VarChar Default \"\","; //��������
            sql= sql + "website VarChar Default \"\","; //����
            sql= sql + "title VarChar Default \"\","; //����
            sql= sql + "httpurl VarChar Default \"\","; //��ַ
            sql= sql + "totitle VarChar Default \"\","; //���ñ���
            sql= sql + "tohttpurl VarChar Default \"\","; //������ַ
            sql= sql + "webState int Default 0,"; //״̬
            sql= sql + "openSpeed int Default 0,"; //��ʱ��
            sql= sql + "charset VarChar Default \"\","; //�ַ�����
            sql= sql + "content_type VarChar Default \"\","; //��������
            sql= sql + "link_count Int Default 0,"; //��������

            sql= sql + "webSize int Default 0,"; //��ҳ��С
            sql= sql + "sortRank Int Default 0,"; //˳��
            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��
            sql= sql + "isThrough YesNo Default No)"; //ͨ�����
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }
        //��վ������ 20160511-------------------------------
        tableName= db_PREFIX + "WebDomain";
        if( checkTable_show(tableName)== false ){
            sql= "Create Table " + tableName + " (Id int Identity(0,1) Primary Key,";
            sql= sql + "bigClassName VarChar Default \"\","; //��������
            sql= sql + "website VarChar Default \"\","; //����
            sql= sql + "webTitle VarChar Default \"\","; //��վ�Զ�����
            sql= sql + "webKeywords VarChar Default \"\","; //��վ�ؼ���
            sql= sql + "webDescription Text Default \"\","; //��վ����

            sql= sql + "webState int Default 0,"; //״̬
            sql= sql + "openSpeed int Default 0,"; //��ʱ��
            sql= sql + "charset VarChar Default \"\","; //�ַ�����
            sql= sql + "content_type VarChar Default \"\","; //��������
            sql= sql + "server_name VarChar Default \"\","; //����������


            sql= sql + "isAsp YesNo Default No,"; //������Asp����
            sql= sql + "isAspx YesNo Default No,"; //������Aspx����
            sql= sql + "isPhp YesNo Default No,"; //������Php����
            sql= sql + "isJsp YesNo Default No,"; //������Jsp����
            sql= sql + "isHtm YesNo Default No,"; //������Htm����
            sql= sql + "isHtml YesNo Default No,"; //������Html����
            sql= sql + "nlinks int Default 0,"; //������
            sql= sql + "links Text Default \"\","; //�����б�


            sql= sql + "homepagelist VarChar Default \"\","; //��

            sql= sql + "flags VarChar Default \"\","; //��
            sql= sql + "webSize int Default 0,"; //��ҳ��С
            sql= sql + "sortRank Int Default 0,"; //˳��
            sql= sql + "addDateTime DateTime Default Now(),"; //����ʱ��
            sql= sql + "upDateTime DateTime Default Date(),"; //�޸�ʱ��

            sql= sql + "isDomain YesNo Default No,"; //�Ƿ�Ϊ����
            sql= sql + "isThrough YesNo Default No)"; //ͨ�����
            if( MDBPath== "" ){ sql= handleSqlServer(sql) ;}//��Access���ݿ�����ת��SqlServer���ݿ�����
            connexecute(sql);
        }


        isAddData= true;
        if( isAddData== true ){
            //Add Admin
            connexecute("insert into " + db_PREFIX + "admin (username,pwd,flags) values('" + loginname + "','" + myMD5(loginpwd) + "','|*|')");



            //Add WebSite
            string fieldNameList=""; string fieldValueList="";
            fieldNameList= "webTitle,webKeywords,webDescription,WebSiteUrl,WebSiteBottom";
            fieldValueList= "'��վ����','��վ�ؼ���','��վ����','http://www.baidu.com','��Ȩ���� С������ �绰��021-66666666 ���棺021-8888888<br>��ϵQQ��313801120 ���䣺313801120@qq.com'";
            connexecute("insert into " + db_PREFIX + "website (" + fieldNameList + ") values(" + fieldValueList + ")");

            //add menu
            string parentid=""; string title=""; string lableName=""; string content=""; string tempS=""; string url=""; string sIsdisplay=""; int nCount=0;
            content= getFText(adminDir + "/��̨�˵�����.ini");
            content= replace(content, vbTab(), "    ");
            splStr= aspSplit(content, vbCrlf());
            nCount= 0;
            foreach(var eachs in splStr){
                s=eachs;
                tempS= s;
                s= aspTrim(s);
                if( tempS != "" && left(phpTrim(s) + " ", 1) != "#" ){
                    nCount= nCount + 1; //����
                    if( left(tempS, 4)== "    " ){
                    }else{
                        parentid= "-1";
                    }

                    if( lCase(tempS)== "end" ){
                        break;
                    }else if( s != "" ){
                        title= mid(s + " ", 1, inStr(s + " ", " ") - 1);
                        lableName= getStrCut(s, "lablename='", "'", 2);
                        url= getStrCut(s, "url='", "'", 2);
                        sIsdisplay= getStrCut(s, "isdisplay='", "'", 2);
                        if( sIsdisplay== "" ){
                            sIsdisplay= "1";
                        }
                        if( title != "" ){
                            echo("lablename", lableName);

                            connexecute("insert into " + db_PREFIX + "ListMenu (title,parentid,sortrank,lablename,isdisplay,customaurl) values('" + title + "'," + parentid + "," + nCount + ",'" + lableName + "'," + sIsdisplay + ",'" + url + "')");
                            if( parentid== "-1" ){
                                rs = new OleDbCommand("select * from " + db_PREFIX + "ListMenu where title='" + title + "'", conn).ExecuteReader();

                                if( rs.Read() ){
                                    parentid= cStr(rs["id"]);
                                }
                            }

                            echo(title, s);
                            echo(url, lableName);
                        }
                    }
                }
            }


        }

    }

    //�������ݿ�
    void createAccess(){
        string loginname=""; string loginpwd="";
        loginname= cStr(Request["loginname"]);
        loginpwd= cStr(Request["loginpwd"]);
        conn=openConn();
        conn.Open();



        webData(db_PREFIX, loginname, loginpwd); //������վ����
    }
    //��ʾĬ��
    void displayDefault(){

        showLayout="step1";
        //Ĭ������
        selectDatabase= "access";
        dbhost= "localhost";
        dbuser= "sa";
        dbpwd= "";
        dbname= "webdata";
        accessDir= "/data/";
        msgStr= "�������ݿ�";
        isYes= true;
        if( cStr(Request.Form["selectDatabase"]) != "" ){
            selectDatabase= cStr(Request.Form["selectDatabase"]);
            dbhost= cStr(Request.Form["dbhost"]);
            dbuser= cStr(Request.Form["dbuser"]);
            dbpwd= cStr(Request.Form["dbpwd"]);
            dbname= cStr(Request.Form["dbname"]);
            accessDir= phpTrim(cStr(Request.Form["accessDir"]));
            accessPath= accessDir + "/data.mdb";

            if( selectDatabase== "access" ){
                if( accessDir== "" ){
                    msgStr= "access���ݿⰲװĿ¼����Ϊ��";
                    isYes= false;
                }else if( checkFolder(accessDir)== false ){
                    msgStr= "access���ݿⰲװĿ¼������";
                    isYes= false;
                }else if( checkFile(accessPath)== true ){
                    msgStr= "access���ݿ���ڣ�ʹ�ã�" + accessPath + "��";
                    //isYes = false
                }else{
                    createMdb(accessPath); // �������ݿ�
                }

                //sqlserver ����
            }else{
                accessPath= "";
                if( EDITORTYPE=="aspx" ){
                    connStr= "server='"+ dbhost +",1433';database='"+ dbname +"';uid='"+ dbuser +"';pwd='"+ dbpwd +"';";
                }else{
                    connStr= " Password = " + dbpwd + "; user id =" + dbuser + "; Initial Catalog =" + dbname + "; data source =" + dbhost + ";Provider = sqloledb;";
                }
                if( checkSqlServer(connStr)== false ){
                    echo("connStr", connStr);
                    msgStr= "����SqlServer���ݿ�ʧ�ܣ�����˺�����";
                    isYes= false;
                }

            }

            //Ϊ���򴴽����ݿ�
            if( isYes== true ){
                content= getFText(configFilePath) ; tempContent= content;

                //���ݿ�
                startStr= "MDBPath = \"" ; endStr= "\"";
                if( inStr(content, startStr) > 0 && inStr(content, startStr) > 0 ){
                    findStr= getStrCut(content, startStr, endStr, 1);
                    if( EDITORTYPE=="aspx" ){
                        if( accessPath!="" ){
                            accessPath=handlePath(accessPath);
                        }
                        accessPath=replace(accessPath,"\\","\\\\");
                    }else{
                        accessPath=replace(accessPath,"\\\\","\\");
                    }
                    replaceStr= startStr + accessPath + endStr;
                    content= replace(content, findStr, replaceStr);
                }
                //���ݿ�
                startStr= "databaseType = \"" ; endStr= "\"";
                if( inStr(content, startStr) > 0 && inStr(content, startStr) > 0 ){
                    findStr= getStrCut(content, startStr, endStr, 1);
                    replaceStr= startStr + selectDatabase + endStr;
                    content= replace(content, findStr, replaceStr);
                }



                startStr= "sqlServerHostIP = \"" ; endStr= "\"";
                if( inStr(content, startStr) > 0 && inStr(content, startStr) > 0 ){
                    findStr= getStrCut(content, startStr, endStr, 1);
                    if( EDITORTYPE=="aspx" ){
                        dbhost=replace(dbhost,"\\","\\\\");
                    }
                    replaceStr= startStr + dbhost + endStr;
                    content= replace(content, findStr, replaceStr);
                }
                startStr= "sqlServerUsername = \"" ; endStr= "\"";
                if( inStr(content, startStr) > 0 && inStr(content, startStr) > 0 ){
                    findStr= getStrCut(content, startStr, endStr, 1);
                    replaceStr= startStr + dbuser + endStr;
                    content= replace(content, findStr, replaceStr);
                }
                startStr= "sqlServerPassword = \"" ; endStr= "\"";
                if( inStr(content, startStr) > 0 && inStr(content, startStr) > 0 ){
                    findStr= getStrCut(content, startStr, endStr, 1);
                    replaceStr= startStr + dbpwd + endStr;
                    content= replace(content, findStr, replaceStr);
                }
                startStr= "sqlServerDatabaseName = \"" ; endStr= "\"";
                if( inStr(content, startStr) > 0 && inStr(content, startStr) > 0 ){
                    findStr= getStrCut(content, startStr, endStr, 1);
                    replaceStr= startStr + dbname + endStr;
                    content= replace(content, findStr, replaceStr);
                }

                //���������ļ�
                if( tempContent != content ){
                    createFile(configFilePath, content);
                }

                showLayout="step2";
            }
        }
    }
    //��ʾ��������
    void displayImportTableLayout(){
        string content=""; string startStr=""; string endStr=""; string findStr=""; string replaceStr="";
        content= getFText(configFilePath);
        startStr= "db_PREFIX = \"";
        endStr= "\"";
        if( inStr(content, startStr) > 0 && inStr(content, startStr) > 0 ){
            findStr= getStrCut(content, startStr, endStr, 1);
            replaceStr= startStr + cStr(Request["db_PREFIX"]) + endStr;
            content= replace(content, findStr, replaceStr);
            createFile(configFilePath, content);
        }

        //ɾ�������ļ�
        WEB_CACHEFile= replace(replace(WEB_CACHEFile, "[adminDir]", adminDir), "[EDITORTYPE]", EDITORTYPE);
        //call echo("WEB_CACHEFile",WEB_CACHEFile)
        deleteFile(WEB_CACHEFile);
        showLayout="step3";
    }
    //��ʾ��������
    void displayImportData(){
        Session["adminusername"]= "PAAJCMS";
        Session["adminflags"]= "|*|";
        showLayout="step4";
    }
    </script>
    <%if( showLayout=="step1" ){%>
        <form id="form1" name="form1" method="post" action="?act=displayDefault">
        <div class="pright">
        <div class="pr-title"><h3>��һ�� ѡ�����ݿ�</h3></div>
        <label for="radio"><input name="selectDatabase" type="radio" id="radio" value="access" <%=IIF(selectDatabase== "access","checked","")%> onClick="selectInsrtDatabase(this.value);" />
        Access���ݿ�</label>
        <label for="radio2"><input name="selectDatabase" type="radio" id="radio2" value="sqlserver" <%=IIF(selectDatabase== "sqlserver","checked","")%> onClick="selectInsrtDatabase(this.value);" />
        SqlServer���ݿ�</label>
        <table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
        <tbody>
        <tr>
        <td class="onetd"><strong>��ʾ��</strong></td>
        <td> <div class="mytipwrap"><%=msgStr%> &nbsp;<a href="<%=WEB_VIEWURL%>" class="mytip" target="_blank">������վ��ҳ</a><a href="<%=WEB_ADMINURL%>" class="mytip" target="_blank">��¼��վ��̨</a> </div> </td>
        </tr>
        </tbody>
        <tbody id="access_layout" <%=IIF(selectDatabase== "sqlserver","style='display:none'","")%>>
        <tr>
        <td class="onetd"><strong>AccessĿ¼��</strong></td>
        <td>
        <div style="float:left;margin-right:3px;"><input name="accessDir" id="accessDir" type="text" value="<%=accessDir%>" class="input-txt" >
        </div>
        <div style="float:left" id="havedbsta"><font color="red"></font></div> </td>
        </tr>
        </tbody>

        <tbody id="sqlserver_layout" <%=IIF(selectDatabase== "access","style='display:none'","")%> >
        <tr>
        <td class="onetd"><strong>���ݿ�������</strong></td>
        <td><input name="dbhost" id="dbhost" type="text" value="<%=dbhost%>" class="input-txt">
        <small>һ��Ϊlocalhost / .\SQLEXPRESS</small> </td>
        </tr>
        <tr>
        <td class="onetd"><strong>���ݿ��û���</strong></td>
        <td><input name="dbuser" id="dbuser" type="text" value="<%=dbuser%>" class="input-txt"></td>
        </tr>
        <tr>
        <td class="onetd"><strong>���ݿ����룺</strong></td>
        <td>
        <div style="float:left;margin-right:3px;"><input name="dbpwd" type="text" value="<%=dbpwd%>" class="input-txt" id="dbpwd" onChange="TestDb()" >
        </div>
        <div style="float:left" id="dbpwdsta"><font color="red"></font></div> </td>
        </tr>
        <tr>
        <td class="onetd"><strong>���ݿ����ƣ�</strong></td>
        <td>
        <div style="float:left;margin-right:3px;"><input name="dbname" id="dbname" type="text" value="<%=dbname%>" class="input-txt" onChange="HaveDB()">
        </div>
        <div style="float:left" id="havedbsta"><font color="red"></font></div> </td>
        </tr>
        </tbody>

        </table>
        <div class="btn-box">
        <input name="�ύ" type="submit" class="btnclick1" onClick="window.location.href='index.php?step=3';" value="����">
        </div>
        </div>
        <script language="javascript">
        function selectInsrtDatabase(sValue){
            //alert(document.getElementById("sqlserver_layout").innerHTML)
            if(sValue=="access"){
                document.getElementById("access_layout").style.display = ""
                document.getElementById("sqlserver_layout").style.display="none";
            }else{
                document.getElementById("access_layout").style.display="none";
                document.getElementById("sqlserver_layout").style.display="";
            }
        }
        </script>
        </form>
        <%;}else if( showLayout=="step2" ){%>
            <form id="form1" name="form1" method="post" action="?act=displayImportTableLayout">
            <div class="pright">
            <div class="pr-title"><h3>�ڶ��� ���õ�¼����</h3></div>
            <table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
            <tbody>
            <tr>
            <td class="onetd"><strong>��ʾ��</strong></td>
            <td> <div class="mytipwrap">�������ݿ�ɹ���(������ָ����ݿ⽫�˳�) &nbsp;<a href="<%=WEB_VIEWURL%>" class="mytip" target="_blank">������վ��ҳ</a><a href="<%=WEB_ADMINURL%>" class="mytip" target="_blank">��¼��վ��̨</a> </div> </td>
            </tr>
            <tr>
            <td class="onetd"><strong>��ǰ׺��</strong></td>
            <td>
            <div style="float:left;margin-right:3px;"><input name="db_PREFIX" id="db_PREFIX" type="text" value="xy_" class="input-txt">
            </div>
            </td>
            </tr>
            <tr>
            <td class="onetd"><strong>��¼�˺ţ�</strong></td>
            <td>
            <div style="float:left;margin-right:3px;"><input name="loginname" type="text" value="admin" class="input-txt">
            </div>
            </td>
            </tr>
            <tr>
            <td class="onetd"><strong>��¼���룺</strong></td>
            <td>
            <div style="float:left;margin-right:3px;"><input name="loginpwd" type="text" value="admin" class="input-txt">
            </div>
            </td>
            </tr>
            </tbody>
            </table>
            <div class="btn-box">
            <input name="�ύ" type="submit" class="btnclick1" value="����">
            </div>
            </div>
            </form>
            <%;}else if( showLayout=="step3" ){%>
                <form id="form1" name="form1" method="post" action="?act=displayImportData">
                <div class="pright">
                <div class="pr-title"><h3>������ ��������</h3></div>
                <table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
                <tbody>
                <tr>
                <td class="onetd"><strong>��ʾ��</strong></td>
                <td> <div class="mytipwrap">��������ɡ�(������ָ����ݿ⽫�˳�) &nbsp;<a href="<%=WEB_VIEWURL%>" class="mytip" target="_blank">������վ��ҳ</a><a href="<%=WEB_ADMINURL%>" class="mytip" target="_blank">��¼��վ��̨</a> </div> </td>
                </tr>
                <tr>
                <td colspan="2">
                <iframe src="?act=createAccess&loginname=<%=cStr(Request["loginname"])%>&loginpwd=<%=cStr(Request["loginpwd"])%>" width="100%" height="350"></iframe>
                </td>
                </tr>
                </tbody>
                </table>
                <div class="btn-box">
                <input name="�ύ" type="submit" class="btnclick1" value="����">
                </div>
                </div>
                </form>
                <%;}else if( showLayout=="step4" ){%>
                    <div class="pright">
                    <div class="pr-title"><h3>���Ĳ� ���</h3></div>
                    <table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
                    <tbody>
                    <tr>
                    <td class="onetd"><strong>��ʾ��</strong></td>
                    <td> <div class="mytipwrap">����������� &nbsp;<a href="<%=WEB_VIEWURL%>" class="mytip" target="_blank">������վ��ҳ</a><a href="<%=WEB_ADMINURL%>" class="mytip" target="_blank">��¼��վ��̨</a> </div> </td>
                    </tr>
                    <tr>
                    <td colspan="2">
                    <iframe src="<%=WEB_ADMINURL%>?act=setAccess&webdataDir=/Templates2015/sharembweb/WebData&login=out" width="100%" height="350"></iframe>
                    </td>
                    </tr>
                    </tbody>
                    </table>
                    </div>
                    <% }%>
                </body>
                </html>
