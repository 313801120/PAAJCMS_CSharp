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
<!--#include File="Inc_csharp/Include/asp.aspx"-->
<!--#include File="Inc_csharp/Include/Conn.aspx"--> 
<!--#include File="Inc_csharp/Include/sys_FSO.aspx"-->
<!--#include File="Inc_csharp/Include/sys_System.aspx"--> 
<!--#include File="Inc_csharp/Include/sys_SessionCookie.aspx"--> 
<!--#include File="Inc_csharp/Include/sys_http.aspx"-->  
<!--#include File="Inc_csharp/Include/sys_patch.aspx"-->  


<!--#include File="Inc_csharp/Inc/2014_Array.aspx"--> 
<!--#include File="Inc_csharp/Inc/2014_Author.aspx"--> 
<!--#include File="Inc_csharp/Inc/2014_Css.aspx"-->  
<!--#include File="Inc_csharp/Inc/2014_Js.aspx"--> 
<!--#include File="Inc_csharp/Inc/2014_Template.aspx"-->
<!--#include File="Inc_csharp/Inc/2015_APGeneral.aspx"--> 
<!--#include File="Inc_csharp/Inc/2015_Color.aspx"--> 
<!--#include File="Inc_csharp/Inc/2015_Editor.aspx"--> 
<!--#include File="Inc_csharp/Inc/2015_Formatting.aspx"--> 
<!--#include File="Inc_csharp/Inc/2015_NewWebFunction.aspx"--> 
<!--#include File="Inc_csharp/Inc/2015_Param.aspx"--> 
<!--#include File="Inc_csharp/Inc/2015_ToMyPHP.aspx"--> 
<!--#include File="Inc_csharp/Inc/2016_Log.aspx"--> 
<!--#include File="Inc_csharp/Inc/2016_SaveData.aspx"--> 
<!--#include File="Inc_csharp/Inc/2016_WebControl.aspx"--> 
<!--#include File="Inc_csharp/Inc/ASPPHPAccess.aspx"--> 
<!--#include File="Inc_csharp/Inc/Cai.aspx"--> 
<!--#include File="Inc_csharp/Inc/Check.aspx"--> 
<!--#include File="Inc_csharp/Inc/Common.aspx"--> 
<!--#include File="Inc_csharp/Inc/EncDec.aspx"--> 
<!--#include File="Inc_csharp/Inc/function.aspx"--> 
<!--#include File="Inc_csharp/Inc/FunHTML.aspx"--> 
<!--#include File="Inc_csharp/Inc/Html.aspx"--> 
<!--#include File="Inc_csharp/Inc/IE.aspx"--> 
<!--#include File="Inc_csharp/Inc/Incpage.aspx"--> 
<!--#include File="Inc_csharp/Inc/PinYin.aspx"--> 
<!--#include File="Inc_csharp/Inc/Print.aspx"--> 
<!--#include File="Inc_csharp/Inc/StringNumber.aspx"--> 
<!--#include File="Inc_csharp/Inc/SystemInfo.aspx"--> 
<!--#include File="Inc_csharp/Inc/Time.aspx"--> 
<!--#include File="Inc_csharp/Inc/URL.aspx"--> 
<!--#include File="Inc_csharp/Inc/2014_GBUTF.aspx"--> 
<!--#include File="Inc_csharp/Inc/2014_Form.aspx"-->   
<!--#include File="Inc_csharp/config.aspx"-->

<!--#include File="Inc_csharp/function.aspx"--> 
<!--#include File="Inc_csharp/setAccess.aspx"-->
 
<script runat="server" language="c#">
protected void Page_Load(object sender, EventArgs e){

	loadRun();	
	//�ر����ݿ�
	if (conn.State != ConnectionState.Closed){
		conn.Close(); 
	}
	
} 



//=========


//������   ReplaceValueParamΪ�����ַ���ʾ��ʽ
string handleAction(string content){
    string startStr=""; string endStr=""; string actionList=""; string[] splStr; string action=""; string s=""; bool isHnadle;
    startStr= "{$" ; endStr= "$}";
    actionList= getArrayNew(content, startStr, endStr, true, true);
    //Call echo("ActionList ", ActionList)
    splStr= aspSplit(actionList, "$Array$");
    foreach(var eachs in splStr){
        s=eachs;
        action= aspTrim(s);
        action= handleInModule(action, "start"); //����\'�滻��
        if( action != "" ){
            action= aspTrim(mid(action, 3, len(action) - 4)) + " ";
            //call echo("s",s)
            isHnadle= true; //����Ϊ��
            //{VB #} �����Ƿ���ͼƬ·���Ŀ����Ϊ����VB�ﲻ�������·��
            if( checkFunValue(action, "# ")== true ){
                action= "";
                //����
            }else if( checkFunValue(action, "GetLableValue ")== true ){
                action= XY_getLableValue(action);
                //�����������������б�
            }else if( checkFunValue(action, "TitleInSearchEngineList ")== true ){
                action= XY_TitleInSearchEngineList(action);

                //�����ļ�
            }else if( checkFunValue(action, "Include ")== true ){
                action= XY_Include(action);
                //��Ŀ�б�
            }else if( checkFunValue(action, "ColumnList ")== true ){
                action= XY_AP_ColumnList(action);
                //�����б�
            }else if( checkFunValue(action, "ArticleList ")== true || checkFunValue(action, "CustomInfoList ")== true ){
                action= XY_AP_ArticleList(action);
                //�����б�
            }else if( checkFunValue(action, "CommentList ")== true ){
                action= XY_AP_CommentList(action);
                //����ͳ���б�
            }else if( checkFunValue(action, "SearchStatList ")== true ){
                action= XY_AP_SearchStatList(action);
                //���������б�
            }else if( checkFunValue(action, "Links ")== true ){
                action= XY_AP_Links(action);

                //��ʾ��ҳ����
            }else if( checkFunValue(action, "GetOnePageBody ")== true || checkFunValue(action, "MainInfo ")== true ){
                action= XY_AP_GetOnePageBody(action);
                //��ʾ��������
            }else if( checkFunValue(action, "GetArticleBody ")== true ){
                action= XY_AP_GetArticleBody(action);
                //��ʾ��Ŀ����
            }else if( checkFunValue(action, "GetColumnBody ")== true ){
                action= XY_AP_GetColumnBody(action);

                //�����ĿURL
            }else if( checkFunValue(action, "GetColumnUrl ")== true ){
                action= XY_GetColumnUrl(action);
                //�����ĿID
            }else if( checkFunValue(action, "GetColumnId ")== true ){
                action= XY_GetColumnId(action);
                //�������URL
            }else if( checkFunValue(action, "GetArticleUrl ")== true ){
                action= XY_GetArticleUrl(action);
                //��õ�ҳURL
            }else if( checkFunValue(action, "GetOnePageUrl ")== true ){
                action= XY_GetOnePageUrl(action);
                //��õ���URL
            }else if( checkFunValue(action, "GetNavUrl ")== true ){
                action= XY_GetNavUrl(action);



                //------------------- ģ��ģ���� -----------------------
                //��ʾ������ ���ò���
            }else if( checkFunValue(action, "DisplayWrap ")== true ){
                action= XY_DisplayWrap(action);
                //��ʾ����
            }else if( checkFunValue(action, "Layout ")== true ){
                action= XY_Layout(action);
                //��ʾģ��
            }else if( checkFunValue(action, "Module ")== true ){
                action= XY_Module(action);
                //��ģ������
            }else if( checkFunValue(action, "ReadTemplateModule ")== true ){
                action= XY_ReadTemplateModule(action);
                //�������ģ�� 20150108
            }else if( checkFunValue(action, "GetContentModule ")== true ){
                action= XY_ReadTemplateModule(action);
                //��ģ����ʽ�����ñ���������   ������и���ĿStyle��������
            }else if( checkFunValue(action, "ReadColumeSetTitle ")== true ){
                action= XY_ReadColumeSetTitle(action);


                //------------------- ������ -----------------------
                //��ʾJS��ȾASP/PHP/VB�ȳ���ı༭��
            }else if( checkFunValue(action, "displayEditor ")== true ){
                action= displayEditor(action);
                //Js����վͳ��
            }else if( checkFunValue(action, "JsWebStat ")== true ){
                action= XY_JsWebStat(action);

                //------------------- ������ -----------------------
                //��ͨ����A
            }else if( checkFunValue(action, "HrefA ")== true ){
                action= XY_HrefA(action);
                //��Ŀ�˵�(���ú�̨��Ŀ����)
            }else if( checkFunValue(action, "ColumnMenu ")== true ){
                action= XY_AP_ColumnMenu(action);

                //------------------- ѭ������ -----------------------
                //Forѭ������
            }else if( checkFunValue(action, "ForArray ")== true ){
                action= XY_ForArray(action);

                //------------------- ������ -----------------------
                //��վ�ײ�
            }else if( checkFunValue(action, "WebSiteBottom ")== true || checkFunValue(action, "WebBottom ")== true ){
                action= XY_AP_WebSiteBottom(action);
                //��ʾ��վ��Ŀ 20160331
            }else if( checkFunValue(action, "DisplayWebColumn ")== true ){
                action= XY_DisplayWebColumn(action);
                //URL����
            }else if( checkFunValue(action, "escape ")== true ){
                action= XY_escape(action);
                //URL����
            }else if( checkFunValue(action, "unescape ")== true ){
                action= XY_unescape(action);
                //asp��php�汾
            }else if( checkFunValue(action, "EDITORTYPE ")== true ){
                action= XY_EDITORTYPE(action);

                //�����ַ
            }else if( checkFunValue(action, "getUrl ")== true ){
                action= XY_getUrl(action);

                //����λ����ʾ��Ϣ{}Ϊ�ж�����
            }else if( checkFunValue(action, "detailPosition ")== true ){
                action= XY_detailPosition(action);




                //��ʱ������
            }else if( checkFunValue(action, "copyTemplateMaterial ")== true ){
                action= "";
            }else if( checkFunValue(action, "clearCache ")== true ){
                action= "";
            }else{
                isHnadle= false; //����Ϊ��
            }
            //ע���������е�����ʾ �� And IsNul(action)=False
            if( isNul(action)== true ){ action= "" ;}
            if( isHnadle== true ){
                content= replace(content, s, action);
            }
        }
    }
    return content;
}

//��ʾ��վ��Ŀ �°� ��֮ǰ��վ ��������Ľ�������
string XY_DisplayWebColumn(string action){
    int i=0; string c=""; string s=""; string url=""; string sql=""; string dropDownMenu=""; string focusType=""; string addSql="";
    bool isConcise; //�����ʾ20150212
    string styleId=""; string styleValue=""; //��ʽID����ʽ����
    string cssNameAddId="";
    string shopnavidwrap=""; //�Ƿ���ʾ��ĿID��

    styleId= phpTrim(rParam(action, "styleID"));
    styleValue= phpTrim(rParam(action, "styleValue"));
    addSql= phpTrim(rParam(action, "addSql"));
    shopnavidwrap= phpTrim(rParam(action, "shopnavidwrap"));
    //If styleId <> "" Then
    //Call ReadNavCSS(styleId, styleValue)
    //End If

    //Ϊ�������� ���Զ���ȡ��ʽ����  20150615
    if( checkStrIsNumberType(styleValue) ){
        cssNameAddId= "_" + styleValue; //Css����׷��Id���
    }
    sql= "select * from " + db_PREFIX + "webcolumn";
    //׷��sql
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    if( checkSql(sql)== false ){ eerr("Sql", sql) ;}
    rs = new OleDbCommand(sql, conn).ExecuteReader();

    dropDownMenu= lCase(rParam(action, "DropDownMenu"));
    focusType= lCase(rParam(action, "FocusType"));
    isConcise= IIF(lCase(rParam(action, "isConcise"))== "true", false, true);

    if( isConcise== true ){ c= c + copyStr(" ", 4) + "<li class=left></li>" + vbCrlf() ;}
    for( i= 1 ; i<= rsRecordcount(sql); i++){


        url= getColumnUrl(cStr(rs["columnname"]), "name");
        if( cStr(rs["columnname"])== glb_columnName ){
            if( focusType== "a" ){
                s= copyStr(" ", 8) + "<li class=focus><a href=\"" + url + "\">" + cStr(rs["columname"]) + "</a>";
            }else{
                s= copyStr(" ", 8) + "<li class=focus>" + cStr(rs["columnname"]);
            }
        }else{
            s= copyStr(" ", 8) + "<li><a href=\"" + url + "\">" + cStr(rs["columnname"]) + "</a>";
        }
        //��վ��Ŀû��pageλ�ô���
        url= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);
        s= handleDisplayOnlineEditDialog(url, s, "", "div|li|span"); //�����Ƿ���������޸Ĺ�����

        c= c + s;

        //С��

        c= c + copyStr(" ", 8) + "</li>" + vbCrlf();

        if( isConcise== true ){ c= c + copyStr(" ", 8) + "<li class=line></li>" + vbCrlf() ;}
    }
    if( isConcise== true ){ c= c + copyStr(" ", 8) + "<li class=right></li>" + vbCrlf() ;}

    if( styleId != "" ){
        c= "<ul class='nav" + styleId + cssNameAddId + "'>" + vbCrlf() + c + vbCrlf() + "</ul>" + vbCrlf();
    }
    if( shopnavidwrap== "1" || shopnavidwrap== "true" ){
        c= "<div id='nav" + styleId + cssNameAddId + "'>" + vbCrlf() + c + vbCrlf() + "</div>" + vbCrlf();
    }

    return c;
}

//�滻ȫ�ֱ��� {$cfg_websiteurl$}
string replaceGlobleVariable( string content){
    content= handleRGV(content, "{$cfg_webSiteUrl$}", cfg_webSiteUrl); //��ַ
    content= handleRGV(content, "{$cfg_webTemplate$}", cfg_webTemplate); //ģ��
    content= handleRGV(content, "{$cfg_webImages$}", cfg_webImages); //ͼƬ·��
    content= handleRGV(content, "{$cfg_webCss$}", cfg_webCss); //css·��
    content= handleRGV(content, "{$cfg_webJs$}", cfg_webJs); //js·��
    content= handleRGV(content, "{$cfg_webTitle$}", cfg_webTitle); //��վ����
    content= handleRGV(content, "{$cfg_webKeywords$}", cfg_webKeywords); //��վ�ؼ���
    content= handleRGV(content, "{$cfg_webDescription$}", cfg_webDescription); //��վ����

    content= handleRGV(content, "{$cfg_webSiteBottom$}", cfg_webSiteBottom); //��վ�ײ�����

    content= handleRGV(content, "{$glb_columnId$}", glb_columnId); //��ĿId
    content= handleRGV(content, "{$glb_columnName$}", glb_columnName); //��Ŀ����
    content= handleRGV(content, "{$glb_columnType$}", glb_columnType); //��Ŀ����
    content= handleRGV(content, "{$glb_columnENType$}", glb_columnENType); //��ĿӢ������

    content= handleRGV(content, "{$glb_Table$}", glb_table); //��
    content= handleRGV(content, "{$glb_Id$}", glb_id); //id

    content= handleRGV(content, "[$ģ��Ŀ¼$]", "Module/"); //Module


    //���ݾɰ汾 ��������ȥ��
    content= handleRGV(content, "{$WebImages$}", cfg_webImages); //ͼƬ·��
    content= handleRGV(content, "{$WebCss$}", cfg_webCss); //css·��
    content= handleRGV(content, "{$WebJs$}", cfg_webJs); //js·��
    content= handleRGV(content, "{$Web_Title$}", cfg_webTitle);
    content= handleRGV(content, "{$Web_KeyWords$}", cfg_webKeywords);
    content= handleRGV(content, "{$Web_Description$}", cfg_webDescription);


    content= handleRGV(content, "{$EDITORTYPE$}", EDITORTYPE); //��׺
    content= handleRGV(content, "{$WEB_VIEWURL$}", WEB_VIEWURL); //��ҳ��ʾ��ַ
    //�����õ�
    content= handleRGV(content, "{$glb_articleAuthor$}", glb_articleAuthor); //��������
    content= handleRGV(content, "{$glb_articleAdddatetime$}", glb_articleAdddatetime); //�������ʱ��
    content= handleRGV(content, "{$glb_articlehits$}", glb_articlehits); //�������ʱ��

    content= handleRGV(content, "{$glb_upArticle$}", glb_upArticle); //��һƪ����
    content= handleRGV(content, "{$glb_downArticle$}", glb_downArticle); //��һƪ����
    content= handleRGV(content, "{$glb_aritcleRelatedTags$}", glb_aritcleRelatedTags); //���±�ǩ��
    content= handleRGV(content, "{$glb_aritcleBigImage$}", glb_aritcleBigImage); //���´�ͼ
    content= handleRGV(content, "{$glb_aritcleSmallImage$}", glb_aritcleSmallImage); //����Сͼ
    content= handleRGV(content, "{$glb_searchKeyWord$}", glb_searchKeyWord); //��ҳ��ʾ��ַ


    return content;
}

//�����滻
string handleRGV( string content, string findStr, string replaceStr){
    string lableName="";
    //��[$$]����
    lableName= mid(findStr, 3, len(findStr) - 4) + " ";
    lableName= mid(lableName, 1, inStr(lableName, " ") - 1);
    content= replaceValueParam(content, lableName, replaceStr);
    content= replaceValueParam(content, lCase(lableName), replaceStr);
    //ֱ���滻{$$}���ַ�ʽ������֮ǰ��վ
    content= replace(content, findStr, replaceStr);
    content= replace(content, lCase(findStr), replaceStr);
    return content;
}

//������վ������Ϣ
void loadWebConfig(){
    string templatedir="";
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "website", conn).ExecuteReader();

    if( rs.Read() ){
        cfg_webSiteUrl= phpTrim(cStr(rs["websiteurl"])); //��ַ
        cfg_webTemplate= webDir + phpTrim(cStr(rs["webtemplate"])); //ģ��·��
        cfg_webImages= webDir + phpTrim(cStr(rs["webimages"])); //ͼƬ·��
        cfg_webCss= webDir + phpTrim(cStr(rs["webcss"])); //css·��
        cfg_webJs= webDir + phpTrim(cStr(rs["webjs"])); //js·��
        cfg_webTitle= cStr(rs["webtitle"]); //��ַ����
        cfg_webKeywords= cStr(rs["webkeywords"]); //��վ�ؼ���
        cfg_webDescription= cStr(rs["webdescription"]); //��վ����
        cfg_webSiteBottom= cStr(rs["websitebottom"]); //��վ�ص�
        cfg_flags= cStr(rs["flags"]); //��

        //�Ļ�ģ��20160202
        if( cStr(Request["templatedir"]) != "" ){
            //ɾ������Ŀ¼ǰ���Ŀ¼������Ҫ�Ǹ�����20160414
            templatedir= replace(handlePath(cStr(Request["templatedir"])), handlePath("/"), "/");
            //call eerr("templatedir",templatedir)

            if((inStr(templatedir, ":") > 0 || inStr(templatedir, "..") > 0) && getIP() != "127.0.0.1" ){
                eerr("��ʾ", "ģ��Ŀ¼�зǷ��ַ�");
            }
            templatedir= handleHttpUrl(replace(templatedir, handlePath("/"), "/"));

            cfg_webImages= replace(cfg_webImages, cfg_webTemplate, templatedir);
            cfg_webCss= replace(cfg_webCss, cfg_webTemplate, templatedir);
            cfg_webJs= replace(cfg_webJs, cfg_webTemplate, templatedir);
            cfg_webTemplate= templatedir;
        }
    }
}

//��վλ�� ������
string thisPosition(string content){
    string c="";
    c= "<a href=\"" + getColumnUrl("��ҳ", "type") + "\">��ҳ</a>";
    if( glb_columnName != "" ){
        c= c + " >> <a href=\"" + getColumnUrl(glb_columnName, "name") + "\">" + glb_columnName + "</a>";
    }
    //20160330
    if( glb_locationType== "detail" ){
        c= c + " >> �鿴����";
    }
    //β��׷������
    c= c + positionEndStr;

    //call echo("glb_locationType",glb_locationType)

    content= replace(content, "[$detailPosition$]", c);
    content= replace(content, "[$detailTitle$]", glb_detailTitle);
    content= replace(content, "[$detailContent$]", glb_bodyContent);

    return content;
}

//��ʾ�����б�
string getDetailList(string action, string content, string actionName, string lableTitle, string fieldNameList, int nPageSize, string sPage, string addSql){
    conn=openConn();
    conn.Open();

    string defaultStr=""; int i=0; string s=""; string c=""; string tableName=""; int j=0; string[] splxx; string sql=""; int nPage=0;
    int nX=0; string url=""; int nCount=0;
    string pageInfo=""; int nModI=0; string startStr=""; string endStr="";

    string fieldName=""; //�ֶ�����
    string[] splFieldName; //�ָ��ֶ�

    string replaceStr=""; //�滻�ַ�
    tableName= lCase(actionName); //������
    string listFileName=""; //�б��ļ�����
    listFileName= rParam(action, "listFileName");
    string abcolorStr=""; //A�Ӵֺ���ɫ
    string atargetStr=""; //A���Ӵ򿪷�ʽ
    string atitleStr=""; //A���ӵ�title20160407
    string anofollowStr=""; //A���ӵ�nofollow

    string id=""; string idPage="";
    id= rq("id");
    checkIDSQL(cStr(Request["id"]));

    if( fieldNameList== "*" ){
        fieldNameList= getHandleFieldList(db_PREFIX + tableName, "�ֶ��б�");
    }

    fieldNameList= specialStrReplace(fieldNameList); //�����ַ�����
    splFieldName= aspSplit(fieldNameList, ","); //�ֶηָ������


    defaultStr= getStrCut(content, "<!--#body start#-->", "<!--#body end#-->", 2);



    pageInfo= getStrCut(content, "[page]", "[/page]", 1);
    if( pageInfo != "" ){
        content= replace(content, pageInfo, "");
    }
    //call eerr("pageInfo",pageInfo)

    sql= "select * from " + db_PREFIX + tableName + " " + addSql;
    //���SQL
    if( checkSql(sql)== false ){
        errorLog("������ʾ��<br>sql=" + sql + "<br>");
        return "";
    }
    rs = new OleDbCommand(sql, conn).ExecuteReader();


    nCount= rsRecordcount(sql);

    //Ϊ��̬��ҳ��ַ
    if( isMakeHtml== true ){
        url= "";
        if( len(listFileName) > 5 ){
            url= mid(listFileName, 1, len(listFileName) - 5) + "[id].html";
            url= urlAddHttpUrl(cfg_webSiteUrl, url);
        }
    }else{
        url= getUrlAddToParam(getUrl(), "?page=[id]", "replace");
    }
    content= replace(content, "[$pageInfo$]", webPageControl(nCount, nPageSize, cStr(sPage), url, pageInfo));

    if( EDITORTYPE== "asp" ){



    }else if( EDITORTYPE== "aspx" ){

        int nCountPage = getCountPage(nCount, nPageSize);
        if(nPage<=1){
            nX=nPageSize;
            if(nX>nCount){
                nX=nCount;
            }
        }else{
            for(int nI2=0;nI2<nPageSize*(nPage-1);nI2++){
                rs.Read();
            }
            if(nPage<nCountPage){
                nX=nPageSize;
            }else{
                nX=nCount-nPageSize*(nPage-1);
            }
        }

    }else{




        rs = new OleDbCommand(sql, conn).ExecuteReader();



    }
    //call echo("sql",sql)
    for( i= 1 ; i<= nX; i++){

        rs.Read();
        startStr= "[list-" + i + "]" ; endStr= "[/list-" + i + "]";

        //�����ʱ����ǰ����20160202
        if( i== nX ){
            startStr= "[list-end]" ; endStr= "[/list-end]";
        }

        //��[list-mod2]  [/list-mod2]    20150112
        for( nModI= 6 ; nModI>= 2 ; nModI--){
            if( inStr(defaultStr, startStr)== 0 && i % nModI== 0 ){
                startStr= "[list-mod" + nModI + "]" ; endStr= "[/list-mod" + nModI + "]";
                if( inStr(defaultStr, startStr) > 0 ){
                    break;
                }
            }
        }

        //û������Ĭ��
        if( inStr(defaultStr, startStr)== 0 || startStr== "" ){
            startStr= "[list]" ; endStr= "[/list]";
        }

        if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
            s= strCut(defaultStr, startStr, endStr, 2);

            //s = defaultStr
            s= replace(s, "[$id$]", cStr(rs["id"]));
            for( j= 0 ; j<= uBound(splFieldName); j++){
                if( splFieldName[j] != "" ){
                    splxx= aspSplit(splFieldName[j] + "|||", "|");
                    fieldName= splxx[0];
                    replaceStr= cStr(rs[fieldName]) + "";
                    s= replaceValueParam(s, fieldName, replaceStr);
                }

                if( isMakeHtml== true ){
                    url= getHandleRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "/detail/detail" + cStr(rs["id"]));
                }else{
                    url= handleWebUrl("?act=detail&id=" + cStr(rs["id"]));
                    if( cStr(rs["customaurl"]) != "" ){
                        url= cStr(rs["customaurl"]);
                    }
                }

                //A���������ɫ
                abcolorStr= "";
                if( inStr(fieldNameList, ",titlecolor,") > 0 ){
                    //A������ɫ
                    if( cStr(rs["titlecolor"]) != "" ){
                        abcolorStr= "color:" + cStr(rs["titlecolor"]) + ";";
                    }
                }
                if( inStr(fieldNameList, ",flags,") > 0 ){
                    //A���ӼӴ�
                    if( inStr(cStr(rs["flags"]), "|b|") > 0 ){
                        abcolorStr= abcolorStr + "font-weight:bold;";
                    }
                }
                if( abcolorStr != "" ){
                    abcolorStr= " style=\"" + abcolorStr + "\"";
                }

                //�򿪷�ʽ2016
                if( inStr(fieldNameList, ",target,") > 0 ){
                    atargetStr= IIF(cStr(rs["target"]) != "", " target=\"" + cStr(rs["target"]) + "\"", "");
                }

                //A��title
                if( inStr(fieldNameList, ",title,") > 0 ){
                    atitleStr= IIF(cStr(rs["title"]) != "", " title=\"" + cStr(rs["title"]) + "\"", "");
                }

                //A��nofollow
                if( inStr(fieldNameList, ",nofollow,") > 0 ){
                    anofollowStr= IIF(cInt(rs["nofollow"]) != 0, " rel=\"nofollow\"", "");
                }



                s= replaceValueParam(s, "url", url);
                s= replaceValueParam(s, "abcolor", abcolorStr); //A���Ӽ���ɫ��Ӵ�
                s= replaceValueParam(s, "atitle", atitleStr); //A����title
                s= replaceValueParam(s, "anofollow", anofollowStr); //A����nofollow
                s= replaceValueParam(s, "atarget", atargetStr); //A���Ӵ򿪷�ʽ


            }
        }
        //call echo("tableName",tableName)
        idPage= getThisIdPage(db_PREFIX + tableName, cStr(rs["id"]), 10);
        //�����ԡ�
        if( tableName== "guestbook" ){
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=GuestBook&lableTitle=����&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page=" + idPage + "&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);

            //��Ĭ����ʾ���¡�
        }else{
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=" + idPage + "&parentid=" + cStr(rs["parentid"]) + "&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);
        }
        s= handleDisplayOnlineEditDialog(url, s, "", "div|li|span");

        c= c + s;
    }
    content= replace(content, "<!--#body start#-->" + defaultStr + "<!--#body end#-->", c);

    if( isMakeHtml== true ){
        url= "";
        if( len(listFileName) > 5 ){
            url= mid(listFileName, 1, len(listFileName) - 5) + "[id].html";
            url= urlAddHttpUrl(cfg_webSiteUrl, url);
        }
    }else{
        url= getUrlAddToParam(getUrl(), "?page=[id]", "replace");
    }

    return content;
}


//****************************************************
//Ĭ���б�ģ��
string defaultListTemplate(string sType, string sName){
    string c=""; string templateHtml=""; string listTemplate=""; string startStr=""; string endStr=""; string lableName="";

    templateHtml= getFText(cfg_webTemplate + "/" + templateName);
    //����Ŀ��������������Ŀ���ͣ���Ĭ��20160630
    lableName= sName + "list";
    startStr= "<!--#" + lableName + " start#-->";
    endStr= "<!--#" + lableName + " end#-->";
    if( inStr(templateHtml, startStr)== 0 || inStr(templateHtml, endStr)== 0 ){
        lableName= sType + "list";
        startStr= "<!--#" + lableName + " start#-->";
        endStr= "<!--#" + lableName + " end#-->";
    }
    if( inStr(templateHtml, startStr)== 0 || inStr(templateHtml, endStr)== 0 ){
        lableName= "list";
        startStr= "<!--#" + lableName + " start#-->";
        endStr= "<!--#" + lableName + " end#-->";
    }

    //call rwend(templateHtml)
    if( inStr(templateHtml, startStr) > 0 && inStr(templateHtml, endStr) > 0 ){
        listTemplate= strCut(templateHtml, startStr, endStr, 2);
    }else{
        startStr= "<!--#" + lableName;
        endStr= "#-->";
        if( inStr(templateHtml, startStr) > 0 && inStr(templateHtml, endStr) > 0 ){
            listTemplate= strCut(templateHtml, startStr, endStr, 2);
        }
    }
    if( listTemplate== "" ){
        c= "<ul class=\"list\"><!--#body start#-->" + vbCrlf();
        c= c + "[list]    <li><a href=\"[$url$]\"[$atitle$][$atarget$][$abcolor$][$anofollow$]>[$title$]</a><span class=\"time\">[$adddatetime format_time='1'$]</span></li>" + vbCrlf();
        c= c + "[/list]11111111111<!--#body end#--> " + vbCrlf();
        c= c + "</ul>" + vbCrlf();
        c= c + "<div class=\"clear10\"></div>" + vbCrlf();
        c= c + "<div>[$pageInfo$]</div>" + vbCrlf();
        listTemplate= c;
    }
    //call rwend(listTemplate)

    return listTemplate;
}


//���ؾ�����
void loadRun(){
    //����Ϊ�˸�.netʹ�õģ���Ϊ��.net����ȫ�ֱ��������б���
    WEB_CACHEFile=replace(replace(WEB_CACHEFile,"[adminDir]",adminDir),"[EDITORTYPE]",EDITORTYPE);

    //���崦��20160622
    cacheHtmlFilePath= "/cache/html/" + setFileName(getThisUrlFileParam()) + ".html";
    //���û���
    if( cStr(Request["cache"]) != "false" && isOnCacheHtml== true ){
        if( checkFile(cacheHtmlFilePath)== true ){
            //call echo("��ȡ�����ļ�","OK")
            rwEnd(getFText(cacheHtmlFilePath));
        }
    }

    //��¼��ǰ׺
    if( cStr(Request["db_PREFIX"]) != "" ){
        db_PREFIX= cStr(Request["db_PREFIX"]);
    }else if( cStr(Session["db_PREFIX"]) != "" ){
        db_PREFIX= cStr(Session["db_PREFIX"]);
    }
    //������ַ����
    loadWebConfig();
    isMakeHtml= false; //Ĭ������HTMLΪ�ر�
    if( cStr(Request["isMakeHtml"])== "1" || cStr(Request["isMakeHtml"])== "true" ){
        isMakeHtml= true;
    }
    templateName= cStr(Request["templateName"]); //ģ������

    //�������ݴ���ҳ
    switch ( cStr(Request["act"]) ){
        case "savedata" : saveData(cStr(Request["stype"])) ; Response.End(); //��������
        break;//'վ��ͳ�� | ����IP[653] | ����PV[9865] | ��ǰ����[65]')
        case "webstat" : webStat(adminDir + "/Data/Stat/") ; Response.End(); //��վͳ��
        break;
        case "saveSiteMap" : isMakeHtml= true ; saveSiteMap() ; Response.End(); //����sitemap.xml
        break;
        case "handleAction":
        if( cStr(Request["ishtml"])== "1" ){
            isMakeHtml= true;
        }
        rwEnd(handleAction(cStr(Request["content"]))); //������
        break;
    }


    //����html
    if( cStr(Request["act"])== "makehtml" ){
        echo("makehtml", "makehtml");
        isMakeHtml= true;
        makeWebHtml(" action actionType='" + cStr(Request["act"]) + "' columnName='" + cStr(Request["columnName"]) + "' id='" + cStr(Request["id"]) + "' ");
        createFileGBK("index.html", code);

        //����Html����վ
    }else if( cStr(Request["act"])== "copyHtmlToWeb" ){
        copyHtmlToWeb();
        //ȫ������
    }else if( cStr(Request["act"])== "makeallhtml" ){
        makeAllHtml("", "", cStr(Request["id"]));

        //���ɵ�ǰҳ��
    }else if( cStr(Request["isMakeHtml"]) != "" && cStr(Request["isSave"]) != "" ){

        handlePower("���ɵ�ǰHTMLҳ��"); //����Ȩ�޴���
        writeSystemLog("", "���ɵ�ǰHTMLҳ��"); //ϵͳ��־

        isMakeHtml= true;


        checkIDSQL(cStr(Request["id"]));
        rw(makeWebHtml(" action actionType='" + cStr(Request["act"]) + "' columnName='" + cStr(Request["columnName"]) + "' columnType='" + cStr(Request["columnType"]) + "' id='" + cStr(Request["id"]) + "' npage='" + cStr(Request["page"]) + "' "));
        glb_filePath= replace(glb_url, cfg_webSiteUrl, "");
        if( right(glb_filePath, 1)== "/" ){
            glb_filePath= glb_filePath + "index.html";
        }else if( glb_filePath== "" && glb_columnType== "��ҳ" ){
            glb_filePath= "index.html";
        }
        //�ļ���Ϊ��  ���ҿ�������html
        if( glb_filePath != "" && glb_isonhtml== true ){
            createDirFolder(getFileAttr(glb_filePath, "1"));
            createFileGBK(glb_filePath, code);
            if( cStr(Request["act"])== "detail" ){
                connexecute("update " + db_PREFIX + "ArticleDetail set ishtml=true where id=" + cStr(Request["id"]));
            }else if( cStr(Request["act"])== "nav" ){
                if( cStr(Request["id"]) != "" ){
                    connexecute("update " + db_PREFIX + "WebColumn set ishtml=true where id=" + cStr(Request["id"]));
                }else{
                    connexecute("update " + db_PREFIX + "WebColumn set ishtml=true where columnname='" + cStr(Request["columnName"]) + "'");
                }
            }
            echo("�����ļ�·��", "<a href=\"" + glb_filePath + "\" target='_blank'>" + glb_filePath + "</a>");

            //�������������� 20160216
            if( glb_columnType== "����" ){
                makeAllHtml("", "", glb_columnId);
            }

        }

        //ȫ������
    }else if( cStr(Request["act"])== "Search" ){
        rw(makeWebHtml("actionType='Search' npage='"+ IIF(cStr(Request["page"])=="","1",cStr(Request["page"])) +"' "));
    }else{
        if( lCase(cStr(Request["issave"]))== "1" ){
            makeAllHtml(cStr(Request["columnType"]), cStr(Request["columnName"]), cStr(Request["columnId"]));
        }else{
            checkIDSQL(cStr(Request["id"]));
            rw(makeWebHtml(" action actionType='" + cStr(Request["act"]) + "' columnName='" + cStr(Request["columnName"]) + "' columnType='" + cStr(Request["columnType"]) + "' id='" + cStr(Request["id"]) + "' npage='" + cStr(Request["page"]) + "' "));
        }
    }
    //��������html
    if( isOnCacheHtml== true ){
        createFile(cacheHtmlFilePath, code); //���浽�����ļ���20160622
    }
}
//���ID�Ƿ�SQL��ȫ
bool checkIDSQL(string id){
    if( checkNumber(id)== false && id != "" ){
        eerr("��ʾ", "id���зǷ��ַ�");
    }
    return false;
}
//http://127.0.0.1/aspweb.asp?act=nav&columnName=ASP
//http://127.0.0.1/aspweb.asp?act=detail&id=75
//����html��̬ҳ
string makeWebHtml(string action){
    string actionType=""; int npagesize=0; int npage=0; string s=""; string url=""; string addSql=""; string sortSql=""; string sortFieldName=""; string ascOrDesc="";
    string serchKeyWordName=""; string parentid=""; //׷����20160716 home
    actionType= rParam(action, "actionType");
    s= rParam(action, "npage");
    s= getNumber(s);
    if( s== "" ){
        npage= 1;
    }else{
        npage= cInt(s);
    }
    //����
    if( actionType== "nav" ){
        glb_columnType= rParam(action, "columnType");
        glb_columnName= rParam(action, "columnName");
        glb_columnId= rParam(action, "columnId");
        if( glb_columnId== "" ){
            glb_columnId= rParam(action, "id");
        }
        if( glb_columnType != "" ){
            addSql= "where columnType='" + glb_columnType + "'";
        }
        if( glb_columnName != "" ){
            addSql= getWhereAnd(addSql, "where columnName='" + glb_columnName + "'");
        }
        if( glb_columnId != "" ){
            addSql= getWhereAnd(addSql, "where id=" + glb_columnId + "");
        }
        //call echo("addsql",addsql)
        rs = new OleDbCommand("Select * from " + db_PREFIX + "webcolumn " + addSql, conn).ExecuteReader();

        if( rs.Read() ){
            glb_columnId= cStr(rs["id"]);
            glb_columnName= cStr(rs["columnname"]);
            glb_columnType= cStr(rs["columntype"]);
            glb_bodyContent= cStr(rs["bodycontent"]);
            glb_detailTitle= glb_columnName;
            glb_flags= cStr(rs["flags"]);
            npagesize= cInt(rs["npagesize"]); //ÿҳ��ʾ����
            glb_isonhtml= IIF(cBool(rs["isonhtml"])== true, true, false); //�Ƿ����ɾ�̬��ҳ
            sortSql= " " + cStr(rs["sortsql"]); //����SQL

            if( cStr(rs["webtitle"]) != "" ){
                cfg_webTitle= cStr(rs["webtitle"]); //��ַ����
            }
            if( cStr(rs["webkeywords"]) != "" ){
                cfg_webKeywords= cStr(rs["webkeywords"]); //��վ�ؼ���
            }
            if( cStr(rs["webdescription"]) != "" ){
                cfg_webDescription= cStr(rs["webdescription"]); //��վ����
            }
            if( templateName== "" ){
                if( aspTrim(cStr(rs["templatepath"])) != "" ){
                    templateName= cStr(rs["templatepath"]);
                }else if( cStr(rs["columntype"]) != "��ҳ" ){
                    templateName= getDateilTemplate(cStr(rs["id"]), "List");
                }
            }
        }
        glb_columnENType= handleColumnType(glb_columnType);
        glb_url= getColumnUrl(glb_columnName, "name");

        //�������б�
        if( inStr("|��Ʒ|����|��Ƶ|����|����|", "|" + glb_columnType + "|") > 0 ){
            glb_bodyContent= getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "��Ŀ�б�", "*", npagesize, cStr(npage), "where parentid=" + glb_columnId + sortSql);
            //�������б�
        }else if( inStr("|����|", "|" + glb_columnType + "|") > 0 ){
            glb_bodyContent= getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "GuestBook", "�����б�", "*", npagesize, cStr(npage), " where isthrough<>0 " + sortSql);
        }else if( glb_columnType== "�ı�" ){
            //������Ŀ�ӹ���
            if( cStr(Request["gl"])== "edit" ){
                glb_bodyContent= "<span>" + glb_bodyContent + "</span>";
            }
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" + glb_columnId + "&n=" + getRnd(11);
            glb_bodyContent= handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span");

        }
        //ϸ��
    }else if( actionType== "detail" ){
        glb_locationType= "detail";
        rs = new OleDbCommand("Select * from " + db_PREFIX + "articledetail where id=" + rParam(action, "id"), conn).ExecuteReader();

        if( rs.Read() ){
            glb_columnName= getColumnName(cStr(rs["parentid"]));
            glb_detailTitle= cStr(rs["title"]);
            glb_flags= cStr(rs["flags"]);
            glb_isonhtml= cBool(rs["isonhtml"]); //�Ƿ����ɾ�̬��ҳ
            glb_id= cStr(rs["id"]); //����ID
            if( isMakeHtml== true ){
                glb_url= getHandleRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "/detail/detail" + cStr(rs["id"]));
            }else{
                glb_url= handleWebUrl("?act=detail&id=" + cStr(rs["id"]));
            }

            if( cStr(rs["webtitle"]) != "" ){
                cfg_webTitle= cStr(rs["webtitle"]); //��ַ����
            }
            if( cStr(rs["webkeywords"]) != "" ){
                cfg_webKeywords= cStr(rs["webkeywords"]); //��վ�ؼ���
            }
            if( cStr(rs["webdescription"]) != "" ){
                cfg_webDescription= cStr(rs["webdescription"]); //��վ����
            }

            //�Ľ�20160628
            sortFieldName= "id";
            ascOrDesc= "asc";
            addSql= aspTrim(getWebColumnSortSql(cStr(rs["parentid"])));
            if( addSql != "" ){
                sortFieldName= aspTrim(replace(replace(replace(addSql, "order by", ""), " desc", ""), " asc", ""));
                if( inStr(addSql, " desc") > 0 ){
                    ascOrDesc= "desc";
                }
            }
            glb_articleAuthor= cStr(rs["author"]);
            glb_articleAdddatetime= cStr(rs["adddatetime"]);
            glb_upArticle= upArticle(cStr(rs["parentid"]), sortFieldName, cStr(rs[sortFieldName]), ascOrDesc);
            glb_downArticle= downArticle(cStr(rs["parentid"]), sortFieldName, cStr(rs[sortFieldName]), ascOrDesc);
            glb_aritcleRelatedTags= aritcleRelatedTags(cStr(rs["relatedtags"]));
            glb_aritcleSmallImage= cStr(rs["smallimage"]);
            glb_aritcleBigImage= cStr(rs["bigimage"]);
            glb_articlehits= cStr(rs["hits"]);
            connexecute("update " + db_PREFIX + "articledetail set hits=hits+1 where id=" + cStr(rs["id"])); //���µ����
            //��������
            //glb_bodyContent = "<div class=""articleinfowrap"">[$articleinfowrap$]</div>" & rs("bodycontent") & "[$relatedtags$]<ul class=""updownarticlewrap"">[$updownArticle$]</ul>"
            //��һƪ���£���һƪ����
            //glb_bodyContent = Replace(glb_bodyContent, "[$updownArticle$]", upArticle(rs("parentid"), "sortrank", rs("sortrank")) & downArticle(rs("parentid"), "sortrank", rs("sortrank")))
            //glb_bodyContent = Replace(glb_bodyContent, "[$articleinfowrap$]", "��Դ��" & rs("author") & " &nbsp; ����ʱ�䣺" & format_Time(rs("adddatetime"), 1))
            //glb_bodyContent = Replace(glb_bodyContent, "[$relatedtags$]", aritcleRelatedTags(rs("relatedtags")))

            glb_bodyContent= cStr(rs["bodycontent"]);

            //������ϸ�ӿ���
            if( cStr(Request["gl"])== "edit" ){
                glb_bodyContent= "<span>" + glb_bodyContent + "</span>";
            }
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=" + rParam(action, "id") + "&n=" + getRnd(11);
            glb_bodyContent= handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span");

            if( templateName== "" ){
                if( aspTrim(cStr(rs["templatepath"])) != "" ){
                    templateName= cStr(rs["templatepath"]);
                }else{
                    templateName= getDateilTemplate(cStr(rs["parentid"]), "Detail");
                }
            }

        }

        //��ҳ
    }else if( actionType== "onepage" ){
        rs = new OleDbCommand("Select * from " + db_PREFIX + "onepage where id=" + rParam(action, "id"), conn).ExecuteReader();

        if( rs.Read() ){
            glb_detailTitle= cStr(rs["title"]);
            glb_isonhtml= cBool(rs["isonhtml"]); //�Ƿ����ɾ�̬��ҳ
            if( isMakeHtml== true ){
                glb_url= getHandleRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "/page/page" + cStr(rs["id"]));
            }else{
                glb_url= handleWebUrl("?act=detail&id=" + cStr(rs["id"]));
            }

            if( cStr(rs["webtitle"]) != "" ){
                cfg_webTitle= cStr(rs["webtitle"]); //��ַ����
            }
            if( cStr(rs["webkeywords"]) != "" ){
                cfg_webKeywords= cStr(rs["webkeywords"]); //��վ�ؼ���
            }
            if( cStr(rs["webdescription"]) != "" ){
                cfg_webDescription= cStr(rs["webdescription"]); //��վ����
            }
            //����
            glb_bodyContent= cStr(rs["bodycontent"]);


            //������ϸ�ӿ���
            if( cStr(Request["gl"])== "edit" ){
                glb_bodyContent= "<span>" + glb_bodyContent + "</span>";
            }
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=" + rParam(action, "id") + "&n=" + getRnd(11);
            glb_bodyContent= handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span");


            if( templateName== "" ){
                if( aspTrim(cStr(rs["templatepath"])) != "" ){
                    templateName= cStr(rs["templatepath"]);
                }else{
                    templateName= "Main_Model.html";
                    //call echo(templateName,"templateName")
                }
            }

        }

        //����
    }else if( actionType== "Search" ){
        templateName= "Main_Model.html";
        serchKeyWordName= cStr(Request["keywordname"]);
        parentid= cStr(Request["parentid"]);
        if( serchKeyWordName== "" ){
            serchKeyWordName= "wd";
        }
        glb_searchKeyWord= replace(cStr(Request[serchKeyWordName]), "<", "&lt;");
        addSql= "";
        if( parentid != "" ){
            addSql= " where parentid=" + parentid;
        }
        addSql= getWhereAnd(addSql, " where title like '%" + glb_searchKeyWord + "%'");
        npagesize= 20;
        glb_bodyContent= getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "��վ��Ŀ", "*", npagesize, cStr(npage), addSql);
        positionEndStr= " >> �������ݡ�" + glb_searchKeyWord + "��";
        //���صȴ�
    }else if( actionType== "loading" ){
        rwEnd("ҳ�����ڼ����С�����");
    }
    //ģ��Ϊ�գ�����Ĭ����ҳģ��
    if( templateName== "" ){
        templateName= "Index_Model.html"; //Ĭ��ģ��
    }
    //��⵱ǰ·���Ƿ���ģ��
    if( inStr(templateName, "/")== 0 ){
        templateName= cfg_webTemplate + "/" + templateName;
    }
    //call echo("templateName",templateName)
    if( checkFile(templateName)== false ){
        eerr("δ�ҵ�ģ���ļ�", templateName);
    }
    code= getFText(templateName);

    code= handleAction(code); //������
    code= thisPosition(code); //λ��
    code= replaceGlobleVariable(code); //�滻ȫ�ֱ�ǩ
    code= handleAction(code); //������    '����һ�Σ��������������ﶯ��

    //call die(code)
    code= handleAction(code); //������
    code= handleAction(code); //������
    code= thisPosition(code); //λ��
    code= replaceGlobleVariable(code); //�滻ȫ�ֱ�ǩ
    code= delTemplateMyNote(code); //ɾ����������
    code= handleAction(code); //������

    //��ʽ��HTML
    if( inStr(cfg_flags, "|formattinghtml|") > 0 ){
        //code = HtmlFormatting(code)        '��
        code= handleHtmlFormatting(code, false, 0, "ɾ������"); //�Զ���
        //��ʽ��HTML�ڶ���
    }else if( inStr(cfg_flags, "|formattinghtmltow|") > 0 ){
        code= htmlFormatting(code); //��
        code= handleHtmlFormatting(code, false, 0, "ɾ������"); //�Զ���
        //ѹ��HTML
    }else if( inStr(cfg_flags, "|ziphtml|") > 0 ){
        code= zipHTML(code);

    }
    //�պϱ�ǩ
    if( inStr(cfg_flags, "|labelclose|") > 0 ){
        code= handleCloseHtml(code, true, ""); //ͼƬ�Զ���alt  "|*|",
    }

    //���߱༭20160127
    if( rq("gl")== "edit" ){
        if( inStr(code, "</head>") > 0 ){
            if( inStr(lCase(code), "jquery.min.js")== 0 ){
                code= replace(code, "</head>", "<script src=\"/Jquery/jquery.Min.js\"></"+"script></head>");
            }
            code= replace(code, "</head>", "<script src=\"/Jquery/Callcontext_menu.js\"></"+"script></head>");
        }
        if( inStr(code, "<body>") > 0 ){
            //Code = Replace(Code,"<body>", "<body onLoad=""ContextMenu.intializeContextMenu()"">")
        }
    }
    //call echo(templateName,templateName)
    return code;
}

//���Ĭ��ϸ��ģ��ҳ
string getDateilTemplate(string parentid, string templateType){
    string templateName="";
    templateName= "Main_Model.html";
    rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn where id=" + parentid, conn).ExecuteReader();

    if( rsx.Read() ){
        //call echo("columntype",rsx("columntype"))
        if( cStr(rsx["columntype"])== "����" ){
            //����ϸ��ҳ
            if( checkFile(cfg_webTemplate + "/News_" + templateType + ".html")== true ){
                templateName= "News_" + templateType + ".html";
            }
        }else if( cStr(rsx["columntype"])== "��Ʒ" ){
            //��Ʒϸ��ҳ
            if( checkFile(cfg_webTemplate + "/Product_" + templateType + ".html")== true ){
                templateName= "Product_" + templateType + ".html";
            }
        }else if( cStr(rsx["columntype"])== "����" ){
            //����ϸ��ҳ
            if( checkFile(cfg_webTemplate + "/Down_" + templateType + ".html")== true ){
                templateName= "Down_" + templateType + ".html";
            }

        }else if( cStr(rsx["columntype"])== "��Ƶ" ){
            //��Ƶϸ��ҳ
            if( checkFile(cfg_webTemplate + "/Video_" + templateType + ".html")== true ){
                templateName= "Video_" + templateType + ".html";
            }
        }else if( cStr(rsx["columntype"])== "����" ){
            //��Ƶϸ��ҳ
            if( checkFile(cfg_webTemplate + "/GuestBook_" + templateType + ".html")== true ){
                templateName= "Video_" + templateType + ".html";
            }
        }else if( cStr(rsx["columntype"])== "�ı�" ){
            //��Ƶϸ��ҳ
            if( checkFile(cfg_webTemplate + "/Page_" + templateType + ".html")== true ){
                templateName= "Page_" + templateType + ".html";
            }
        }
    }
    //call echo(templateType,templateName)
    string getDateilTemplate= templateName;

    return getDateilTemplate;
}


//����ȫ��htmlҳ��
void makeAllHtml(string columnType, string columnName, string columnId){
    string action=""; string s=""; int i=0; int nPageSize=0; int nCountSize=0; int nPage=0; string addSql=""; string url=""; string articleSql="";
    handlePower("����ȫ��HTMLҳ��"); //����Ȩ�޴���
    writeSystemLog("", "����ȫ��HTMLҳ��"); //ϵͳ��־

    isMakeHtml= true;
    //��Ŀ
    echo("��Ŀ", "");
    if( columnType != "" ){
        addSql= "where columnType='" + columnType + "'";
    }
    if( columnName != "" ){
        addSql= getWhereAnd(addSql, "where columnName='" + columnName + "'");
    }
    if( columnId != "" ){
        addSql= getWhereAnd(addSql, "where id in(" + columnId + ")");
    }
    rss = new OleDbCommand("select * from " + db_PREFIX + "webcolumn " + addSql + " order by sortrank asc", conn).ExecuteReader();

    while( rss.Read()){
        glb_columnName= "";
        //��������html
        if( cBool(rss["isonhtml"])== true ){
            if( inStr("|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����|", "|" + cStr(rss["columntype"]) + "|") > 0 ){
                if( cStr(rss["columntype"])== "����" ){
                    nCountSize= getRecordCount(db_PREFIX + "guestbook", ""); //��¼��
                }else{
                    nCountSize= getRecordCount(db_PREFIX + "articledetail", " where parentid=" + cStr(rss["id"])); //��¼��
                }
                nPageSize= cInt(rss["npagesize"]);
                nPage= getPageNumb(cInt(nCountSize), cInt(nPageSize));
                if( nPage <= 0 ){
                    nPage= 1;
                }
                for( i= 1 ; i<= nPage; i++){
                    url= getHandleRsUrl(cStr(rss["filename"]), cStr(rss["customaurl"]), "/nav" + cStr(rss["id"]));
                    glb_filePath= replace(url, cfg_webSiteUrl, "");
                    if( right(glb_filePath, 1)== "/" || glb_filePath== "" ){
                        glb_filePath= glb_filePath + "index.html";
                    }
                    //call echo("glb_filePath",glb_filePath)
                    action= " action actionType='nav' columnName='" + cStr(rss["columnname"]) + "' npage='" + i + "' listfilename='" + glb_filePath + "' ";
                    //call echo("action",action)
                    makeWebHtml(action);
                    if( i > 1 ){
                        glb_filePath= mid(glb_filePath, 1, len(glb_filePath) - 5) + i + ".html";
                    }
                    s= "<a href=\"" + glb_filePath + "\" target='_blank'>" + glb_filePath + "</a>(" + cBool(rss["isonhtml"]) + ")";
                    echo(action, s);
                    if( glb_filePath != "" ){
                        createDirFolder(getFileAttr(glb_filePath, "1"));
                        createFileGBK(glb_filePath, code);
                    }
                    doEvents();
                    templateName= ""; //���ģ���ļ�����
                }
            }else{
                action= " action actionType='nav' columnName='" + cStr(rss["columnname"]) + "'";
                makeWebHtml(action);
                glb_filePath= replace(getColumnUrl(cStr(rss["columnname"]), "name"), cfg_webSiteUrl, "");
                if( right(glb_filePath, 1)== "/" || glb_filePath== "" ){
                    glb_filePath= glb_filePath + "index.html";
                }
                s= "<a href=\"" + glb_filePath + "\" target='_blank'>" + glb_filePath + "</a>(" + cBool(rss["isonhtml"]) + ")";
                echo(action, s);
                if( glb_filePath != "" ){
                    createDirFolder(getFileAttr(glb_filePath, "1"));
                    createFileGBK(glb_filePath, code);
                }
                doEvents();
                templateName= "";
            }
            connexecute("update " + db_PREFIX + "WebColumn set ishtml=true where id=" + cStr(rss["id"])); //���µ���Ϊ����״̬
        }
    }

    //��������ָ����Ŀ��Ӧ����
    if( columnId != "" ){
        articleSql= "select * from " + db_PREFIX + "articledetail where parentid=" + columnId + " order by sortrank asc";
        //������������
    }else if( addSql== "" ){
        articleSql= "select * from " + db_PREFIX + "articledetail order by sortrank asc";
    }
    if( articleSql != "" ){
        //����
        echo("����", "");
        rss = new OleDbCommand(articleSql, conn).ExecuteReader();

        while( rss.Read()){
            glb_columnName= "";
            action= " action actionType='detail' columnName='" + cStr(rss["parentid"]) + "' id='" + cStr(rss["id"]) + "'";
            //call echo("action",action)
            makeWebHtml(action);
            glb_filePath= replace(glb_url, cfg_webSiteUrl, "");
            if( right(glb_filePath, 1)== "/" ){
                glb_filePath= glb_filePath + "index.html";
            }
            s= "<a href=\"" + glb_filePath + "\" target='_blank'>" + glb_filePath + "</a>(" + cBool(rss["isonhtml"]) + ")";
            echo(action, s);
            //�ļ���Ϊ��  ���ҿ�������html
            if( glb_filePath != "" && cBool(rss["isonhtml"])== true ){
                createDirFolder(getFileAttr(glb_filePath, "1"));
                createFileGBK(glb_filePath, code);
                connexecute("update " + db_PREFIX + "ArticleDetail set ishtml=true where id=" + cStr(rss["id"])); //��������Ϊ����״̬
            }
            templateName= ""; //���ģ���ļ�����
        }
    }

    if( addSql== "" ){
        //��ҳ
        echo("��ҳ", "");
        rss = new OleDbCommand("select * from " + db_PREFIX + "onepage order by sortrank asc", conn).ExecuteReader();

        while( rss.Read()){
            glb_columnName= "";
            action= " action actionType='onepage' id='" + cStr(rss["id"]) + "'";
            //call echo("action",action)
            makeWebHtml(action);
            glb_filePath= replace(glb_url, cfg_webSiteUrl, "");
            if( right(glb_filePath, 1)== "/" ){
                glb_filePath= glb_filePath + "index.html";
            }
            s= "<a href=\"" + glb_filePath + "\" target='_blank'>" + glb_filePath + "</a>(" + cBool(rss["isonhtml"]) + ")";
            echo(action, s);
            //�ļ���Ϊ��  ���ҿ�������html
            if( glb_filePath != "" && cBool(rss["isonhtml"])== true ){
                createDirFolder(getFileAttr(glb_filePath, "1"));
                createFileGBK(glb_filePath, code);
                connexecute("update " + db_PREFIX + "onepage set ishtml=true where id=" + cStr(rss["id"])); //���µ�ҳΪ����״̬
            }
            templateName= ""; //���ģ���ļ�����
        }

    }
}

//����html����վ
void copyHtmlToWeb(){
    string webDir=""; string toWebDir=""; string toFilePath=""; string filePath=""; string fileName=""; string fileList=""; string[] splStr; string content=""; string s=""; string s1=""; string c=""; string webImages=""; string webCss=""; string webJs=""; string[] splJs;
    string webFolderName=""; string jsFileList=""; string setFileCode=""; int nErrLevel=0; string jsFilePath=""; string url="";

    setFileCode= cStr(Request["setcode"]); //�����ļ��������

    handlePower("��������HTMLҳ��"); //����Ȩ�޴���
    writeSystemLog("", "��������HTMLҳ��"); //ϵͳ��־

    webFolderName= cfg_webTemplate;
    if( left(webFolderName, 1)== "/" ){
        webFolderName= mid(webFolderName, 2,-1);
    }
    if( right(webFolderName, 1)== "/" ){
        webFolderName= mid(webFolderName, 1, len(webFolderName) - 1);
    }
    if( inStr(webFolderName, "/") > 0 ){
        webFolderName= mid(webFolderName, inStr(webFolderName, "/") + 1,-1);
    }
    webDir= "/htmladmin/" + webFolderName + "/";
    toWebDir= "/htmlw" + "eb/viewweb/";
    createDirFolder(toWebDir);
    toWebDir= toWebDir + pinYin2(webFolderName) + "/";

    deleteFolder(toWebDir); //ɾ��
    createFolder("/htmlweb/web"); //�����ļ��� ��ֹweb�ļ��в�����20160504
    deleteFolder(webDir);
    createDirFolder(webDir);
    webImages= webDir + "Images/";
    webCss= webDir + "Css/";
    webJs= webDir + "Js/";
    copyFolder(cfg_webImages, webImages);
    copyFolder(cfg_webCss, webCss);
    createFolder(webJs); //����Js�ļ���


    //����Js�ļ���
    splJs= aspSplit(getDirJsList(webJs), vbCrlf());
    foreach(var eachfilePath in splJs){
        filePath=eachfilePath;
        if( filePath != "" ){
            toFilePath= webJs + getFileName(filePath);
            echo("js", filePath);
            moveFile(filePath, toFilePath);
        }
    }
    //����Css�ļ���
    splStr= aspSplit(getDirCssList(webCss), vbCrlf());
    foreach(var eachfilePath in splStr){
        filePath=eachfilePath;
        if( filePath != "" ){
            content= getFText(filePath);
            content= replace(content, cfg_webImages, "../images/");

            content= deleteCssNote(content);
            content= phpTrim(content);
            //����Ϊutf-8���� 20160527
            if( lCase(setFileCode)== "utf-8" ){
                content= replace(content, "gb2312", "utf-8");
            }
            writeToFile(filePath, content, setFileCode);
            echo("css", cfg_webImages);
        }
    }
    //������ĿHTML
    isMakeHtml= true;
    rss = new OleDbCommand("select * from " + db_PREFIX + "webcolumn where isonhtml=true", conn).ExecuteReader();

    while( rss.Read()){
        glb_filePath= replace(getColumnUrl(cStr(rss["columnname"]), "name"), cfg_webSiteUrl, "");

        if( right(glb_filePath, 1)== "/" || right(glb_filePath, 1)== "" ){
            glb_filePath= glb_filePath + "index.html";
        }
        if( right(glb_filePath, 5)== ".html" ){
            if( right(glb_filePath, 11)== "/index.html" ){
                fileList= fileList + glb_filePath + vbCrlf();
            }else{
                fileList= glb_filePath + vbCrlf() + fileList;
            }
            fileName= replace(glb_filePath, "/", "_");
            toFilePath= webDir + fileName;
            copyFile(glb_filePath, toFilePath);
            echo("����", glb_filePath);
        }
    }
    //��������HTML
    rss = new OleDbCommand("select * from " + db_PREFIX + "articledetail where isonhtml=true", conn).ExecuteReader();

    while( rss.Read()){
        glb_url= getHandleRsUrl(cStr(rss["filename"]), cStr(rss["customaurl"]), "/detail/detail" + cStr(rss["id"]));
        glb_filePath= replace(glb_url, cfg_webSiteUrl, "");
        if( right(glb_filePath, 1)== "/" || right(glb_filePath, 1)== "" ){
            glb_filePath= glb_filePath + "index.html";
        }
        if( right(glb_filePath, 5)== ".html" ){
            if( right(glb_filePath, 11)== "/index.html" ){
                fileList= fileList + glb_filePath + vbCrlf();
            }else{
                fileList= glb_filePath + vbCrlf() + fileList;
            }
            fileName= replace(glb_filePath, "/", "_");
            toFilePath= webDir + fileName;
            copyFile(glb_filePath, toFilePath);
            echo("����" + cStr(rss["title"]), glb_filePath);
        }
    }
    //���Ƶ���HTML
    rss = new OleDbCommand("select * from " + db_PREFIX + "onepage where isonhtml=true", conn).ExecuteReader();

    while( rss.Read()){
        glb_url= getHandleRsUrl(cStr(rss["filename"]), cStr(rss["customaurl"]), "/page/page" + cStr(rss["id"]));
        glb_filePath= replace(glb_url, cfg_webSiteUrl, "");
        if( right(glb_filePath, 1)== "/" || right(glb_filePath, 1)== "" ){
            glb_filePath= glb_filePath + "index.html";
        }
        if( right(glb_filePath, 5)== ".html" ){
            if( right(glb_filePath, 11)== "/index.html" ){
                fileList= fileList + glb_filePath + vbCrlf();
            }else{
                fileList= glb_filePath + vbCrlf() + fileList;
            }
            fileName= replace(glb_filePath, "/", "_");
            toFilePath= webDir + fileName;
            copyFile(glb_filePath, toFilePath);
            echo("��ҳ" + cStr(rss["title"]), glb_filePath);
        }
    }
    //��������html�ļ��б�
    //call echo(cfg_webSiteUrl,cfg_webTemplate)
    //call rwend(fileList)
    string sourceUrl=""; string replaceUrl="";
    splStr= aspSplit(fileList, vbCrlf());
    foreach(var eachfilePath in splStr){
        filePath=eachfilePath;
        if( filePath != "" ){
            filePath= webDir + replace(filePath, "/", "_");
            echo("filePath", filePath);
            content= getFText(filePath);
            foreach(var eachs in splStr){
                s=eachs;
                s1= s;
                if( right(s1, 11)== "/index.html" ){
                    s1= left(s1, len(s1) - 11) + "/";
                }
                sourceUrl= cfg_webSiteUrl + s1;
                replaceUrl= cfg_webSiteUrl + replace(s, "/", "_");
                //Call echo(sourceUrl, replaceUrl)                             '����  ���������ʾ20160613
                content= replace(content, sourceUrl, replaceUrl);
            }
            content= replace(content, cfg_webSiteUrl, ""); //ɾ����ַ
            content= replace(content, cfg_webTemplate + "/", ""); //ɾ��ģ��·�� ��

            //content=nullLinkAddDefaultName(content)
            foreach(var eachs in splJs){
                s=eachs;
                if( s != "" ){
                    fileName= getFileName(s);
                    content= replace(content, "Images/" + fileName, "js/" + fileName);
                }
            }
            if( inStr(content, "/Jquery/Jquery.Min.js") > 0 ){
                content= replace(content, "/Jquery/Jquery.Min.js", "js/Jquery.Min.js");
                copyFile("/Jquery/Jquery.Min.js", webJs + "/Jquery.Min.js");
            }
            content= replace(content, "<a href=\"\" ", "<a href=\"index.html\" "); //����ҳ��index.html

            createFileGBK(filePath, content);
        }
    }

    //�Ѹ�����վ���µ�images/�ļ����µ�js�Ƶ�js/�ļ�����  20160315
    string htmlFileList=""; string[] splHtmlFile; string[] splJsFile; string htmlFilePath=""; string jsFileName="";
    jsFileList= getDirJsNameList(webImages);
    htmlFileList= getDirHtmlList(webDir);
    splHtmlFile= aspSplit(htmlFileList, vbCrlf());
    splJsFile= aspSplit(jsFileList, vbCrlf());
    foreach(var eachhtmlFilePath in splHtmlFile){
        htmlFilePath=eachhtmlFilePath;
        content= getFText(htmlFilePath);
        foreach(var eachjsFileName in splJsFile){
            jsFileName=eachjsFileName;
            content= regExp_Replace(content, "Images/" + jsFileName, "js/" + jsFileName);
        }

        nErrLevel= 0;
        content= handleHtmlFormatting(content, false, nErrLevel, "|ɾ������|"); //|ɾ������|
        content= handleCloseHtml(content, true, ""); //�պϱ�ǩ
        //nErrLevel = checkHtmlFormatting(content)
        if( checkHtmlFormatting(content)== false ){
            echoRed(htmlFilePath + "(��ʽ������)", nErrLevel); //ע��
        }
        //����Ϊutf-8����
        if( lCase(setFileCode)== "utf-8" ){
            content= replace(content, "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\" />", "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />");
        }
        content= phpTrim(content);
        writeToFile(htmlFilePath, content, setFileCode);
    }
    //images��js�ƶ���js��
    foreach(var eachjsFileName in splJsFile){
        jsFileName=eachjsFileName;
        jsFilePath= webImages + jsFileName;
        content= getFText(jsFilePath);
        content= phpTrim(content);
        writeToFile(webJs + jsFileName, content, setFileCode);
        deleteFile(jsFilePath);
    }

    copyFolder(webDir, toWebDir);
    //ʹhtmlWeb�ļ�����phpѹ��
    if( cStr(Request["isMakeZip"])== "1" ){
        makeHtmlWebToZip(webDir);
    }
    //ʹ��վ��xml���20160612
    if( cStr(Request["isMakeXml"])== "1" ){
        makeHtmlWebToXmlZip("/htmladmin/", webFolderName);
    }
    //�����ַ
    url= "http://10.10.10.57/" + toWebDir;
    echo("���", "<a href='" + url + "' target='_blank'>" + url + "</a>");
}
//ʹhtmlWeb�ļ�����phpѹ��
string makeHtmlWebToZip(string webDir){
    string content=""; string[] splStr; string filePath=""; string c=""; string[] arrayFile; string fileName=""; string fileType=""; bool isTrue;
    string webFolderName="";
    string cleanFileList="";
    splStr= aspSplit(webDir, "/");
    webFolderName= splStr[2];
    //content = getFileFolderList(webDir, true, "ȫ��", "", "ȫ���ļ���", "", "") 		'��������
    content=getDirAllFileList(webDir,"");
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachfilePath in splStr){
        filePath=eachfilePath;
        if( checkFolder(filePath)== false ){
            arrayFile= handleFilePathArray(filePath);
            fileName= lCase(arrayFile[2]);
            fileType= lCase(arrayFile[4]);
            fileName= remoteNumber(fileName);
            isTrue= true;

            if( inStr("|" + cleanFileList + "|", "|" + fileName + "|") > 0 && fileType== "html" ){
                isTrue= false;
            }
            if( isTrue== true ){
                //call echo(fileType,fileName)
                if( c != "" ){ c= c + "|" ;}
                c= c + replace(filePath, handlePath("/"), "");
                cleanFileList= cleanFileList + fileName + "|";
            }
        }
    }
    rw(c);
    c= c + "|||||";
    createFileGBK("htmlweb/1.txt", c);
    echo("<hr>cccccccccccc", c);
    //���ж�����ļ�����20160309
    if( checkFile("/myZIP.php")== true ){
        echo("", xmlPost(getHost() + "/myZIP.php?webFolderName=" + webFolderName, "content=" + escape(c)));
    }

    return "";
}
//ʹ��վ��xml���20160612
string makeHtmlWebToXmlZip(string sNewWebDir,string rootDir){//���պ���
    return "";
}


//���ɸ���sitemap.xml 20160118
void saveSiteMap(){
    bool isWebRunHtml; //�Ƿ�Ϊhtml��ʽ��ʾ��վ
    string changefreg=""; //����Ƶ��
    string priority=""; //���ȼ�
    string s=""; string c=""; string url="";string sql="";
    handlePower("�޸�����SiteMap"); //����Ȩ�޴���

    changefreg= cStr(Request["changefreg"]);
    priority= cStr(Request["priority"]);
    loadWebConfig(); //��������
    //call eerr("cfg_flags",cfg_flags)
    if( inStr(cfg_flags, "|htmlrun|") > 0 ){
        isWebRunHtml= true;
    }else{
        isWebRunHtml= false;
    }

    c= c + "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + vbCrlf();
    c= c + vbTab() + "<urlset xmlns=\"http://www.sitemaps.org/schemas/sitemap/0.9\">" + vbCrlf();

    //��Ŀ
    rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn where isonhtml<>0 order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(rsx["nofollow"])== 0 ){
            c= c + copyStr(vbTab(), 2) + "<url>" + vbCrlf();
            if( isWebRunHtml== true ){
                url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/nav" + cStr(rsx["id"]));
                url= handleAction(url);
            }else{
                url= escape("?act=nav&columnName=" + cStr(rsx["columnname"]));
            }
            url= urlAddHttpUrl(cfg_webSiteUrl, url);
            //call echo(cfg_webSiteUrl,url)

            c= c + copyStr(vbTab(), 3) + "<loc>" + url + "</loc>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<lastmod>" + format_Time(cStr(rsx["updatetime"]), 2) + "</lastmod>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<changefreq>" + changefreg + "</changefreq>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<priority>" + priority + "</priority>" + vbCrlf();
            c= c + copyStr(vbTab(), 2) + "</url>" + vbCrlf();
            echo("��Ŀ", "<a href=\"" + url + "\" target='_blank'>" + url + "</a>");
        }
    }

    //����
    sql="select * from " + db_PREFIX + "articledetail  where isonhtml<>0 order by sortrank asc";
    rsx = new OleDbCommand(sql, conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(rsx["nofollow"])== 0 ){
            c= c + copyStr(vbTab(), 2) + "<url>" + vbCrlf();
            if( isWebRunHtml== true ){
                url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/detail/detail" + cStr(rsx["id"]));
                url= handleAction(url);
            }else{
                url= "?act=detail&id=" + cStr(rsx["id"]);
            }
            url= urlAddHttpUrl(cfg_webSiteUrl, url);
            //call echo(cfg_webSiteUrl,url)

            c= c + copyStr(vbTab(), 3) + "<loc>" + url + "</loc>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<lastmod>" + format_Time(cStr(rsx["updatetime"]), 2) + "</lastmod>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<changefreq>" + changefreg + "</changefreq>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<priority>" + priority + "</priority>" + vbCrlf();
            c= c + copyStr(vbTab(), 2) + "</url>" + vbCrlf();
            echo("����", "<a href=\"" + url + "\">" + url + "</a>");
        }
    }

    //��ҳ
    rsx = new OleDbCommand("select * from " + db_PREFIX + "onepage where isonhtml<>0 order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(rsx["nofollow"])== 0 ){
            c= c + copyStr(vbTab(), 2) + "<url>" + vbCrlf();
            if( isWebRunHtml== true ){
                url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/page/detail" + cStr(rsx["id"]));
                url= handleAction(url);
            }else{
                url= "?act=onepage&id=" + cStr(rsx["id"]);
            }
            url= urlAddHttpUrl(cfg_webSiteUrl, url);
            //call echo(cfg_webSiteUrl,url)

            c= c + copyStr(vbTab(), 3) + "<loc>" + url + "</loc>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<lastmod>" + format_Time(cStr(rsx["updatetime"]), 2) + "</lastmod>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<changefreq>" + changefreg + "</changefreq>" + vbCrlf();
            c= c + copyStr(vbTab(), 3) + "<priority>" + priority + "</priority>" + vbCrlf();
            c= c + copyStr(vbTab(), 2) + "</url>" + vbCrlf();
            echo("��ҳ", "<a href=\"" + url + "\">" + url + "</a>");
        }
    }
    c= c + vbTab() + "</urlset>" + vbCrlf();
    loadWebConfig();
    createFile("sitemap.xml", c);
    echo("����sitemap.xml�ļ��ɹ�", "<a href='/sitemap.xml' target='_blank'>���Ԥ��sitemap.xml</a>");


    //�ж��Ƿ�����sitemap.html
    if( cStr(Request["issitemaphtml"])== "1" ){
        c= "";
        //�ڶ���
        //��Ŀ
        rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn order by sortrank asc", conn).ExecuteReader();

        while( rsx.Read()){
            if( cInt(rsx["nofollow"])== 0 ){
                if( isWebRunHtml== true ){
                    url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/nav" + cStr(rsx["id"]));
                    url= handleAction(url);
                }else{
                    url= escape("?act=nav&columnName=" + cStr(rsx["columnname"]));
                }
                url= urlAddHttpUrl(cfg_webSiteUrl, url);

                //�ж��Ƿ�����html
                if( cBool(rsx["isonhtml"])== true ){
                    s= "<a href=\"" + url + "\">" + cStr(rsx["columnname"]) + "</a>";
                }else{
                    s= "<span>" + cStr(rsx["columnname"]) + "</span>";
                }
                c= c + "<li style=\"width:20%;\">" + s + vbCrlf() + "<ul>" + vbCrlf();

                //����
                sql="select * from " + db_PREFIX + "articledetail where parentId=" + cStr(rsx["id"]) + " order by sortrank asc";
                rss = new OleDbCommand(sql, conn).ExecuteReader();

                while( rss.Read()){
                    if( cInt(rss["nofollow"])== 0 ){
                        if( isWebRunHtml== true ){
                            url= getRsUrl(cStr(rss["filename"]), cStr(rss["customaurl"]), "/detail/detail" + cStr(rss["id"]));
                            url= handleAction(url);
                        }else{
                            url= "?act=detail&id=" + cStr(rss["id"]);
                        }
                        url= urlAddHttpUrl(cfg_webSiteUrl, url);
                        //�ж��Ƿ�����html
                        if( cBool(rss["isonhtml"])== true ){
                            s= "<a href=\"" + url + "\">" + cStr(rss["title"]) + "</a>";
                        }else{
                            s= "<span>" + cStr(rss["title"]) + "</span>";
                        }
                        c= c + "<li style=\"width:20%;\">" + s + "</li>" + vbCrlf();
                    }
                }
                c= c + "</ul>" + vbCrlf() + "</li>" + vbCrlf();


            }
        }

        //����
        c= c + "<li style=\"width:20%;\"><a href=\"javascript:;\">�����б�</a>" + vbCrlf() + "<ul>" + vbCrlf();
        rsx = new OleDbCommand("select * from " + db_PREFIX + "onepage order by sortrank asc", conn).ExecuteReader();

        while( rsx.Read()){
            if( cInt(rsx["nofollow"])== 0 ){
                c= c + copyStr(vbTab(), 2) + "<url>" + vbCrlf();
                if( isWebRunHtml== true ){
                    url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/page/detail" + cStr(rsx["id"]));
                    url= handleAction(url);
                }else{
                    url= "?act=onepage&id=" + cStr(rsx["id"]);
                }
                //�ж��Ƿ�����html
                if( cBool(rsx["isonhtml"])== true ){
                    s= "<a href=\"" + url + "\">" + cStr(rsx["title"]) + "</a>";
                }else{
                    s= "<span>" + cStr(rsx["title"]) + "</span>";
                }

                c= c + "<li style=\"width:20%;\">" + s + "</li>" + vbCrlf(); //target=""_blank""  ȥ��
            }
        }
        c= c + "</ul>" + vbCrlf() + "</li>" + vbCrlf();

        string templateContent="";
        templateContent= getFText(adminDir + "/template_SiteMap.html");


        templateContent= replace(templateContent, "{$content$}", c);
        templateContent= replace(templateContent, "{$Web_Title$}", cfg_webTitle);


        createFile("sitemap.html", templateContent);
        echo("����sitemap.html�ļ��ɹ�", "<a href='/sitemap.html' target='_blank'>���Ԥ��sitemap.html</a>");
    }
    writeSystemLog("", "����sitemap.xml"); //ϵͳ��־
}
</script>
