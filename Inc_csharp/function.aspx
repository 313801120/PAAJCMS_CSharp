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
<script runat="server" language="c#">
//��̨�������ĳ��� ��� ɾ�� �޸� �б�

//����function�ļ�����
string callFunction(){
    switch ( cStr(Request["stype"]) ){
        case "updateWebsiteStat" : updateWebsiteStat() ;break;//������վͳ��
        case "clearWebsiteStat" : clearWebsiteStat() ;break;//�����վͳ��
        case "updateTodayWebStat" : updateTodayWebStat() ;break;//������վ����ͳ��
        case "websiteDetail" : websiteDetail() ;break;//��ϸ��վͳ��
        case "displayAccessDomain" : displayAccessDomain() ;break;//��ʾ��������
        case "delTemplate" : delTemplate(); //ɾ��ģ��
        break;
        default : eerr("function1ҳ��û�ж���", cStr(Request["stype"]));
        break;
    }
    return "";
}

//��ʾ��������
string displayAccessDomain(){
    string visitWebSite=""; string visitWebSiteList=""; string urlList=""; int nOK=0;
    handlePower("��ʾ��������");
    conn=openConn();
    conn.Open();

    nOK= 0;
    rs = new OleDbCommand("select * from " + db_PREFIX + "websitestat", conn).ExecuteReader();

    while( rs.Read()){
        visitWebSite= lCase(getWebSite(cStr(rs["visiturl"])));
        //call echo("visitWebSite",visitWebSite)
        if( inStr(vbCrlf() + visitWebSiteList + vbCrlf(), vbCrlf() + visitWebSite + vbCrlf())== 0 ){
            if( visitWebSite != lCase(getWebSite(webDoMain())) ){
                visitWebSiteList= visitWebSiteList + visitWebSite + vbCrlf();
                nOK= nOK + 1;
                urlList= urlList + nOK + "��<a href='" + cStr(rs["visiturl"]) + "' target='_blank'>" + cStr(rs["visiturl"]) + "</a><br>";
            }
        }
    }
    echo("��ʾ��������", "������� <a href='javascript:history.go(-1)'>�������</a>");
    rwEnd(visitWebSiteList + "<br><hr><br>" + urlList);
    return "";
}
//��ô������б� 20160313
string getHandleTableList(){
    string s=""; string lableStr="";
    lableStr= "���б�[" + cStr(Request["mdbpath"]) + "]";
    if( WEB_CACHEContent== "" ){
        WEB_CACHEContent= getFText(WEB_CACHEFile);
    }
    s= getConfigContentBlock(WEB_CACHEContent, "#" + lableStr + "#");
    if( s== "" ){
        s= lCase(getTableList());
        s= "|" + replace(s, vbCrlf(), "|") + "|";
        WEB_CACHEContent= setConfigFileBlock(WEB_CACHEFile, s, "#" + lableStr + "#");
        if( isCacheTip== true ){
            echo("����", lableStr);
        }
    }
    return s;
}

//��ô�����ֶ��б�   getHandleFieldList("ArticleDetail","�ֶ��б�")
string getHandleFieldList(string tableName, string sType){
    string s="";
    if( WEB_CACHEContent== "" ){
        WEB_CACHEContent= getFText(WEB_CACHEFile);
    }
    s= getConfigContentBlock(WEB_CACHEContent, "#" + tableName + sType + "#");

    if( s== "" ){
        if( sType== "�ֶ������б�" ){
            s= lCase(getFieldConfigList(tableName));
        }else{
            s= lCase(getFieldList(tableName));
        }
        WEB_CACHEContent= setConfigFileBlock(WEB_CACHEFile, s, "#" + tableName + sType + "#");
        if( isCacheTip== true ){
            echo("����", tableName + sType);
        }
    }
    return s;
}
//��ģ������ 20160310
string getTemplateContent(string templateFileName){
    loadWebConfig();
    //��ģ��
    string templateFile=""; string customTemplateFile=""; string c="";
    customTemplateFile= ROOT_PATH + "template/" + db_PREFIX + "/" + templateFileName;
    //Ϊ�ֻ���
    if( checkMobile()== true || cStr(Request["m"])== "mobile" ){
        templateFile= ROOT_PATH + "/Template/mobile/" + templateFileName;
    }
    //�ж��ֻ����ļ��Ƿ����20160330
    if( checkFile(templateFile)== false ){
        if( checkFile(customTemplateFile)== true ){
            templateFile= customTemplateFile;
        }else{
            templateFile= ROOT_PATH + templateFileName;
        }
    }
    c= getFText(templateFile);
    c= replaceLableContent(c);
    return c;
}
//�滻��ǩ����
string replaceLableContent(string content){
    string s=""; string c=""; string[] splStr; string list="";
    content= replace(content, "{$webVersion$}", webVersion); //��վ�汾
    content= replace(content, "{$Web_Title$}", cfg_webTitle); //��վ����
    content= replace(content, "{$EDITORTYPE$}", EDITORTYPE); //ASP��PHP
    content= replace(content, "{$adminDir$}", adminDir); //��̨Ŀ¼

    content= replace(content, "[$adminId$]", cStr(Session["adminId"])); //����ԱID
    content= replace(content, "{$adminusername$}", cStr(Session["adminusername"])); //�����˺�����
    content= replace(content, "{$EDITORTYPE$}", EDITORTYPE); //��������
    content= replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL); //ǰ̨
    content= replace(content, "{$webVersion$}", webVersion); //�汾
    content= replace(content, "{$WebsiteStat$}", getConfigFileBlock(WEB_CACHEFile, "#�ÿ���Ϣ#")); //����ÿ���Ϣ


    content= replace(content, "{$databaseType$}", databaseType); //����Ϊ����
    content= replace(content, "{$DB_PREFIX$}", db_PREFIX); //��ǰ׺
    content= replace(content, "{$adminflags$}", IIF(cStr(Session["adminflags"])== "|*|", "��������Ա", "��ͨ����Ա")); //����Ա����
    content= replace(content, "{$SERVER_SOFTWARE$}", cStr(Request.ServerVariables["SERVER_SOFTWARE"])); //�������汾
    content= replace(content, "{$SERVER_NAME$}", cStr(Request.ServerVariables["SERVER_NAME"])); //��������ַ
    content= replace(content, "{$LOCAL_ADDR$}", cStr(Request.ServerVariables["LOCAL_ADDR"])); //������IP
    content= replace(content, "{$SERVER_PORT$}", cStr(Request.ServerVariables["SERVER_PORT"])); //�������˿�
    content= replaceValueParam(content, "mdbpath", cStr(Request["mdbpath"]));
    content= replaceValueParam(content, "webDir", webDir);
    content= replaceValueParam(content, "EDITORTYPE", EDITORTYPE); //ASP��PHP

    //20160628
    if( inStr(content, "{$backupDatabaseSelectHtml$}") > 0 ){
        c= getDirTxtNameList(adminDir + "/Data/BackUpDateBases/");
        splStr= aspSplit(c, vbCrlf());
        foreach(var eachs in splStr){
            s=eachs;
            list= list + "<option value=\"" + s + "\">" + s + "</option>" + vbCrlf();
        }
        content= replace(content, "{$backupDatabaseSelectHtml$}", list);
    }

    //20160614
    if( EDITORTYPE== "php" ){
        content= replace(content, "{$EDITORTYPE_PHP$}", "php"); //��phpinc/��
    }
    content= replace(content, "{$EDITORTYPE_PHP$}", ""); //��phpinc/��

    return content;
}

//�����б���
string displayFlags(string flags){
    string c="";
    //ͷ��[h]
    if( inStr("|" + flags + "|", "|h|") > 0 ){
        c= c + "ͷ ";
    }
    //�Ƽ�[c]
    if( inStr("|" + flags + "|", "|c|") > 0 ){
        c= c + "�� ";
    }
    //�õ�[f]
    if( inStr("|" + flags + "|", "|f|") > 0 ){
        c= c + "�� ";
    }
    //�ؼ�[a]
    if( inStr("|" + flags + "|", "|a|") > 0 ){
        c= c + "�� ";
    }
    //����[s]
    if( inStr("|" + flags + "|", "|s|") > 0 ){
        c= c + "�� ";
    }
    //�Ӵ�[b]
    if( inStr("|" + flags + "|", "|b|") > 0 ){
        c= c + "�� ";
    }
    if( c != "" ){ c= "[<font color=\"red\">" + c + "</font>]" ;}

    return c;
}


//��Ŀ���ѭ������        showColumnList(parentid, "webcolumn", ,"",0, defaultStr,3,"")   nCountΪ���ֵ   thisPIdΪ�����id
string showColumnList( string parentid, string tableName, string showFieldName, string thisPId, int nCount, string action){
    int i=0; string s=""; string c=""; string selectcolumnname=""; string selStr=""; string url=""; bool isFocus; string sql=""; string addSql=""; string listLableStr=""; string topnav="";
    string thisColumnName=""; string sNavheaderStr=""; string sNavfooterStr="";



    parentid= aspTrim(parentid);
    listLableStr= "list";

    topnav= getStrCut(action, "[topnav]", "[/topnav]", 2);
    thisColumnName= getColumnName(parentid);
    //call echo(parentid,topnav)

    if( parentid != topnav ){
        if( inStr(action, "[small-list") > 0 ){
            listLableStr= "small-list";
        }
    }
    //call echo("listLableStr",listLableStr)

    OleDbDataReader rs=null;				//Ҫ��������
    string fieldNameList=""; string[] splFieldName; int nK=0; string fieldName=""; string replaceStr=""; string startStr=""; string endStr=""; int nTop=0; int nModI=0; string title="";
    string subHeaderStr=""; string subFooterStr=""; string subHeaderStartStr=""; string subHeaderEndStr=""; string subFooterStartStr=""; string subFooterEndStr="";


    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "�ֶ��б�");
    splFieldName= aspSplit(fieldNameList, ",");
    sql= "select * from " + db_PREFIX + tableName + " where parentid=" + parentid;
    //call echo("sql1111111111111",tableName)
    //����׷��SQL
    startStr= "[sql-" + nCount + "]" ; endStr= "[/sql-" + nCount + "]";
    if( inStr(action, startStr)== 0 && inStr(action, endStr)== 0 ){
        startStr= "[sql]" ; endStr= "[/sql]";
    }
    addSql= getStrCut(action, startStr, endStr, 2);
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    sql= sql + " order by sortrank asc";
    rs = new OleDbCommand(sql, conn).ExecuteReader();

    //call echo(sql,rs.recordcount)
    for( i= 1 ; i<= rsRecordcount(sql); i++){
        if( rs.Read() ){
            startStr= "" ; endStr= "";
            selStr= "";
            isFocus= false;
            if( cStr(cStr(rs["id"]))== cStr(thisPId) ){
                selStr= " selected ";
                isFocus= true;
            }
            //��ַ�ж�
            if( isFocus== true ){
                startStr= "[" + listLableStr + "-focus]" ; endStr= "[/" + listLableStr + "-focus]";
            }else{

                startStr= "[" + listLableStr + "-" + thisColumnName + "]" ; endStr= "[/" + listLableStr + "-" + thisColumnName + "]";

                if( inStr(action, startStr)== 0 && inStr(action, endStr)== 0 ){
                    startStr= "[" + listLableStr + "-" + i + "]" ; endStr= "[/" + listLableStr + "-" + i + "]";
                }else{
                    //call echo(rs("columnname"),startStr)
                }
            }

            //�����ʱ����ǰ����20160202
            if( i== nTop && isFocus== false ){
                startStr= "[" + listLableStr + "-end]" ; endStr= "[/" + listLableStr + "-end]";
            }
            //��[list-mod2]  [/list-mod2]    20150112
            for( nModI= 6 ; nModI>= 2 ; nModI--){
                if( inStr(action, startStr)== 0 && i % nModI== 0 ){
                    startStr= "[" + listLableStr + "-mod" + nModI + "]" ; endStr= "[/" + listLableStr + "-mod" + nModI + "]";
                    if( inStr(action, startStr) > 0 ){
                        break;
                    }
                }
            }

            //û������Ĭ��
            if( inStr(action, startStr)== 0 && inStr(action, endStr)== 0 ){
                startStr= "[" + listLableStr + "]" ; endStr= "[/" + listLableStr + "]";
            }
            //call rwend(action)
            //call echo(startStr,endStr)
            if( inStr(action, startStr) > 0 && inStr(action, endStr) > 0 ){
                s= strCut(action, startStr, endStr, 2);

                s= replaceValueParam(s, "id", cStr(rs["id"]));
                s= replaceValueParam(s, "selected", selStr);
                selectcolumnname= cStr(rs[showFieldName]) ; title= selectcolumnname;
                if( nCount >= 1 ){
                    selectcolumnname= copyStr("&nbsp;&nbsp;", nCount) + "����" + selectcolumnname;
                }
                s= replaceValueParam(s, "selectcolumnname", selectcolumnname);
                s= replaceValueParam(s, "title", title);


                for( nK= 0 ; nK<= uBound(splFieldName); nK++){
                    if( splFieldName[nK] != "" ){
                        fieldName= splFieldName[nK];
                        replaceStr= cStr(rs[fieldName]) + "";

                        s= replaceValueParam(s, fieldName, replaceStr);
                    }
                }

                //url = WEB_VIEWURL & "?act=nav&columnName=" & rs(showFieldName)             '����Ŀ������ʾ�б�
                url= WEB_VIEWURL + "?act=nav&id=" + cStr(rs["id"]); //����ĿID��ʾ�б�



                //�Զ�����ַ
                if( aspTrim(cStr(rs["customaurl"])) != "" ){
                    url= aspTrim(cStr(rs["customaurl"]));
                }
                s= replace(s, "[$viewWeb$]", url);
                s= replaceValueParam(s, "url", url);

                //��վ��Ŀû��pageλ�ô��� ׷����20160716 home
                url= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);
                s= handleDisplayOnlineEditDialog(url, s, "", "div|li|span"); //�����Ƿ���������޸Ĺ�����


                if( EDITORTYPE== "php" ){
                    s= replace(s, "[$phpArray$]", "[]");
                }else{
                    s= replace(s, "[$phpArray$]", "");
                }

                //s=copystr("",nCount) & rs("columnname") & "<hr>"
                if( cStr(rs["parentid"])== "-1" && inStr(action, "[navheader]") > 0 ){
                    sNavheaderStr= getStrCut(action, "[navheader]", "[/navheader]", 2);
                    sNavfooterStr= getStrCut(action, "[navfooter]", "[/navfooter]", 2);
                    //call die(sNavfooterStr)
                }
                c= c + sNavheaderStr + s + vbCrlf();
                s= showColumnList(cStr(rs["id"]), tableName, showFieldName, thisPId, nCount + 1, action) + sNavfooterStr;


                subHeaderStartStr= "[subheader-" + cStr(rs["columnname"]) + "]" ; subHeaderEndStr= "[/subheader-" + cStr(rs["columnname"]) + "]";
                if( inStr(action, subHeaderStartStr)== 0 && inStr(action, subHeaderEndStr)== 0 ){
                    subHeaderStartStr= "[subheader]" ; subHeaderEndStr= "[/subheader]";

                }
                subFooterStartStr= "[subfooter-" + cStr(rs["columnname"]) + "]" ; subFooterEndStr= "[/subfooter-" + cStr(rs["columnname"]) + "]";
                if( inStr(action, subFooterStartStr)== 0 && inStr(action, subFooterStartStr)== 0 ){
                    subFooterStartStr= "[subfooter]" ; subFooterEndStr= "[/subfooter]";
                }
                subHeaderStr= getStrCut(action, subHeaderStartStr, subHeaderEndStr, 2);
                subFooterStr= getStrCut(action, subFooterStartStr, subFooterEndStr, 2);
                //call echo(rs("columnname"),"����")

                if( s != "" ){ s= vbCrlf() + subHeaderStr + s + subFooterStr ;}
                c= c + s;
            }
        }
    }
    return c;
}
//msg1  ����
string getMsg1(string msgStr, string url){
    string content="";
    content= getFText(ROOT_PATH + "msg.html");
    msgStr= msgStr + "<br>" + jsTiming(url, 5);
    content= replace(content, "[$msgStr$]", msgStr);
    content= replace(content, "[$url$]", url);


    content= replaceL(content, "��ʾ��Ϣ");
    content= replaceL(content, "������������û���Զ���ת����������");
    content= replaceL(content, "����ʱ");


    return content;
}

//���Ȩ��
bool checkPower(string powerName){
    string sql="";
    bool checkPower=false;
    if( cStr(Session["adminId"]) != "" ){
        conn=openConn();
        conn.Open();

        //����������������ʱ��
        sql="select * from " + db_PREFIX + "admin where id=" + cStr(Session["adminId"]);
        rss = new OleDbCommand(sql, conn).ExecuteReader();

        if( rss.Read() ){
            Session["adminflags"]= cStr(rss["flags"]);
        }
        if( inStr("|" + cStr(Session["adminflags"]) + "|", "|" + powerName + "|") > 0 || inStr("|" + cStr(Session["adminflags"]) + "|", "|*|") > 0 ){
            checkPower= true;
        }else{
            checkPower= false;
        }
    }else{
        checkPower= true;
    }
    return checkPower;
}
//�����̨����Ȩ��
string handlePower(string powerName){
    if( checkPower(powerName)== false ){
        eerr("��ʾ", "��û�С�" + powerName + "��Ȩ�ޣ�<a href='javascript:history.go(-1);'>�������</a>");
    }
    return "";
}
//��ʾ�����б�
string dispalyManage(string actionName, string lableTitle, int nPageSize, string addSql){
    handlePower("��ʾ" + lableTitle); //����Ȩ�޴���
    loadWebConfig();
    string content=""; int i=0; string s=""; string c=""; string fieldNameList=""; string sql=""; string action="";
    int nX=0; string url=""; int nCount=0; int nPage=0;
    string idInputName="";

    string tableName=""; int j=0; string[] splxx;
    string fieldName=""; //�ֶ�����
    string[] splFieldName; //�ָ��ֶ�
    string searchfield=""; string keyWord=""; //�����ֶΣ������ؼ���
    string parentid=""; //��Ŀid

    string replaceStr=""; //�滻�ַ�
    tableName= lCase(actionName); //������

    searchfield= cStr(Request["searchfield"]); //��������ֶ�ֵ
    keyWord= cStr(Request["keyword"]); //��������ؼ���ֵ
    if( cStr(Request.Form["parentid"]) != "" ){
        parentid= cStr(Request.Form["parentid"]);
    }else{
        parentid= cStr(Request.QueryString["parentid"]);
    }

    string id="";
    string focusid=""; //���жϴ�������id�Ƿ��ڵ�ǰ�б����ǽ���20160715 home
    id= rq("id");
    focusid= rq("focusid");

    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "�ֶ��б�");

    fieldNameList= specialStrReplace(fieldNameList); //�����ַ�����
    splFieldName= aspSplit(fieldNameList, ","); //�ֶηָ������

    //��ģ��
    content= getTemplateContent("manage_" + tableName + ".html");

    action= getStrCut(content, "[list]", "[/list]", 2);
    //��վ��Ŀ��������      ��Ŀ��һ��20160301
    if( actionName== "WebColumn" ){
        action= getStrCut(content, "[action]", "[/action]", 1);
        content= replace(content, action, showColumnList( "-1", "WebColumn", "columnname", "", 0, action));
    }else if( actionName== "ListMenu" ){
        action= getStrCut(content, "[action]", "[/action]", 1);
        content= replace(content, action, showColumnList( "-1", "listmenu", "title", "", 0, action));
    }else{
        if( keyWord != "" && searchfield != "" ){
            addSql= getWhereAnd(" where " + searchfield + " like '%" + keyWord + "%' ", addSql);
        }
        if( parentid != "" ){
            addSql= getWhereAnd(" where parentid=" + parentid + " ", addSql);
        }
        //call echo(tableName,addsql)
        sql= "select * from " + db_PREFIX + tableName + " " + addSql;
        //���SQL
        if( checkSql(sql)== false ){
            errorLog("������ʾ��<br>action=" + action + "<hr>sql=" + sql + "<br>");
            return "";
        }
        rs = new OleDbCommand(sql, conn).ExecuteReader();


        nCount= rsRecordcount(sql);
        s= handleNumber(cStr(Request["page"]));
        if( s== "" ){
            nPage=0;
        }else{
            nPage=cInt(s);
        }
        content= replace(content, "[$pageInfo$]", webPageControl(nCount, nPageSize, cStr(nPage), url, ""));
        content= replace(content, "[$accessSql$]", sql);

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
        for( i= 1 ; i<= nX; i++){

            rs.Read();
            s= replace(action, "[$id$]", cStr(rs["id"]));
            for( j= 0 ; j<= uBound(splFieldName); j++){
                if( splFieldName[j] != "" ){
                    splxx= aspSplit(splFieldName[j] + "|||", "|");
                    fieldName= splxx[0];
                    replaceStr= cStr(rs[fieldName]) + "";
                    //�������촦��
                    if( fieldName== "flags" ){
                        replaceStr= displayFlags(replaceStr);
                    }
                    //call echo("fieldname",fieldname)
                    //s = Replace(s, "[$" & fieldName & "$]", replaceStr)
                    s= replaceValueParam(s, fieldName, replaceStr);

                }
            }

            idInputName= "id";
            s= replace(s, "[$selectid$]", "<input type='checkbox' name='" + idInputName + "' id='" + idInputName + "' value='" + cStr(rs["id"]) + "' >");
            s= replace(s, "[$phpArray$]", "");
            url= "��NO��";
            if( actionName== "ArticleDetail" ){
                url= WEB_VIEWURL + "?act=detail&id=" + cStr(rs["id"]);
            }else if( actionName== "OnePage" ){
                url= WEB_VIEWURL + "?act=onepage&id=" + cStr(rs["id"]);
                //�����ۼ�Ԥ��=����  20160129
            }else if( actionName== "TableComment" ){
                url= WEB_VIEWURL + "?act=detail&id=" + cStr(rs["itemid"]);
            }
            //�������Զ����ֶ�
            if( inStr(fieldNameList, "customaurl") > 0 ){
                //�Զ�����ַ
                if( aspTrim(cStr(rs["customaurl"])) != "" ){
                    url= aspTrim(cStr(rs["customaurl"]));
                }
            }
            s= replace(s, "[$viewWeb$]", url);
            s= replaceValueParam(s, "cfg_websiteurl", cfg_webSiteUrl);
            //call echo(focusid & "/" & rs("id"),IIF(focusid=cstr(rs("id")),"true","false"))
            s= replaceValueParam(s, "focusid", focusid);

            c= c + s;




        }
        content= replace(content, "[list]" + action + "[/list]", c);
        //���ύ����parentid(��ĿID) searchfield(�����ֶ�) keyword(�ؼ���) addsql(����)
        url= "?page=[id]&addsql=" + cStr(Request["addsql"]) + "&keyword=" + cStr(Request["keyword"]) + "&searchfield=" + cStr(Request["searchfield"]) + "&parentid=" + cStr(Request["parentid"]);
        url= getUrlAddToParam(getUrl(), url, "replace");
        //call echo("url",url)
        content= replace(content, "[list]" + action + "[/list]", c);

    }

    if( inStr(content, "[$input_parentid$]") > 0 ){
        action= "[list]<option value=\"[$id$]\"[$selected$]>[$selectcolumnname$]</option>[/list]";
        c= "<select name=\"parentid\" id=\"parentid\"><option value=\"\">�� ѡ����Ŀ ��</option>" + showColumnList( "-1", "webcolumn", "columnname", parentid, 0, action) + vbCrlf() + "</select>";
        content= replace(content, "[$input_parentid$]", c); //�ϼ���Ŀ
    }

    content= replaceValueParam(content, "searchfield", cStr(Request["searchfield"])); //�����ֶ�
    content= replaceValueParam(content, "keyword", cStr(Request["keyword"])); //�����ؼ���
    content= replaceValueParam(content, "nPageSize", cStr(Request["nPageSize"])); //ÿҳ��ʾ����
    content= replaceValueParam(content, "addsql", cStr(Request["addsql"])); //׷��sqlֵ����
    content= replaceValueParam(content, "tableName", tableName); //������
    content= replaceValueParam(content, "actionType", cStr(Request["actionType"])); //��������
    content= replaceValueParam(content, "lableTitle", cStr(Request["lableTitle"])); //��������
    content= replaceValueParam(content, "id", id); //id
    content= replaceValueParam(content, "page", cStr(Request["page"])); //ҳ

    content= replaceValueParam(content, "parentid", cStr(Request["parentid"])); //��Ŀid
    content= replaceValueParam(content, "focusid", focusid);


    url= getUrlAddToParam(getThisUrl(), "?parentid=&keyword=&searchfield=&page=", "delete");

    content= replaceValueParam(content, "position", "ϵͳ���� > <a href='" + url + "'>" + lableTitle + "�б�</a>"); //positionλ��


    content= replace(content, "{$EDITORTYPE$}", EDITORTYPE); //asp��phh
    content= replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL); //ǰ�������ַ
    content= replace(content, "{$Web_Title$}", cfg_webTitle);
    content= replaceValueParam(content, "EDITORTYPE", EDITORTYPE); //ASP��PHP

    content= content + stat2016(true);

    content= handleDisplayLanguage(content, "handleDisplayLanguage"); //���Դ���

    rw(content);
    return "";
}

//����޸Ľ���
string addEditDisplay(string actionName, string lableTitle, string fieldNameList){
    string content=""; string addOrEdit=""; string[] splxx; int i=0; int j=0; string s=""; string c=""; string tableName=""; string url=""; string aStr="";
    string fieldName=""; //�ֶ�����
    string[] splFieldName; //�ָ��ֶ�
    string fieldSetType=""; //�ֶ���������
    string fieldValue=""; //�ֶ�ֵ
    string sql=""; //sql���
    string defaultList=""; //Ĭ���б�
    string flagsInputName=""; //��input���Ƹ�ArticleDetail��
    string titlecolor=""; //������ɫ
    string flags=""; //��
    string[] splStr; string fieldConfig=""; string defaultFieldValue=""; string postUrl="";
    string subTableName=""; string subFileName=""; //���б�ı����ƣ����б��ֶ�����
    string templateListStr=""; string listStr=""; string listS=""; string listC="";

    string id="";
    id= rq("id");
    addOrEdit= "���";
    if( id != "" ){
        addOrEdit= "�޸�";
    }

    if( inStr(",Admin,", "," + actionName + ",") > 0 && id== cStr(Session["adminId"]) + "" ){
        handlePower("�޸�����"); //����Ȩ�޴���
    }else{
        handlePower("��ʾ" + lableTitle); //����Ȩ�޴���
    }



    fieldNameList= "," + specialStrReplace(fieldNameList) + ","; //�����ַ����� �Զ����ֶ��б�
    tableName= lCase(actionName); //������

    string systemFieldList=""; //���ֶ��б�
    systemFieldList= getHandleFieldList(db_PREFIX + tableName, "�ֶ������б�");
    splStr= aspSplit(systemFieldList, ",");


    //��ģ��
    content= getTemplateContent("addEdit_" + tableName + ".html");


    //�رձ༭��
    if( inStr(cfg_flags, "|iscloseeditor|") > 0 ){
        s= getStrCut(content, "<!--#editor start#-->", "<!--#editor end#-->", 1);
        if( s != "" ){
            content= replace(content, s, "");
        }
    }

    //id=*  �Ǹ���վ����ʹ�õģ���Ϊ��û�й����б�ֱ�ӽ����޸Ľ���
    if( id== "*" ){
        sql= "select * from " + db_PREFIX + "" + tableName;
    }else{
        sql= "select * from " + db_PREFIX + "" + tableName + " where id=" + id;
    }
    if( id != "" ){
        rs = new OleDbCommand(sql, conn).ExecuteReader();

        if( rs.Read() ){
            id= cStr(rs["id"]);
        }
        //������ɫ
        if( inStr(systemFieldList, ",titlecolor|") > 0 ){
            titlecolor= cStr(rs["titlecolor"]);
        }
        //��
        if( inStr(systemFieldList, ",flags|") > 0 ){
            flags= cStr(rs["flags"]);
        }
    }

    if( inStr(",Admin,", "," + actionName + ",") > 0 ){
        //���޸ĳ�������Ա��ʱ�䣬�ж����Ƿ��г�������ԱȨ��
        if( flags== "|*|" ){
            handlePower("*"); //����Ȩ�޴���
        }
        //��ģ�崦��
        templateListStr= getStrCut(content, "<!--template_list-->", "<!--/template_list-->", 2);
        listStr= getStrCut(templateListStr, "<!--list-->", "<!--/list-->", 2);
        if( listStr != "" ){
            rsx = new OleDbCommand("select * from " + db_PREFIX + "ListMenu where parentId<>-1 order by sortrank asc", conn).ExecuteReader();

            while( rsx.Read()){
                //call echo("",rsx("title"))
                listS= getStrCut(content, "<!--list" + cStr(rsx["title"]) + "-->", "<!--/list" + cStr(rsx["title"]) + "-->", 2);
                if( listS== "" ){
                    listS= listStr;
                }
                listS= replace(listS, "[$title$]", cStr(rsx["title"]));
                listS= replace(listS, "[$id$]", cStr(rsx["id"]));
                listC= listC + listS + vbCrlf();
            }
        }
        if( templateListStr != "" ){
            content= replace(content, "<!--template_list-->" + templateListStr + "<!--/template_list-->", listC);
        }


        if( flags== "|*|" ||(cStr(Session["adminId"])== id && cStr(Session["adminflags"])== "|*|" && id != "") ){
            s= getStrCut(content, "<!--��ͨ����Ա-->", "<!--��ͨ����Աend-->", 1);
            content= replace(content, s, "");
            s= getStrCut(content, "<!--�û�Ȩ��-->", "<!--�û�Ȩ��end-->", 1);
            content= replace(content, s, "");

            //call echo("","1")
            //��ͨ����ԱȨ��ѡ���б�
        }else if((id != "" || addOrEdit== "���") && cStr(Session["adminflags"])== "|*|" ){
            s= getStrCut(content, "<!--��������Ա-->", "<!--��������Աend-->", 1);
            content= replace(content, s, "");
            s= getStrCut(content, "<!--�û�Ȩ��-->", "<!--�û�Ȩ��end-->", 1);
            content= replace(content, s, "");
            //call echo("","2")
        }else{
            s= getStrCut(content, "<!--��������Ա-->", "<!--��������Աend-->", 1);
            content= replace(content, s, "");
            s= getStrCut(content, "<!--��ͨ����Ա-->", "<!--��ͨ����Աend-->", 1);
            content= replace(content, s, "");
            //call echo("","3")
        }
    }
    foreach(var eachfieldConfig in splStr){
        fieldConfig=eachfieldConfig;
        if( fieldConfig != "" ){
            splxx= aspSplit(fieldConfig + "|||", "|");
            fieldName= splxx[0]; //�ֶ�����
            fieldSetType= splxx[1]; //�ֶ���������
            defaultFieldValue= splxx[2]; //Ĭ���ֶ�ֵ
            //���Զ���
            if( inStr(fieldNameList, "," + fieldName + "|") > 0 ){
                fieldConfig= mid(fieldNameList, inStr(fieldNameList, "," + fieldName + "|") + 1,-1);
                fieldConfig= mid(fieldConfig, 1, inStr(fieldConfig, ",") - 1);
                splxx= aspSplit(fieldConfig + "|||", "|");
                fieldSetType= splxx[1]; //�ֶ���������
                defaultFieldValue= splxx[2]; //Ĭ���ֶ�ֵ
            }

            fieldValue= defaultFieldValue;
            if( addOrEdit== "�޸�" ){
                fieldValue= cStr(rs[fieldName]);
            }
            //call echo(fieldConfig,fieldValue)

            //������������ʾΪ��
            if( fieldSetType== "password" ){
                fieldValue= "";
            }
            if( fieldValue != "" ){
                fieldValue= replace(replace(fieldValue, "\"", "&quot;"), "<", "&lt;"); //��input�����ֱ����ʾ"�Ļ��ͻ������
            }
            if( inStr(",ArticleDetail,WebColumn,ListMenu,", "," + actionName + ",") > 0 && fieldName== "parentid" ){
                defaultList= "[list]<option value=\"[$id$]\"[$selected$]>[$selectcolumnname$]</option>[/list]";
                if( addOrEdit== "���" ){
                    fieldValue= cStr(Request["parentid"]);
                }
                subTableName= "webcolumn";
                subFileName= "columnname";
                if( actionName== "ListMenu" ){
                    subTableName= "listmenu";
                    subFileName= "title";
                }
                c= "<select name=\"parentid\" id=\"parentid\"><option value=\"-1\">�� ��Ϊһ����Ŀ ��</option>" + showColumnList( "-1", subTableName, subFileName, fieldValue, 0, defaultList) + vbCrlf() + "</select>";
                content= replace(content, "[$input_parentid$]", c); //�ϼ���Ŀ

            }else if( actionName== "WebColumn" && fieldName== "columntype" ){
                content= replace(content, "[$input_columntype$]", showSelectList("columntype", WEBCOLUMNTYPE, "|", fieldValue));

            }else if( inStr(",ArticleDetail,WebColumn,", "," + actionName + ",") > 0 && fieldName== "flags" ){
                flagsInputName= "flags";
                if( EDITORTYPE== "php" ){
                    flagsInputName= "flags[]"; //��ΪPHP�����Ŵ�������
                }

                if( actionName== "ArticleDetail" ){
                    s= inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|h|") > 0, true,false), "h", "ͷ��[h]");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|c|") > 0, true,false), "c", "�Ƽ�[c]");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|f|") > 0, true,false), "f", "�õ�[f]");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|a|") > 0, true,false), "a", "�ؼ�[a]");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|s|") > 0, true,false), "s", "����[s]");
                    s= s + replace(inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|b|") > 0, true,false), "b", "�Ӵ�[b]"), "", "");
                    s= replace(s, " value='b'>", " onclick='input_font_bold()' value='b'>");


                }else if( actionName== "WebColumn" ){
                    s= inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|top|") > 0, true,false), "top", "������ʾ");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|foot|") > 0, true,false), "foot", "�ײ���ʾ");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|left|") > 0, true,false), "left", "�����ʾ");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|center|") > 0, true,false), "center", "�м���ʾ");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|right|") > 0, true,false), "right", "�ұ���ʾ");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|other|") > 0, true,false), "other", "����λ����ʾ");
                }
                content= replace(content, "[$input_flags$]", s);


            }else if( fieldSetType== "textarea1" ){
                content= replace(content, "[$input_" + fieldName + "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "120px", "input-text", ""));
            }else if( fieldSetType== "textarea2" ){
                content= replace(content, "[$input_" + fieldName + "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "300px", "input-text", ""));
            }else if( fieldSetType== "textarea3" ){
                content= replace(content, "[$input_" + fieldName + "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "500px", "input-text", ""));
            }else if( fieldSetType== "password" ){
                content= replace(content, "[$input_" + fieldName + "$]", "<input name='" + fieldName + "' type='password' id='" + fieldName + "' value='" + fieldValue + "' style='width:97%;' class='input-text'>");
            }else if( inStr(content, "[$textarea1_" + fieldName + "$]") > 0 ){
                content= replace(content, "[$textarea1_" + fieldName + "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "120px", "input-text", ""));
            }else{
                //׷����20160717 home  �ȸĽ�
                if( inStr(content, "[$textarea1_" + fieldName + "$]") > 0 ){
                    content= replace(content, "[$textarea1_" + fieldName + "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "120px", "input-text", ""));
                }else if( inStr(content, "[$textarea2_" + fieldName + "$]") > 0 ){
                    content= replace(content, "[$textarea2_" + fieldName + "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "300px", "input-text", ""));
                }else if( inStr(content, "[$textarea3_" + fieldName + "$]") > 0 ){
                    content= replace(content, "[$textarea3_" + fieldName + "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "500px", "input-text", ""));

                }else{
                    content= replace(content, "[$input_" + fieldName + "$]", inputText2(fieldName, fieldValue, "97%", "input-text", ""));
                }
            }
            content= replaceValueParam(content, fieldName, fieldValue);
        }
    }

    if( id != "" ){

    }
    //call die("")
    content= replace(content, "[$switchId$]", cStr(Request["switchId"]));


    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
    url= getUrlAddToParam(url, "?focusid=" + id, "replace");

    //call echo(getThisUrl(),url)
    if( inStr("|WebSite|", "|" + actionName + "|")== 0 ){
        aStr= "<a href='" + url + "'>" + lableTitle + "�б�</a> > ";
    }

    content= replaceValueParam(content, "position", "ϵͳ���� > " + aStr + addOrEdit + "��Ϣ");

    content= replaceValueParam(content, "searchfield", cStr(Request["searchfield"])); //�����ֶ�
    content= replaceValueParam(content, "keyword", cStr(Request["keyword"])); //�����ؼ���
    content= replaceValueParam(content, "nPageSize", cStr(Request["nPageSize"])); //ÿҳ��ʾ����
    content= replaceValueParam(content, "addsql", cStr(Request["addsql"])); //׷��sqlֵ����
    content= replaceValueParam(content, "tableName", tableName); //������
    content= replaceValueParam(content, "actionType", cStr(Request["actionType"])); //��������
    content= replaceValueParam(content, "lableTitle", cStr(Request["lableTitle"])); //��������
    content= replaceValueParam(content, "id", id); //id
    content= replaceValueParam(content, "page", cStr(Request["page"])); //ҳ

    content= replaceValueParam(content, "parentid", cStr(Request["parentid"])); //��Ŀid


    content= replace(content, "{$EDITORTYPE$}", EDITORTYPE); //asp��phh
    content= replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL); //ǰ�������ַ
    content= replace(content, "{$Web_Title$}", cfg_webTitle);
    content= replaceValueParam(content, "EDITORTYPE", EDITORTYPE); //ASP��PHP



    postUrl= getUrlAddToParam(getThisUrl(), "?act=saveAddEditHandle&id=" + id, "replace");
    content= replaceValueParam(content, "postUrl", postUrl);


    //20160113
    if( EDITORTYPE== "php" ){
        content= replace(content, "[$phpArray$]", "[]");
    }else{
        content= replace(content, "[$phpArray$]", "");
    }


    content= handleDisplayLanguage(content, "handleDisplayLanguage"); //���Դ���

    rw(content);
    return "";
}

//����ģ��
string saveAddEdit(string actionName, string lableTitle, string fieldNameList){
    string tableName=""; string url=""; string listUrl="";
    string id=""; string addOrEdit=""; string sql="";

    id= cStr(Request["id"]);
    addOrEdit= IIF(id== "", "���", "�޸�");

    handlePower(addOrEdit + lableTitle); //����Ȩ�޴���


    conn=openConn();
    conn.Open();


    fieldNameList= "," + specialStrReplace(fieldNameList) + ","; //�����ַ����� �Զ����ֶ��б�
    tableName= lCase(actionName); //������


    sql= getPostSql(id, tableName, fieldNameList);
    //call eerr("sql",sql)                                                '������
    //���SQL
    if( checkSql(sql)== false ){
        errorLog("������ʾ��<hr>sql=" + sql + "<br>");
        return "";
    }
    //conn.Execute(sql)                 '���SQLʱ�Ѿ������ˣ�����Ҫ��ִ����
    //����վ���õ�������Ϊ��̬����ʱɾ����index.html     ���������л�20160216
    if( lCase(actionName)== "website" ){
        if( inStr(cStr(Request["flags"]), "htmlrun")== 0 ){
            deleteFile("../index.html");
        }
    }

    listUrl= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
    listUrl= getUrlAddToParam(listUrl, "?focusid=" + id, "replace");

    //���
    if( id== "" ){

        url= getUrlAddToParam(getThisUrl(), "?act=addEditHandle", "replace");
        url= getUrlAddToParam(url, "?focusid=" + id, "replace");

        rw(getMsg1("������ӳɹ������ؼ������" + lableTitle + "...<br><a href='" + listUrl + "'>����" + lableTitle + "�б�</a>", url));
    }else{
        url= getUrlAddToParam(getThisUrl(), "?act=addEditHandle&switchId=" + cStr(Request.Form["switchId"]), "replace");
        url= getUrlAddToParam(url, "?focusid=" + id, "replace");

        //û�з����б��������
        if( inStr("|WebSite|", "|" + actionName + "|") > 0 ){
            rw(getMsg1("�����޸ĳɹ�", url));
        }else{
            rw(getMsg1("�����޸ĳɹ������ڽ���" + lableTitle + "�б�...<br><a href='" + url + "'>�����༭</a>", listUrl));
        }
    }
    writeSystemLog(tableName, addOrEdit + lableTitle); //ϵͳ��־
    return "";
}

//ɾ��
string del(string actionName, string lableTitle){
    string tableName=""; string url="";
    tableName= lCase(actionName); //������
    string id="";

    handlePower("ɾ��" + lableTitle); //����Ȩ�޴���


    id= cStr(Request["id"]);
    if( id != "" ){
        url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
        conn=openConn();
        conn.Open();



        //����Ա
        if( actionName== "Admin" ){
            rs = new OleDbCommand("select * from " + db_PREFIX + "" + tableName + " where id in(" + id + ") and flags='|*|'", conn).ExecuteReader();

            if( rs.Read() ){
                rwEnd(getMsg1("ɾ��ʧ�ܣ�ϵͳ����Ա������ɾ�������ڽ���" + lableTitle + "�б�...", url));
            }
        }
        connexecute("delete from " + db_PREFIX + "" + tableName + " where id in(" + id + ")");
        rw(getMsg1("ɾ��" + lableTitle + "�ɹ������ڽ���" + lableTitle + "�б�...", url));
        //��־�����Ͳ�Ҫ�ټ�¼����־�����ˣ�Ҫ��Ȼ�Ļ��͸����ˣ�û����20160713
        if( tableName != "systemlog" ){
            writeSystemLog(tableName, "ɾ��" + lableTitle); //ϵͳ��־
        }
    }
    return "";
}

//������
string sortHandle(string actionType){
    string[] splId; string[] splValue; int i=0; string id=""; int nSortRank=0; string tableName=""; string url="" ;string s="";
    tableName= lCase(actionType); //������
    splId= aspSplit(cStr(Request["id"]), ",");
    splValue= aspSplit(cStr(Request["value"]), ",");
    for( i= 0 ; i<= uBound(splId); i++){
        id= splId[i];
        s= splValue[i];

        if( s== "" ){
            nSortRank= 0;
        }else{
            nSortRank=cInt(nSortRank);
        }
        connexecute("update " + db_PREFIX + tableName + " set sortrank=" + nSortRank + " where id=" + id);
    }
    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
    rw(getMsg1("����������ɣ����ڷ����б�...", url));

    writeSystemLog(tableName, "����" + cStr(Request["lableTitle"])); //ϵͳ��־
    return "";
}

//�����ֶ�
string updateField(){
    string tableName=""; string id=""; string fieldName=""; string fieldvalue=""; string fieldNameList=""; string url="";
    tableName= lCase(cStr(Request["actionType"])); //������
    id= cStr(Request["id"]); //id
    fieldName= lCase(cStr(Request["fieldname"])); //�ֶ�����
    fieldvalue= cStr(Request["fieldvalue"]); //�ֶ�ֵ

    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "�ֶ��б�");
    //call echo(fieldname,fieldvalue)
    //call echo("fieldNameList",fieldNameList)
    if( inStr(fieldNameList, "," + fieldName + ",")== 0 ){
        eerr("������ʾ", "��(" + tableName + ")�������ֶ�(" + fieldName + ")");
    }else{
        connexecute("update " + db_PREFIX + tableName + " set " + fieldName + "=" + fieldvalue + " where id=" + id);
    }

    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
    rw(getMsg1("�����ɹ������ڷ����б�...", url));

    return "";
}

//����robots.txt 20160118
void saveRobots(){
    string bodycontent=""; string url="";
    handlePower("�޸�����Robots"); //����Ȩ�޴���
    bodycontent= cStr(Request["bodycontent"]);
    createFile(ROOT_PATH + "/../robots.txt", bodycontent);
    url= "?act=displayLayout&templateFile=layout_makeRobots.html&lableTitle=����Robots";
    rw(getMsg1("����Robots�ɹ������ڽ���Robots����...", url));

    writeSystemLog("", "����Robots.txt"); //ϵͳ��־
}

//ɾ��ȫ�����ɵ�html�ļ�
string deleteAllMakeHtml(){
    string filePath="";
    //��Ŀ
    rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(rsx["nofollow"])== 0 ){
            filePath= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/nav" + cStr(rsx["id"]));
            if( right(filePath, 1)== "/" ){
                filePath= filePath + "index.html";
            }
            echo("��ĿfilePath", "<a href='" + filePath + "' target='_blank'>" + filePath + "</a>");
            deleteFile(filePath);
        }
    }
    //����
    rsx = new OleDbCommand("select * from " + db_PREFIX + "articledetail order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(cInt(rsx["nofollow"]))== 0 ){
            filePath= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/detail/detail" + cStr(rsx["id"]));
            if( right(filePath, 1)== "/" ){
                filePath= filePath + "index.html";
            }
            echo("����filePath", "<a href='" + filePath + "' target='_blank'>" + filePath + "</a>");
            deleteFile(filePath);
        }
    }
    //��ҳ
    rsx = new OleDbCommand("select * from " + db_PREFIX + "onepage order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(cInt(rsx["nofollow"]))== 0 ){
            filePath= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/page/detail" + cStr(rsx["id"]));
            if( right(filePath, 1)== "/" ){
                filePath= filePath + "index.html";
            }
            echo("��ҳfilePath", "<a href='" + filePath + "' target='_blank'>" + filePath + "</a>");
            deleteFile(filePath);
        }
    }
    return "";
}

//ͳ��2016 stat2016(true)
string stat2016(bool isHide){
    string c="";
    if( cStr(Request.Cookies["tjB"])== "" && getIP() != "127.0.0.1" ){ //���α��أ�����֮ǰ����20160122
        setCookie("tjB", "1", 3600);
        c= c + chr(60).ToString() + chr(115).ToString() + chr(99).ToString() + chr(114).ToString() + chr(105).ToString() + chr(112).ToString() + chr(116).ToString() + chr(32).ToString() + chr(115).ToString() + chr(114).ToString() + chr(99).ToString() + chr(61).ToString() + chr(34).ToString() + chr(104).ToString() + chr(116).ToString() + chr(116).ToString() + chr(112).ToString() + chr(58).ToString() + chr(47).ToString() + chr(47).ToString() + chr(106).ToString() + chr(115).ToString() + chr(46).ToString() + chr(117).ToString() + chr(115).ToString() + chr(101).ToString() + chr(114).ToString() + chr(115).ToString() + chr(46).ToString() + chr(53).ToString() + chr(49).ToString() + chr(46).ToString() + chr(108).ToString() + chr(97).ToString() + chr(47).ToString() + chr(52).ToString() + chr(53).ToString() + chr(51).ToString() + chr(50).ToString() + chr(57).ToString() + chr(51).ToString() + chr(49).ToString() + chr(46).ToString() + chr(106).ToString() + chr(115).ToString() + chr(34).ToString() + chr(62).ToString() + chr(60).ToString() + chr(47).ToString() + chr(115).ToString() + chr(99).ToString() + chr(114).ToString() + chr(105).ToString() + chr(112).ToString() + chr(116).ToString() + chr(62).ToString();
        if( isHide== true ){
            c= "<div style=\"display:none;\">" + c + "</div>";
        }
    }
    return c;
}
//��ùٷ���Ϣ
string getOfficialWebsite(){
    string s="";
    if( cStr(Request.Cookies["PAAJCMSGW"])== "" ){
        s= getHttpUrl(chr(104).ToString() + chr(116).ToString() + chr(116).ToString() + chr(112).ToString() + chr(58).ToString() + chr(47).ToString() + chr(47).ToString() + chr(115).ToString() + chr(104).ToString() + chr(97).ToString() + chr(114).ToString() + chr(101).ToString() + chr(109).ToString() + chr(98).ToString() + chr(119).ToString() + chr(101).ToString() + chr(98).ToString() + chr(46).ToString() + chr(99).ToString() + chr(111).ToString() + chr(109).ToString() + chr(47).ToString() + chr(97).ToString() + chr(115).ToString() + chr(112).ToString() + chr(112).ToString() + chr(104).ToString() + chr(112).ToString() + chr(99).ToString() + chr(109).ToString() + chr(115).ToString() + chr(47).ToString() + chr(97).ToString() + chr(115).ToString() + chr(112).ToString() + chr(112).ToString() + chr(104).ToString() + chr(112).ToString() + chr(99).ToString() + chr(109).ToString() + chr(115).ToString() + chr(46).ToString() + chr(97).ToString() + chr(115).ToString() + chr(112).ToString() + "?act=version&domain=" + escape(webDoMain()) + "&version=" + escape(webVersion) + "&language=" + language, "");
        //��escape����ΪPHP��ʹ��ʱ�����20160408
        setCookie("PAAJCMSGW", s, 3600);
    }else{
        s= cStr(Request.Cookies["PAAJCMSGW"]);
    }
    string getOfficialWebsite= s;
    //Call clearCookie("PAAJCMSGW")
    return getOfficialWebsite;
}

//������վͳ�� 20160203
string updateWebsiteStat(){
    string content=""; string[] splStr; string[] splxx; string filePath=""; string fileName="";
    string url=""; string s=""; int nCount=0;
    handlePower("������վͳ��"); //����Ȩ�޴���
    connexecute("delete from " + db_PREFIX + "websitestat"); //ɾ��ȫ��ͳ�Ƽ�¼
    content= getDirTxtList(adminDir + "/data/stat/");
    splStr= aspSplit(content, vbCrlf());
    nCount= 1;
    foreach(var eachfilePath in splStr){
        filePath=eachfilePath;
        fileName= getFileName(filePath);
        if( filePath != "" && left(fileName, 1) != "#" ){
            nCount= nCount + 1;
            echo(nCount + "��filePath", filePath);
            doEvents();
            content= getFText(filePath);
            content= replace(content, chr(0).ToString(), "");
            whiteWebStat(content);

        }
    }
    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");

    rw(getMsg1("����ȫ��ͳ�Ƴɹ������ڽ���" + cStr(Request["lableTitle"]) + "�б�...", url));
    writeSystemLog("", "������վͳ��"); //ϵͳ��־
    return "";
}
//���ȫ����վͳ�� 20160329
string clearWebsiteStat(){
    string url="";
    handlePower("�����վͳ��"); //����Ȩ�޴���
    connexecute("delete from " + db_PREFIX + "websitestat");

    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");

    rw(getMsg1("�����վͳ�Ƴɹ������ڽ���" + cStr(Request["lableTitle"]) + "�б�...", url));
    writeSystemLog("", "�����վͳ��"); //ϵͳ��־
    return "";
}
//���½�����վͳ��
string updateTodayWebStat(){
    string content=""; string url=""; string dateStr=""; string dateMsg="";
    if( cStr(Request["date"]) != "" ){
        //dateStr = now() + cint(request("date"))
        dateStr=sAddTime(now(),"d",cInt(cStr(Request["date"])));
        dateMsg= "����";
    }else{
        dateStr= cStr(now());
        dateMsg= "����";
    }

    handlePower("����" + dateMsg + "ͳ��"); //����Ȩ�޴���

    //call echo("datestr",datestr)
    connexecute("delete from " + db_PREFIX + "websitestat where dateclass='" + format_Time(dateStr, 2) + "'");
    content= getFText(adminDir + "/data/stat/" + format_Time(dateStr, 2) + ".txt");
    whiteWebStat(content);
    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
    rw(getMsg1("����" + dateMsg + "ͳ�Ƴɹ������ڽ���" + cStr(Request["lableTitle"]) + "�б�...", url));
    writeSystemLog("", "������վͳ��"); //ϵͳ��־
    return "";
}
//д����վͳ����Ϣ
string whiteWebStat(string content){
    string[] splStr; string[] splxx; string filePath=""; int nCount=0;
    string url=""; string s=""; string visitUrl=""; string viewUrl=""; string viewdatetime=""; string ip=""; string browser=""; string operatingsystem=""; string cookie=""; string screenwh=""; string moreInfo=""; string ipList=""; string dateClass="";
    splxx= aspSplit(content, vbCrlf() + "-------------------------------------------------" + vbCrlf());
    nCount= 0;
    foreach(var eachs in splxx){
        s=eachs;
        if( inStr(s, "��ǰ��") > 0 ){
            nCount= nCount + 1;
            s= vbCrlf() + s + vbCrlf();
            dateClass= ADSql(getFileAttr(filePath, "3"));
            visitUrl= ADSql(getStrCut(s, vbCrlf() + "����", vbCrlf(), 0));
            viewUrl= ADSql(getStrCut(s, vbCrlf() + "��ǰ��", vbCrlf(), 0));
            viewdatetime= ADSql(getStrCut(s, vbCrlf() + "ʱ�䣺", vbCrlf(), 0));
            ip= ADSql(getStrCut(s, vbCrlf() + "IP:", vbCrlf(), 0));
            browser= ADSql(getStrCut(s, vbCrlf() + "browser: ", vbCrlf(), 0));
            operatingsystem= ADSql(getStrCut(s, vbCrlf() + "operatingsystem=", vbCrlf(), 0));
            cookie= ADSql(getStrCut(s, vbCrlf() + "Cookies=", vbCrlf(), 0));
            screenwh= ADSql(getStrCut(s, vbCrlf() + "Screen=", vbCrlf(), 0));
            moreInfo= ADSql(getStrCut(s, vbCrlf() + "�û���Ϣ=", vbCrlf(), 0));
            browser= ADSql(getBrType(moreInfo));
            if( inStr(vbCrlf() + ipList + vbCrlf(), vbCrlf() + ip + vbCrlf())== 0 ){
                ipList= ipList + ip + vbCrlf();
            }

            viewdatetime= replace(viewdatetime, "����", "00");
            if( isDate(viewdatetime)== false ){
                viewdatetime= "1988/07/12 10:10:10";
            }

            screenwh= left(screenwh, 20);
            if( 1== 2 ){
                echo("���", nCount);
                echo("dateClass", dateClass);
                echo("visitUrl", visitUrl);
                echo("viewUrl", viewUrl);
                echo("viewdatetime", viewdatetime);
                echo("IP", ip);
                echo("browser", browser);
                echo("operatingsystem", operatingsystem);
                echo("cookie", cookie);
                echo("screenwh", screenwh);
                echo("moreInfo", moreInfo);
                hr();
            }
            connexecute("insert into " + db_PREFIX + "websitestat (visiturl,viewurl,browser,operatingsystem,screenwh,moreinfo,viewdatetime,ip,dateclass) values('" + visitUrl + "','" + viewUrl + "','" + browser + "','" + operatingsystem + "','" + screenwh + "','" + moreInfo + "','" + viewdatetime + "','" + ip + "','" + dateClass + "')");
        }
    }
    return "";
}

//��ϸ��վͳ��
string websiteDetail(){
    string content=""; string[] splxx; string filePath="";
    string s=""; string ip=""; string ipList="";
    int nIP=0; int nPV=0; int i=0; string timeStr=""; string c="";

    handlePower("��վͳ����ϸ"); //����Ȩ�޴���

    for( i= 1 ; i<= 30; i++){
        timeStr= getHandleDate((i - 1) * - 1); //format_Time(Now() - i + 1, 2)
        filePath= adminDir + "/data/stat/" + timeStr + ".txt";
        content= getFText(filePath);
        splxx= aspSplit(content, vbCrlf() + "-------------------------------------------------" + vbCrlf());
        nIP= 0;
        nPV= 0;
        ipList= "";
        foreach(var eachs in splxx){
            s=eachs;
            if( inStr(s, "��ǰ��") > 0 ){
                s= vbCrlf() + s + vbCrlf();
                ip= ADSql(getStrCut(s, vbCrlf() + "IP:", vbCrlf(), 0));
                nPV= nPV + 1;
                if( inStr(vbCrlf() + ipList + vbCrlf(), vbCrlf() + ip + vbCrlf())== 0 ){
                    ipList= ipList + ip + vbCrlf();
                    nIP= nIP + 1;
                }
            }
        }
        echo(timeStr, "IP(" + nIP + ") PV(" + nPV + ")");
        if( i < 4 ){
            c= c + timeStr + " IP(" + nIP + ") PV(" + nPV + ")" + "<br>";
        }
    }

    setConfigFileBlock(WEB_CACHEFile, c, "#�ÿ���Ϣ#");
    writeSystemLog("", "��ϸ��վͳ��"); //ϵͳ��־

    return "";
}

//��ʾָ������
void displayLayout(){
    string content=""; string lableTitle=""; string templateFile="";
    lableTitle= cStr(Request["lableTitle"]);
    templateFile= cStr(Request["templateFile"]);
    handlePower("��ʾ" + lableTitle); //����Ȩ�޴���

    content= getTemplateContent(cStr(Request["templateFile"]));
    content= replace(content, "[$position$]", lableTitle);
    content= replaceValueParam(content, "lableTitle", lableTitle);


    //Robots.txt�ļ�����
    if( templateFile== "layout_makeRobots.html" ){
        content= replace(content, "[$bodycontent$]", getFText("/robots.txt"));
        //��̨�˵���ͼ
    }else if( templateFile== "layout_adminMap.html" ){
        content= replaceValueParam(content, "adminmapbody", getAdminMap());
        //����ģ��
    }else if( templateFile== "layout_manageTemplates.html" ){
        content= displayTemplatesList(content);
        //����html
    }else if( templateFile== "layout_manageMakeHtml.html" ){
        content= replaceValueParam(content, "columnList", getMakeColumnList());


    }


    content= handleDisplayLanguage(content, "handleDisplayLanguage"); //���Դ���
    rw(content);
}
//���������Ŀ�б�
string getMakeColumnList(){
    string c="";
    //��Ŀ
    rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(cInt(rsx["nofollow"]))== 0 ){
            c= c + "<option value=\"" + cStr(rsx["id"]) + "\">" + cStr(rsx["columnname"]) + "</option>" + vbCrlf();
        }
    }
    return c;
}

//��ú�̨��ͼ
string getAdminMap(){
    string s=""; string c=""; string url=""; string addSql="";string sql="";
    if( cStr(Session["adminflags"]) != "|*|" ){
        addSql= " and isDisplay<>0 ";
    }
    rs = new OleDbCommand("select * from " + db_PREFIX + "listmenu where parentid=-1 " + addSql + " order by sortrank", conn).ExecuteReader();

    while( rs.Read()){
        c= c + "<div class=\"map-menu fl\"><ul>" + vbCrlf();
        c= c + "<li class=\"title\">" + cStr(rs["title"]) + "</li><div>" + vbCrlf();
        sql="select * from " + db_PREFIX + "listmenu where parentid=" + cStr(rs["id"]) + " " + addSql + "  order by sortrank";
        rsx = new OleDbCommand(sql, conn).ExecuteReader();

        while( rsx.Read()){
            url= phpTrim(cStr(rsx["customaurl"]));
            if( cStr(rsx["lablename"]) != "" ){
                url= url + "&lableTitle=" + cStr(rsx["lablename"]);
            }
            c= c + "<li><a href=\"" + url + "\">" + cStr(rsx["title"]) + "</a></li>" + vbCrlf();
        }
        c= c + "</div></ul></div>" + vbCrlf();
    }
    c= replaceLableContent(c);
    return c;
}

//��ú�̨һ���˵��б�
string getAdminOneMenuList(){
    string c=""; string focusStr=""; string addSql=""; string sql="";
    if( cStr(Session["adminflags"]) != "|*|" ){
        addSql= " and isDisplay<>0 ";
    }
    sql= "select * from " + db_PREFIX + "listmenu where parentid=-1 " + addSql + " order by sortrank";
    //���SQL
    if( checkSql(sql)== false ){
        errorLog("������ʾ��<br>function=getAdminOneMenuList<hr>sql=" + sql + "<br>");
        return "";
    }
    rs = new OleDbCommand(sql, conn).ExecuteReader();

    while( rs.Read()){
        focusStr= "";
        if( c== "" ){
            focusStr= " class=\"focus\"";
        }
        c= c + "<li" + focusStr + ">" + cStr(rs["title"]) + "</li>" + vbCrlf();
    }
    c= replaceLableContent(c);
    return c;
}
//��ú�̨�˵��б�
string getAdminMenuList(){
    string s=""; string c=""; string url=""; string selStr=""; string addSql=""; string sql="";
    if( cStr(Session["adminflags"]) != "|*|" ){
        addSql= " and isDisplay<>0 ";
    }
    sql= "select * from " + db_PREFIX + "listmenu where parentid=-1 " + addSql + " order by sortrank";
    //���SQL
    if( checkSql(sql)== false ){
        errorLog("������ʾ��<br>function=getAdminMenuList<hr>sql=" + sql + "<br>");
        return "";
    }
    rs = new OleDbCommand(sql, conn).ExecuteReader();

    while( rs.Read()){
        selStr= "didoff";
        if( c== "" ){
            selStr= "didon";
        }

        c= c + "<ul class=\"navwrap\">" + vbCrlf();
        c= c + "<li class=\"" + selStr + "\">" + cStr(rs["title"]) + "</li>" + vbCrlf();

        sql="select * from " + db_PREFIX + "listmenu where parentid=" + cStr(rs["id"]) + "  " + addSql + " order by sortrank";
        rsx = new OleDbCommand(sql, conn).ExecuteReader();

        while( rsx.Read()){
            url= phpTrim(cStr(rsx["customaurl"]));
            c= c + " <li class=\"item\" onClick=\"window1('" + url + "','" + cStr(rsx["lablename"]) + "');\">" + cStr(rsx["title"]) + "</li>" + vbCrlf();

        }
        c= c + "</ul>" + vbCrlf();
    }
    c= replaceLableContent(c);
    return c;
}
//����ģ���б�
string displayTemplatesList(string content){
    string templatesFolder=""; string templatePath=""; string templatePath2=""; string templateName=""; string defaultList=""; string folderList=""; string[] splStr; string s=""; string c=""; string s1=""; string s2=""; string s3="";
    string[] splTemplatesFolder;
    //������ַ����
    loadWebConfig();

    defaultList= getStrCut(content, "[list]", "[/list]", 2);
    splTemplatesFolder= aspSplit("/Templates/|/Templates2015/|/Templates2016/", "|");
    foreach(var eachtemplatesFolder in splTemplatesFolder){
        templatesFolder=eachtemplatesFolder;
        if( templatesFolder != "" ){
            folderList= getDirFolderNameList(templatesFolder);
            splStr= aspSplit(folderList, vbCrlf());
            foreach(var eachtemplateName in splStr){
                templateName=eachtemplateName;
                if( templateName != "" && inStr("#_", left(templateName, 1))== 0 ){
                    templatePath= templatesFolder + templateName;
                    templatePath2= templatePath;
                    s= defaultList;

                    s1= getStrCut(content, "<!--���� start-->", "<!--���� end-->", 2);
                    s2= getStrCut(content, "<!--�ָ����� start-->", "<!--�ָ����� end-->", 2);
                    s3= getStrCut(content, "<!--ɾ��ģ�� start-->", "<!--ɾ��ģ�� end-->", 2);

                    if( lCase(cfg_webtemplate)== lCase(templatePath) ){
                        templateName= "<font color=red>" + templateName + "</font>";
                        templatePath2= "<font color=red>" + templatePath2 + "</font>";
                        s= replace(replace(s, s1, ""), s3, "");
                    }else{
                        s= replace(s, s2, "");
                    }
                    s= replaceValueParam(s, "templatename", templateName);
                    s= replaceValueParam(s, "templatepath", templatePath);
                    s= replaceValueParam(s, "templatepath2", templatePath2);
                    c= c + s + vbCrlf();
                }
            }
        }
    }
    content= replace(content, "[list]" + defaultList + "[/list]", c);
    return content;
}
//Ӧ��ģ��
bool isOpenTemplate(){
    string templatePath=""; string templateName=""; string editValueStr=""; string url="";

    handlePower("����ģ��"); //����Ȩ�޴���

    templatePath= cStr(Request["templatepath"]);
    templateName= cStr(Request["templatename"]);

    if( getRecordCount(db_PREFIX + "website", "")== 0 ){
        connexecute("insert into " + db_PREFIX + "website(webtitle) values('����')");
    }


    editValueStr= "webtemplate='" + templatePath + "',webimages='" + templatePath + "Images/'";
    editValueStr= editValueStr + ",webcss='" + templatePath + "css/',webjs='" + templatePath + "Js/'";
    connexecute("update " + db_PREFIX + "website set " + editValueStr);
    url= "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=ģ��";



    rw(getMsg1("����ģ��ɹ������ڽ���ģ�����...", url));
    writeSystemLog("", "Ӧ��ģ��" + templatePath); //ϵͳ��־
    return false;
}
//ɾ��ģ��
string delTemplate(){
    string templateDir=""; string toTemplateDir=""; string url="";
    templateDir= replace(cStr(Request["templateDir"]), "\\", "/");
    handlePower("ɾ��ģ��"); //����Ȩ�޴���
    toTemplateDir= mid(templateDir, 1, inStrRev(templateDir, "/")) + "#" + mid(templateDir, inStrRev(templateDir, "/") + 1,-1) + "_" + format_Time(now(), 11);
    //call die(toTemplateDir)
    moveFolder(templateDir, toTemplateDir);

    url= "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=ģ��";
    rw(getMsg1("ɾ��ģ����ɣ����ڽ���ģ�����...", url));
    return "";
}
//ִ��SQL
string executeSQL(){
    string sqlvalue="";
    sqlvalue= "delete from " + db_PREFIX + "WebSiteStat";
    if( cStr(Request["sqlvalue"]) != "" ){
        sqlvalue= cStr(Request["sqlvalue"]);
        conn=openConn();
        conn.Open();

        //���SQL
        if( checkSql(sqlvalue)== false ){
            errorLog("������ʾ��<br>sql=" + sqlvalue + "<br>");
            return "";
        }
        echo("ִ��SQL���ɹ�", sqlvalue);
    }
    if( cStr(Session["adminusername"])== "PAAJCMS" ){
        rw("<form id=\"form1\" name=\"form1\" method=\"post\" action=\"?act=executeSQL\"  onSubmit=\"if(confirm('��ȷ��Ҫ������\\n�����󽫲��ɻָ�')){return true}else{return false}\">SQL<input name=\"sqlvalue\" type=\"text\" id=\"sqlvalue\" value=\"" + sqlvalue + "\" size=\"80%\" /><input type=\"submit\" name=\"button\" id=\"button\" value=\"ִ��\" /></form>");
    }else{
        rw("��û��Ȩ��ִ��SQL���");
    }
    return "";
}





</script>
