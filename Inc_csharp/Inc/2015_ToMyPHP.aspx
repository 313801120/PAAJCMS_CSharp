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
//��phpͨ��   �ҵĺ�̨

//�����վ�ײ�����aa
string XY_AP_WebSiteBottom(string action){
    string s="";string url="";
    if( inStr(cfg_webSiteBottom, "[$aoutadd$]") > 0 ){
        cfg_webSiteBottom= getDefaultValue(action); //���Ĭ������
        connexecute("update " + db_PREFIX + "website set websitebottom='" + ADSql(cfg_webSiteBottom) + "'");
    }

    s=cfg_webSiteBottom;
    //��վ�ײ�
    if( cStr(Request["gl"])== "edit" ){
        s= "<span>" + s + "</span>";
    }
    url= WEB_ADMINURL + "?act=addEditHandle&switchId=2&id=*&actionType=WebSite&lableTitle=վ������&n=" + getRnd(11);
    s= handleDisplayOnlineEditDialog(url, s, "", "span");

    return s;
}

//asp��php�汾
string XY_EDITORTYPE(string action){
    string aspValue="";string phpValue="";string aspxValue="";string s="";
    aspValue= lCase(rParam(action, "asp"));
    phpValue= lCase(rParam(action, "php"));
    aspxValue= lCase(rParam(action, "aspx")); 
    if( EDITORTYPE=="asp" ){
        s=aspValue;
    }else if( EDITORTYPE=="aspx" ){
        s=aspxValue;
    }else{
        s=phpValue;
    }
    return s;
}


//�����ļ�
string XY_Include(string action){
    string templateFilePath=""; string Block=""; string startStr=""; string endStr=""; string content="";
    templateFilePath= lCase(rParam(action, "File"));
    Block= lCase(rParam(action, "Block"));

    string findstr=""; string replaceStr=""; //�����ַ����滻�ַ�
    findstr= moduleFindContent(action, "findstr"); //���ҿ�
    replaceStr= moduleFindContent(action, "replacestr"); //���ҿ�

    templateFilePath= handleFileUrl(templateFilePath); //�����ļ�·��
    if( checkFile(templateFilePath)== false ){
        templateFilePath= cfg_webTemplate + "/" + templateFilePath;
    }
    content= getFText(templateFilePath);
    if( Block != "" ){
        startStr= "<!--#" + Block + " start#-->";
        endStr= "<!--#" + Block + " end#-->";
        if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
            content= strCut(content, startStr, endStr, 2);
        }
    }
    //�滻������������
    if( findstr != "" ){
        content= replace(content, findstr, replaceStr);
    }

    return content;
}

//��Ŀ�˵�
string XY_AP_ColumnMenu(string action){
    string defaultStr=""; string thisId=""; string parentid="";string c="";string s="";
    parentid= aspTrim(rParam(action, "parentid"));
    parentid=getColumnId(parentid);

    if( parentid== "" ){ parentid= "-1" ;}

    thisId= glb_columnId;
    if( thisId== "" ){ thisId= "-1" ;}
    defaultStr= getDefaultValue(action); //���Ĭ������

    defaultStr=defaultStr + "[topnav]"+ parentid +"[/topnav]";
    string XY_AP_ColumnMenu= showColumnList( parentid, "webcolumn", "columnname",thisId , 0, defaultStr);

    return XY_AP_ColumnMenu;
}



//��ʾ��Ŀ�б�
string XY_AP_ColumnList(string action){
    string sql=""; string flags=""; string addSql=""; string columnname="";
    sql= rParam(action, "sql");
    flags= rParam(action, "flags");
    addSql= rParam(action, "addSql");
    columnname= rParam(action, "columnname");
    if( flags != "" ){
        sql= " where flags like'%" + flags + "%'";
    }
    if( columnname != "" ){
        sql= getWhereAnd(sql, "where parentid=" + getColumnId(columnname));
        //call echo(sql,columnName)
    }
    //׷��sql
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    string XY_AP_ColumnList= XY_AP_GeneralList(action, "WebColumn", sql);

    return XY_AP_ColumnList;
}

//��ʾ�����б�
string XY_AP_ArticleList(string action){
    string sql=""; string addSql=""; string columnName=""; string columnId=""; string topNumb=""; string idRand=""; string[] splStr; string s=""; string columnIdList="";
    action= replaceGlobleVariable(action); //�������滻��ǩ
    sql= rParam(action, "sql");
    topNumb= rParam(action, "topNumb");


    //id���
    idRand= lCase(rParam(action, "rand"));
    if( idRand== "true" || idRand== "1" ){
        sql= sql + " where id in(" + getRandArticleId("", topNumb) + ")";
    }

    //��Ŀ���� ����Ŀ���鴦���� ģ���������[Array]CSS3[Array]HTML5
    s= rParam(action, "columnName");
    if( s== "" ){
        s= rParam(action, "did");
    }
    if( s != "" ){
        splStr= aspSplit(s, "[Array]");
        foreach(var eachcolumnName in splStr){
            columnName=eachcolumnName;
            columnId= getColumnId(columnName);
            if( columnId != "" ){
                if( columnIdList != "" ){
                    columnIdList= columnIdList + ",";
                }
                columnIdList= columnIdList + columnId;
            }
        }
    }
    if( columnIdList != "" ){
        sql= getWhereAnd(sql, "where parentId in(" + columnIdList + ")");
    }
    //׷��sql
    addSql= rParam(action, "addSql");
    //call echo("addsql",addsql)
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    sql= replaceGlobleVariable(sql);
    //call echo(RParam(action, "columnName") ,sql)
    return XY_AP_GeneralList(action, "ArticleDetail", sql);
}

//��ʾ�����б�
string XY_AP_CommentList(string action){
    string itemID=""; string sql=""; string addSql="";
    addSql= rParam(action, "addsql");
    itemID= rParam(action, "itemID");
    itemID= replaceGlobleVariable(itemID);

    if( itemID != "" ){
        sql= " where itemID=" + itemID;
    }
    //׷��sql
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    return XY_AP_GeneralList(action, "TableComment", sql);
}

//��ʾ����ͳ��
string XY_AP_SearchStatList(string action){
    string addSql="";
    addSql= rParam(action, "addSql");
    return XY_AP_GeneralList(action, "SearchStat", addSql);
}
//��ʾ��������
string XY_AP_Links(string action){
    string addSql="";
    addSql= rParam(action, "addSql");
    return XY_AP_GeneralList(action, "FriendLink", addSql);
}

string XY_GetColumnId(string action){
    string columnName="";
    columnName=rParam(action, "columnName");
    return getColumnId(columnName);
}
//ͨ����Ϣ�б�
string XY_AP_GeneralList(string action, string tableName, string addSql){
    string title=""; string topNumb=""; int nTop=0; bool isB; string sql="";
    string columnName=""; string columnEnName=""; string aboutcontent=""; string bodyContent=""; string showTitle="";
    string bannerImage=""; string smallImage=""; string bigImage=""; string id="";
    string defaultStr=""; int i=0; int j=0; string s=""; string c=""; string startStr=""; string endStr=""; string url="";
    string sNoFollow=""; //��׷�� 20141222
    defaultStr= getDefaultValue(action); //���Ĭ������
    int nModI=0; //��ѭ��20150112
    sNoFollow= aspTrim(lCase(rParam(action, "noFollow"))); //��׷��
    string lableTitle=""; //�������
    string target=""; //a���Ӵ�Ŀ�귽ʽ
    string adddatetime=""; //���ʱ��
    bool isFocus;
    string fieldNameList=""; //�ֶ��б�
    string abcolorStr=""; //A�Ӵֺ���ɫ
    string atargetStr=""; //A���Ӵ򿪷�ʽ
    string atitleStr=""; //A���ӵ�title20160407
    string anofollowStr=""; //A���ӵ�nofollow
    string[] splFieldName; string fieldName=""; string replaceStr=""; int nI2=0;string idPage="";


    tableName=lCase(tableName);		//תСд
    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "�ֶ��б�");
    splFieldName= aspSplit(fieldNameList, ",");

    topNumb= rParam(action, "topNumb");
    if( topNumb != "" ){
        nTop= cInt(topNumb);
    }else{
        nTop= 999;
    }
    if( sql== "" ){
        if( topNumb != "" ){
            topNumb= " top " + topNumb + " ";
        }
        sql= "Select " + topNumb + "* From " + db_PREFIX + tableName;
    }
    //׷��sql
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    sql= replaceGlobleVariable(sql); //�滻ȫ�ֱ���

    //���SQL
    if( checkSql(sql)== false ){
        errorLog("������ʾ��<br>action=" + action + "<hr>sql=" + sql + "<br>");
        return "";
    }
    rs = new OleDbCommand(sql, conn).ExecuteReader();

    //call echo(nTop,rs.RecordCount)
    for( i= 1 ; i<= rsRecordcount(sql); i++){
        startStr= "" ; endStr= "";
        //call echo(sql,i & "," & nTop)
        if( i > nTop ){
            break;
        }
        //#��PHP��$rs=mysql_fetch_array($rsObj);                                            //��PHP�ã���Ϊ�� asptophpת��������
        rs.Read();
        isFocus= false; //����Ϊ��
        id= cStr(rs["id"]);
        //��������
        if( tableName== "webcolumn" ){
            if( isMakeHtml== true ){
                url= getRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "/nav" + cStr(rs["id"]));
            }else{
                url= handleWebUrl("?act=nav&columnName=" + cStr(rs["columnname"])); //��׷��gl�Ȳ���
                if( cStr(rs["customaurl"]) != "" ){
                    url= cStr(rs["customaurl"]);
                    url= replaceGlobleVariable(url);
                }
            }
            //ȫ����Ŀ����Ϊ����Ϊ�Զ���λ��ҳ ׷��(20160128)
            if( glb_columnName== "" && cStr(rs["columntype"])== "��ҳ" ){
                glb_columnName= cStr(rs["columnname"]);
            }
            if( cStr(rs["columnname"])== glb_columnName ){
                isFocus= true;
            }
            //�����¡�
        }else if( tableName== "articledetail" ){
            if( isMakeHtml== true ){
                url= getRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "detail/detail" + cStr(rs["id"]));
            }else{
                url= handleWebUrl("?act=detail&id=" + cStr(rs["id"])); //��׷��gl�Ȳ���
                if( cStr(rs["customaurl"]) != "" ){
                    url= cStr(rs["customaurl"]);
                }
            }
            //����
        }else if( tableName== "tablecomment" ){

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

        //�����ж�(����Ŀ�����õ�)
        if( isFocus== true ){
            startStr= "[list-focus]" ; endStr= "[/list-focus]";
        }else if( inStr(defaultStr, "[list-start]")>0 && inStr(defaultStr, "[/list-start]")>0 ){
            startStr= "[list-start]" ; endStr= "[/list-start]";
        }else{
            startStr= "[list-" + i + "]" ; endStr= "[/list-" + i + "]";
        }

        //�����ʱ����ǰ����20160202
        if( i== nTop && isFocus== false ){
            if( inStr(defaultStr, "[list-end]")>0 && inStr(defaultStr, "[/list-end]")>0 ){
                startStr= "[list-end]" ; endStr= "[/list-end]";
            }
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
        if( inStr(defaultStr, startStr)== 0 ){
            startStr= "[list]" ; endStr= "[/list]";
        }

        if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){

            s= strCut(defaultStr, startStr, endStr, 2);

            s= replaceValueParam(s, "i", i); //ѭ�����
            s= replaceValueParam(s, "���", i); //ѭ�����
            s= replaceValueParam(s, "id", cStr(rs["id"])); //id��� ��Ϊ����ֶ��������id
            s= replaceValueParam(s, "url", url); //��ַ
            s= replaceValueParam(s, "aurl", "href=\"" + url + "\""); //��ַ
            s= replaceValueParam(s, "abcolor", abcolorStr); //A���Ӽ���ɫ��Ӵ�
            s= replaceValueParam(s, "atitle", atitleStr); //A����title
            s= replaceValueParam(s, "anofollow", anofollowStr); //A����nofollow
            s= replaceValueParam(s, "atarget", atargetStr); //A���Ӵ򿪷�ʽ



            for( nI2= 0 ; nI2<= uBound(splFieldName); nI2++){
                if( splFieldName[nI2] != "" ){
                    fieldName= splFieldName[nI2];
                    replaceStr= cStr(rs[fieldName]) + "";
                    s= replaceValueParam(s, fieldName, replaceStr);
                }
            }


            //��ʼλ�ü�Dialog����
            startStr= "[list-" + i + " startdialog]" ; endStr= "[/list-" + i + " startdialog]";
            if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
                s= strCut(defaultStr, startStr, endStr, 2) + s;
            }
            //����λ�ü�Dialog����
            startStr= "[list-" + i + " enddialog]" ; endStr= "[/list-" + i + " enddialog]";
            if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
                s= s + strCut(defaultStr, startStr, endStr, 2);
            }

            //�ӿ���
            //��������
            if( tableName== "webcolumn" ){
                url= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);

                //�����ۡ�
            }else if( tableName=="tablecomment" ){
                idPage=getThisIdPage(db_PREFIX + tableName ,cStr(rs["id"]),10);
                url= WEB_ADMINURL + "?act=addEditHandle&actionType=TableComment&lableTitle=����&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page="+ idPage +"&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);

                //�����¡�
            }else if( tableName== "articledetail" ){
                idPage=getThisIdPage(db_PREFIX + tableName ,cStr(rs["id"]),10);
                url= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page="+ idPage +"&parentid="+ cStr(rs["parentid"]) +"&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);


                s= replaceValueParam(s, "columnurl", getColumnUrl(cStr(rs["parentid"]), "")); //���¶�Ӧ��ĿURL 20160304
                s= replaceValueParam(s, "columnname", getColumnName(cStr(rs["parentid"]))); //���¶�Ӧ��Ŀ���� 20160304
            }
            s= handleDisplayOnlineEditDialog(url, s, "", "div|li|span"); //�����Ƿ���������޸Ĺ�����
            c= c + s;
        }
    }

    //��ʼ���ݼ�Dialog����
    startStr= "[dialog start]" ; endStr= "[/dialog start]";
    if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
        c= strCut(defaultStr, startStr, endStr, 2) + c;
    }
    //�������ݼ�Dialog����
    startStr= "[dialog end]" ; endStr= "[/dialog end]";
    if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
        c= c + strCut(defaultStr, startStr, endStr, 2);
    }
    return c;
}


//�����ñ�����
string XY_handleGetTableBody(string action, string tableName, string fieldParamName, string defaultFileName, string adminUrl){
    string url=""; string content=""; string id=""; string sql=""; string addSql=""; string fieldName=""; string fieldParamValue=""; string fieldNameList=""; string sLen=""; string delHtmlYes=""; string trimYes="";string defaultStr="";
    string snoisonhtml="";string intoFieldStr="";string valuesStr="";
    fieldName= rParam(action, "fieldname"); //�ֶ�����
    snoisonhtml= rParam(action, "noisonhtml");					 //������html
    //dim nonull nonull=RParam(action, "noisonhtml")							'���ݲ���Ϊ��20160716 home

    if( snoisonhtml=="true" ){
        intoFieldStr=",isonhtml";
        valuesStr=",0";
    }

    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "�ֶ��б�");
    //�ֶ����Ʋ�Ϊ�գ�����Ҫ�ڱ��ֶ���
    if( fieldName== "" || inStr(fieldNameList, "," + fieldName + ",")== 0 ){
        fieldName= defaultFileName;
    }
    fieldName= lCase(fieldName); //תΪСд����Ϊ��PHP����ȫСд��

    fieldParamValue= rParam(action, fieldParamName); //��ȡ�ֶ�����
    id= handleNumber(rParam(action, "id")); //���ID
    addSql= " where " + fieldParamName + "='" + fieldParamValue + "'";
    if( id != "" ){
        addSql= " where id=" + id;
    }

    content= getDefaultValue(action) ;defaultStr=content; //���Ĭ������
    sql= "select * from " + db_PREFIX + tableName + addSql;
    rs = new OleDbCommand(sql, conn).ExecuteReader();

    if( rs.Read() ){
        id= cStr(rs["id"]);
        content= cStr(rs[fieldName]);
        if( len(content)<=0 ){
            content=defaultStr;
            connexecute("update " + db_PREFIX + tableName + " set " + fieldName + "='"+ content +"' where id=" + cStr(rs["id"]));
        }
    }else{
        //�Զ���� 20160113
        if( rParam(action, "autoadd")== "true" ){
            connexecute("insert into " + db_PREFIX + tableName + " (" + fieldParamName + "," + fieldName + intoFieldStr + ") values('" + fieldParamValue + "','" + ADSql(content) + "'"+ valuesStr +")");
        }
    }

    //ɾ��Html
    delHtmlYes= rParam(action, "delHtml"); //�Ƿ�ɾ��Html
    if( delHtmlYes== "true" ){ content= replace(delHtml(content), "<", "&lt;") ;}//HTML����
    //ɾ�����߿ո�
    trimYes= rParam(action, "trim"); //�Ƿ�ɾ�����߿ո�
    if( trimYes== "true" ){ content= trimVbCrlf(content) ;}

    //��ȡ�ַ�����
    sLen= rParam(action, "len"); //�ַ�����ֵ
    sLen= handleNumber(sLen);
    //If sLen<>"" Then ReplaceStr = CutStr(ReplaceStr,sLen,"null")' Left(ReplaceStr,sLen)
    if( sLen != "" ){ content= cutStr(content, cInt(sLen), "...") ;}//Left(ReplaceStr,sLen)


    if( id== "" ){
        id= XY_AP_GetFieldValue("", sql, "id");
    }
    url= adminUrl + "&id=" + id + "&n=" + getRnd(11);
    if( cStr(Request["gl"])== "edit" ){
        content= "<span>" + content + "</span>";
    }

    //call echo(sql,url)
    content= handleDisplayOnlineEditDialog(url, content, "", "span");
    string XY_handleGetTableBody= content;

    return XY_handleGetTableBody;
}

//��õ�ҳ����
string XY_AP_GetOnePageBody(string action){
    string adminUrl="";
    adminUrl= WEB_ADMINURL + "?act=addEditHandle&actionType=OnePage&lableTitle=��ҳ&nPageSize=10&page=&switchId=2";
    return XY_handleGetTableBody(action, "onepage", "title", "bodycontent", adminUrl);
}

//��õ�������
string XY_AP_GetColumnBody(string action){
    string adminUrl="";
    adminUrl= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&switchId=2";
    return XY_handleGetTableBody(action, "webcolumn", "columnname", "bodycontent", adminUrl);
}

//��ʾ��������
string XY_AP_GetArticleBody(string action){
    string adminUrl="";
    adminUrl= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&switchId=2";
    return XY_handleGetTableBody(action, "articledetail", "title", "bodycontent", adminUrl);
}


//�����ĿURL
string XY_GetColumnUrl(string action){
    string columnName=""; string url="";
    columnName= rParam(action, "columnName");
    url= getColumnUrl(columnName, "name");
    //handleWebUrl  �ж�gl����

    //If Request("gl") <> "" Then
    //    url = url & "&gl=" & Request("gl")
    //End If
    return url;
}
//��õ���url ����֮ǰ��վ����ʱ�õ� 20160722  ����ɾ��
string XY_GetNavUrl(string action){
    string columnName=""; string url="";string did="";string sid="";
    did= rParam(action, "did");
    sid= rParam(action, "sid");
    columnName=did;
    if( sid!="" ){
        columnName=sid;
    }
    url= getColumnUrl(columnName, "name");
    //handleWebUrl  �ж�gl����

    //If Request("gl") <> "" Then
    //    url = url & "&gl=" & Request("gl")
    //End If
    return url;
}

//�������URL
string XY_GetArticleUrl(string action){
    string title=""; string url="";
    title= rParam(action, "title");
    url= getArticleUrl(title);
    //If Request("gl") <> "" Then
    //    url = url & "&gl=" & Request("gl")
    //End If
    return url;
}

//��õ�ҳURL
string XY_GetOnePageUrl(string action){
    string title=""; string url="";
    title= rParam(action, "title");
    url= getOnePageUrl(title);
    //If Request("gl") <> "" Then
    //    url = url & "&gl=" & Request("gl")
    //End If
    return url;
}


//��õ����ֶ�����
string XY_AP_GetFieldValue(string action, string sql, string fieldName){
    string title=""; string content="";
    rs = new OleDbCommand(sql, conn).ExecuteReader();

    if( rs.Read() ){
        content= cStr(rs[fieldName]);
    }
    return content;
}


//Js����վͳ��
string XY_JsWebStat(string action){
    string s=""; string fileName=""; string sType="";
    sType= rParam(action, "stype");
    fileName= aspTrim(rParam(action, "fileName"));
    if( fileName== "" ){
        fileName= "[$WEB_VIEWURL$]?act=webstat&stype=" + sType;
    }
    fileName= replace(fileName, "/", "\\/");
    s= "<script>document.writeln(\"<script src=\\'" + fileName + "&GoToUrl=\"";
    s= s + "+escape(document.referrer)+\"&ThisUrl=\"+escape(window.location.href)+\"&screen=\"+escape(window.screen.width+\"x\"+window.screen.height)";
    s= s + "+\"&co=\"+escape(document.cookie)"; //�ռ�cookie ����Ҫ�����ε�
    s= s + "+\" \\'><\\/script>\");</"+"script>";
    return s;
}



//��ͨ����A
string XY_HrefA(string action){
    string content=""; string Href=""; string c=""; string AContent=""; string AType=""; string url=""; string title="";
    action= handleInModule(action, "start");
    content= rParam(action, "Content");
    AType= rParam(action, "Type");
    if( AType== "�ղ�" ){
        //��һ�ַ���
        //Url = "window.external.addFavorite('"& WebUrl &"','"& WebTitle &"')"
        url= "shoucang(document.title,window.location)";
        c= "<a href='javascript:;' onClick=\"" + url + "\" " + setHtmlParam(action, "target|title|alt|id|class|style") + ">" + content + "</a>";
    }else if( AType== "��Ϊ��ҳ" ){
        //��һ�ַ���
        //Url = "var strHref=window.location.href;this.style.behavior='url(#default#homepage)';this.setHomePage('"& WebUrl &"');"
        url= "SetHome(this,window.location)";
        c= "<a href='javascript:;' onClick=\"" + url + "\"" + setHtmlParam(action, "target|title|alt|id|class|style") + ">" + content + "</a>";
    }else{
        content= rParam(action, "Title");
    }

    content= handleInModule(content, "end");
    if( c== "" ){ c= "<a" + setHtmlParam(action, "href|target|title|alt|id|class|rel|style") + ">" + content + "</a>" ;}

    return c;
}



//����20151231
string XY_Layout(string action){
    string layoutName=""; string s=""; string c=""; string sourceStr=""; string replaceStr=""; string[] splSource; string[] splReplace; int i=0;

    layoutName= rParam(action, "layoutname");
    rs = new OleDbCommand("select * from " + db_PREFIX + "weblayout where layoutname='" + layoutName + "'", conn).ExecuteReader();

    if( rs.Read() ){
        c= cStr(rs["bodycontent"]);

        sourceStr= cStr(rs["sourcestr"]); //Դ���� ���滻����
        replaceStr= cStr(rs["replacestr"]); //�滻����
        splSource= aspSplit(sourceStr, "[Array]"); //Դ��������
        splReplace= aspSplit(replaceStr, "[Array]"); //�滻��������

        for( i= 0 ; i<= uBound(splSource); i++){
            sourceStr= splSource[i];
            replaceStr= splReplace[i];
            if( sourceStr != "" ){
                c= replace(c, sourceStr, replaceStr);
                //call echo(sourceStr,replaceStr)
                //call echo(c,instr(c,sourcestr))
            }
        }
        //call rwend(c)
    }
    return c;
}

//ģ��20151231
string XY_Module(string action){
    string moduleName=""; string s=""; string c=""; string sourceStr=""; string replaceStr=""; string[] splSource; string[] splReplace; int i=0;
    moduleName= rParam(action, "modulename");
    rs = new OleDbCommand("select * from " + db_PREFIX + "webmodule where modulename='" + moduleName + "'", conn).ExecuteReader();

    if( rs.Read() ){
        c= cStr(rs["bodycontent"]);

        sourceStr= rParam(action, "sourceStr"); //Դ���� ���滻����
        replaceStr= rParam(action, "replaceStr"); //�滻����

        splSource= aspSplit(sourceStr, "[Array]"); //Դ��������
        splReplace= aspSplit(replaceStr, "[Array]"); //�滻��������

        for( i= 0 ; i<= uBound(splSource); i++){
            sourceStr= splSource[i];
            replaceStr= splReplace[i];
            if( sourceStr != "" ){
                c= replace(c, sourceStr, replaceStr);
                //call echo(sourceStr,replaceStr)
                //call echo(c,instr(c,sourcestr))
            }
        }
        //call rwend(c)
    }
    return c;
}

//��ʾ������20160127
string XY_DisplayWrap( string action){
    string content="";
    content= getDefaultValue(action);
    return content;
}




//Ƕ�ױ��� ����
string XY_getLableValue(string action){
    string title=""; string content=""; string c="";
    //call echo("Action",Action)
    title= rParam(action, "title");
    content= rParam(action, "content");
    c= c + "title=" + getContentRunStr(title) + "<hr>";
    c= c + "content=" + getContentRunStr(content) + "<hr>";
    string XY_getLableValue= c;
    echo("title", title);
    return "��title=����" + title + "��";
}
//���������������������б�
string XY_TitleInSearchEngineList(string action){
    string title=""; string sType="";string divclass="";string spanclass="";string s="";string c="";

    title= rParam(action, "title");
    sType= rParam(action, "sType");
    divclass= rParam(action, "divclass");
    spanclass= rParam(action, "spanclass");

    s="<strong>������ڡ�" + title + "��</strong>";
    if( divclass!="" ){
        s="<div class=\""+ divclass +"\">"+ s +"</div>";
    }else if( spanclass!="" ){
        s="<span class=\""+ spanclass +"\">"+ s +"</span>" + "<br>";
    }else{
        s=s + "<br>";
    }
    c= c + s + vbCrlf();
    c= c + "<ul class=\"list\"> " + vbCrlf();
    c= c + "<li><a href=\"https://www.baidu.com/s?ie=gb2312&word=" + title + "\" rel=\"nofollow\" target=\"_blank\">��baidu�������ڰٶ�������(" + title + ")</a></li>" + vbCrlf();
    c= c + "<li><a href=\"http://www.haosou.com/s?ie=gb2312&q=" + title + "\" rel=\"nofollow\" target=\"_blank\">��haosou�������ں���������(" + title + ")</a></li>" + vbCrlf();
    c= c + "<li><a href=\"https://search.yahoo.com/search;_ylt=A86.JmbkJatWH5YARmebvZx4?toggle=1&cop=mss&ei=gb2312&fr=yfp-t-901&fp=1&p=" + title + "\" rel=\"nofollow\" target=\"_blank\">��yahoo���������Ż�������(" + title + ")</a></li>" + vbCrlf();

    c= c + "<li><a href=\"https://www.sogou.com/sogou?ie=utf8&query=" + toUTFChar(title) + "\" rel=\"nofollow\" target=\"_blank\">��sogou���������ѹ�������(" + title + ")</a></li>" + vbCrlf();
    c= c + "<li><a href=\"http://www.youdao.com/search?ue=utf8&q=" + toUTFChar(title) + "\" rel=\"nofollow\" target=\"_blank\">��youdao���������е�������(" + title + ")</a></li>" + vbCrlf();
    c= c + "<li><a href=\"http://search.yam.com/Search/Web/DefaultKSA.aspx?SearchType=web&l=0&p=0&k=" + toUTFChar(title) + "\" rel=\"nofollow\" target=\"_blank\">��yam����(google�ṩ����)����ެ����������(" + title + ")</a></li>" + vbCrlf();


    c= c + "<li><a href=\"http://cn.bing.com/search?q=" + toUTFChar(title) + "\" rel=\"nofollow\" target=\"_blank\">��bing�������ڱ�Ӧ������(" + title + ")</a></li>" + vbCrlf();
    c= c + "</ul>" + vbCrlf();

    return c;
}

//URL����
string XY_escape(string action){
    string content="";
    content= rParam(action, "content");
    return escape(content);
}
//URL����
string XY_unescape(string action){
    string content="";
    content= rParam(action, "content");
    return escape(content);
}
//�����ַ
string XY_getUrl(string action){
    string stype="";
    stype= rParam(action, "stype");
    return getThisUrlNoParam();
}



//Forѭ������
string XY_ForArray(string action){
    string sListArray=""; string sSplitStr=""; string defaultStr=""; string[] splStr; int nForI=0; string title=""; string s=""; string c=""; string sNLoop="";
    sListArray= atRParam(action, "arraylist"); //atRParam��ý���������������滻��������
    sSplitStr= rParam(action, "splitstr");
    sNLoop= rParam(action, "nloop"); //ѭ����

    if( sListArray== "" && sNLoop!="" ){
        sListArray= copyStr("ѭ��" + sSplitStr, cInt(sNLoop));
    }

    defaultStr= getDefaultValue(action);

    splStr= aspSplit(sListArray, sSplitStr);
    for( nForI= 0 ; nForI<= uBound(splStr); nForI++){
        title= splStr[nForI];
        if( title != "" ){
            s= defaultStr;
            s= replaceValueParam(s, "fortitle", title);
            s= replaceValueParam(s, "forid", nForI+1);
            s= replaceValueParam(s, "fori", nForI);
            s= replaceValueParam(s, "forcount", uBound(splStr) + 1);
            c= c + s;
        }
    }
    return c;
}

//����λ����ʾ��Ϣ{}Ϊ�ж�����
string XY_detailPosition(string action){
    string c="";string stype="";
    stype= rParam(action, "stype");
    stype="|"+ stype +"|";
    c= "<a href=\""+ getColumnUrl("��ҳ","type") +"\">��ҳ</a>";
    //call echo("type",getColumnUrl("��ҳ","type"))
    if( glb_columnName != "" ){
        c= c + " >> <a href=\"" + getColumnUrl(glb_columnName, "name") + "\">" + glb_columnName + "</a>";
    }
    //20160330
    if( glb_locationType== "detail" ){
        if( inStr(stype,"|��ʾ���±���|")>0 ){
            c= c + " >> " + glb_detailTitle;
        }else if( inStr(stype,"|�������±���|")>0 ){
        }else{
            c= c + " >> �鿴����";
        }
    }
    //β��׷������

    return c;
}

</script>
