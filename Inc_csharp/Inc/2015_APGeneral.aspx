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
//ASP��PHPͨ�ú���


//������ر�ǩ ��  ���Ľ�
string aritcleRelatedTags(string relatedTags){
    string c=""; string[] splStr; string s=""; string url="";
    splStr= aspSplit(relatedTags, ",");
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            if( c != "" ){
                c= c + ",";
            }
            url= getColumnUrl(s, "name");
            c= c + "<a href=\"" + url + "\" rel=\"category tag\" class=\"ablue\">" + s + "</a>" + vbCrlf();
        }
    }

    c= "<footer class=\"articlefooter\">" + vbCrlf() + "��ǩ�� " + c + "</footer>" + vbCrlf();
    return c;
}


//����������id�б�
string getRandArticleId(string addSql, int nTopNumb){
    string[] splStr; string s=""; string c=""; int nIndex=0;
    rs = new OleDbCommand("select * from " + db_PREFIX + "articledetail " + addSql, conn).ExecuteReader();

    while( rs.Read()){
        if( c != "" ){ c= c + "," ;}
        c= c + cStr(rs["id"]);
    }
    string getRandArticleId= randomShow(c, ",", 4);
    splStr= aspSplit(c, ",") ; c= "" ; nIndex= 0;
    foreach(var eachs in splStr){
        s=eachs;
        if( c != "" ){ c= c + "," ;}
        c= c + s;
        nIndex= nIndex + 1;
        if( nIndex >= nTopNumb ){ break; }
    }
    return c;
}
//�����վ��Ŀ����SQL
string getWebColumnSortSql(string id){
    string sql="";
    tempRs2 = new OleDbCommand("select * from " + db_PREFIX + "webcolumn where id=" + id, conn).ExecuteReader();

    if( tempRs2.Read() ){

        sql= cStr(cStr(tempRs2["sortsql"]));
    }
    return sql;
}

//��һƪ���� �������sortrank(����)Ҳ���Ը�Ϊid,�����õ�ʱ���Ҫ��id
string upArticle(string parentid, string lableName, string lableValue, string ascOrDesc){
    return handleUpDownArticle("��һƪ��", "uppage", parentid, lableName, lableValue, ascOrDesc);
}
//��һƪ����
string downArticle(string parentid, string lableName, string lableValue, string ascOrDesc){
    return handleUpDownArticle("��һƪ��", "downpage", parentid, lableName, lableValue, ascOrDesc);
}
//��������ҳ
string handleUpDownArticle(string lableTitle, string sType, string parentid, string lableName, string lableValue, string ascOrDesc){
    string c=""; string url=""; string target=""; string targetStr="";

    string sql="";
    if( lableName== "adddatetime" ){
        lableValue= "#" + lableValue + "#";
    }
    //λ�û���
    if( ascOrDesc== "desc" ){
        if( sType== "uppage" ){
            sType= "downpage";
        }else{
            sType= "uppage";
        }
    }
    if( sType== "uppage" ){
        sql= "select * from " + db_PREFIX + "articledetail where parentid=" + parentid + " and " + lableName + "<" + lableValue + " order by " + lableName + " desc";
    }else{
        sql= "select * from " + db_PREFIX + "articledetail where parentid=" + parentid + " and " + lableName + ">" + lableValue + " order by " + lableName + " asc";
    }

    //call echo("sql",sql)
    rsx = new OleDbCommand(sql, conn).ExecuteReader();

    if( rsx.Read() ){
        target= cStr(rsx["target"]);
        if( target != "" ){
            targetStr= " target=\"" + target + "\"";
        }
        if( isMakeHtml== true ){
            url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/detail/detail" + cStr(rsx["id"]));
        }else{
            if( cStr(rsx["customaurl"])== "" ){
                url= handleWebUrl("?act=detail&id=" + cStr(rsx["id"]));
            }else{
                url= handleWebUrl(cStr(rsx["customaurl"]));
            }
        }
        c= "<a href=\"" + url + "\"" + targetStr + ">" + lableTitle + cStr(rsx["title"]) + "</a>";
    }else{
        c= lableTitle + "û��";
    }
    return c;
}
//���RS��ַ ������һҳ ��һҳ
string getRsUrl( string fileName, string customAUrl, string defaultFileName){
    string url="";
    //��Ĭ���ļ�����
    if( fileName== "" ){
        fileName= defaultFileName;
    }
    //��ַ
    if( fileName != "" ){
        fileName= lCase(fileName); //���ļ�����Сд20160315
        url= fileName;
        if( inStr(lCase(url), ".html")== 0 && right(url, 1) != "/" ){
            url= url + ".html";
        }
    }
    if( aspTrim(customAUrl) != "" ){
        url= aspTrim(customAUrl);
    }
    //׷�������Ϊ�������ɾ�̬�ļ�ʱ�����Ի����ҳ���ļ����ƣ�����index.html#about  ���� 20160728
    if( url=="/" ){
        url="/index.html";
    }
    if( inStr(cfg_flags, "|addwebsite|") > 0 ){
        //url = replaceGlobleVariable(url)   '�滻ȫ�ֱ���
        if( inStr(url, "$cfg_websiteurl$")== 0 && inStr(url, "{$GetColumnUrl ")== 0 && inStr(url, "{$GetArticleUrl ")== 0 && inStr(url, "{$GetOnePageUrl ")== 0 ){
            url= urlAddHttpUrl(cfg_webSiteUrl, url);
        }
    }
    return url;
}
//��ô����RS��ַ
string getHandleRsUrl(string fileName, string customAUrl, string defaultFileName){
    string url="";
    url= getRsUrl(fileName, customAUrl, defaultFileName);
    //��ΪURL���Ϊ�Զ��������Ҫ������ȫ�ֱ������������������ֻ���������Ϳ���ʹ������HTML�������������⣬20160308
    url= replaceGlobleVariable(url);
    return url;
}

//��õ�ҳurl 20160114
string getOnePageUrl(string title){
    string url="";
    rsx = new OleDbCommand("select * from " + db_PREFIX + "onepage where title='" + title + "'", conn).ExecuteReader();

    if( rsx.Read() ){
        if( isMakeHtml== true ){
            url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/page/page" + cStr(rsx["id"]));
        }else{
            url= handleWebUrl("?act=onepage&id=" + cStr(rsx["id"]));
            if( cStr(rsx["customaurl"]) != "" ){
                url= cStr(rsx["customaurl"]);
            }
        }
    }

    return url;
}
//�������URL
string getArticleUrl(string title){
    string url="";
    rsx = new OleDbCommand("select * from " + db_PREFIX + "articledetail where title='" + title + "'", conn).ExecuteReader();

    if( rsx.Read() ){
        if( isMakeHtml== true ){
            url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/detail/" + cStr(rsx["id"]));
        }else{
            url= handleWebUrl("?act=article&id=" + cStr(rsx["id"]));
            if( cStr(rsx["customaurl"]) != "" ){
                url= cStr(rsx["customaurl"]);
            }
        }
    }

    return url;
}
//�����ĿURL 20160114 getColumnUrl("��ҳ","type")
string getColumnUrl(string columnNameOrId, string sType){
    string url=""; string addSql="";

    columnNameOrId= replaceGlobleVariable(columnNameOrId); //������ <a href="{$GetColumnUrl columnname='[$glb_columnName$]' $}" >����ͼƬ</a>

    if( sType== "name" ){
        addSql= " where columnname='" + replace(columnNameOrId, "'", "''") + "'"; //��'�Ŵ���Ҫ��Ȼsql��ѯ����20160716
    }else if( sType== "type" ){
        addSql= " where columntype='"+ columnNameOrId +"'"; //��'�Ŵ���Ҫ��Ȼsql��ѯ����20160716
    }else{
        addSql= " where id=" + columnNameOrId + "";
    }
    rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn" + addSql, conn).ExecuteReader();

    if( rsx.Read() ){
        if( isMakeHtml== true ){
            url= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/nav" + cStr(rsx["id"]));
            //call echo(rsx("columnName"),url)
        }else{
            url= handleWebUrl("?act=nav&columnName=" + cStr(rsx["columnname"]));
            if( cStr(rsx["customaurl"]) != "" ){
                url= cStr(rsx["customaurl"]);
            }
        }
    }

    return url;
}

//������±����Ӧ��id
string getArticleId(string title){
    title= replace(title, "'", ""); //ע�⣬���������
    string getArticleId= "-1";
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "ArticleDetail where title='" + title + "'", conn).ExecuteReader();

    if( rsx.Read() ){
        getArticleId= cStr(rsx["id"]);
    }
    return getArticleId;
}

//�����Ŀid
string getColumnId(string columnName){
    //columnName = Replace(columnName, "'", "")           'ע�⣬���������  ��Ϊsql���Ѿ������� 20160716 home ����д��Խ��Խ��߼�Խ��
    string getColumnId= "-1";
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "webcolumn where columnName='" + columnName + "'", conn).ExecuteReader();

    if( rsx.Read() ){
        getColumnId= cStr(rsx["id"]);
    }
    return getColumnId;
}

//�����Ŀ����
string getColumnName(string sID){
    string getColumnName="";
    if( sID != "" ){
        rsx = new OleDbCommand("Select * from " + db_PREFIX + "webcolumn where id=" + sID, conn).ExecuteReader();

        if( rsx.Read() ){
            getColumnName= cStr(rsx["columnname"]);
        }
    }
    return getColumnName;
}

//�����Ŀ����
string getColumnType(string columnID){
    string getColumnType="";
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "webcolumn where id=" + columnID, conn).ExecuteReader();

    if( rsx.Read() ){
        getColumnType= cStr(rsx["columntype"]);
    }
    return getColumnType;
}

//�����Ŀ����
string getColumnBodyContent(string columnID){
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "webcolumn where id=" + columnID, conn).ExecuteReader();

    string getColumnBodyContent="";
    if( rsx.Read() ){
        getColumnBodyContent= cStr(rsx["bodycontent"]);
    }
    return getColumnBodyContent;
}


//��ú�̨�˵�����
string getListMenuId(string title){
    title= replace(title, "'", ""); //ע�⣬���������
    string getListMenuId= "-1";
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "listmenu where title='" + title + "'", conn).ExecuteReader();

    if( rsx.Read() ){
        getListMenuId= cStr(rsx["id"]);
    }
    return getListMenuId;
}

//��ú�̨�˵�ID
string getListMenuName(string sID){
    string getListMenuName="";
    if( sID != "" ){
        rsx = new OleDbCommand("Select * from " + db_PREFIX + "listmenu where id=" + sID, conn).ExecuteReader();

        if( rsx.Read() ){
            getListMenuName= cStr(rsx["title"]);
        }
    }
    return getListMenuName;
}

//��վͳ��2014
string webStat(string folderPath){
    System.DateTime dateTime; string content=""; string[] splStr;
    string thisUrl=""; string goToUrl=""; string caiShu=""; string c=""; string fileName=""; string co=""; string ie=""; string xp="";
    createDirFolder(folderPath);		//����ͳ��ָ���ļ���
    goToUrl= cStr(Request.ServerVariables["HTTP_REFERER"]);
    thisUrl= "http://" + cStr(Request.ServerVariables["HTTP_HOST"]) + cStr(Request.ServerVariables["SCRIPT_NAME"]);
    caiShu= cStr(Request.ServerVariables["QUERY_STRING"]);
    if( caiShu != "" ){
        thisUrl= thisUrl + "?" + caiShu;
    }
    goToUrl= cStr(Request["GoToUrl"]);
    thisUrl= cStr(Request["ThisUrl"]);
    co= cStr(Request.QueryString["co"]);
    dateTime= now();
    content= cStr(Request.ServerVariables["HTTP_USER_AGENT"]);
    content= replace(content, "MSIE", "Internet Explorer");
    content= replace(content, "NT 5.0", "2000");
    content= replace(content, "NT 5.1", "XP");
    content= replace(content, "NT 5.2", "2003");

    splStr= aspSplit(content + ";;;;", ";");
    ie= splStr[1];
    xp= aspTrim(splStr[2]);
    if( right(xp, 1)== ")" ){ xp= mid(xp, 1, len(xp) - 1) ;}
    c= "����" + goToUrl + vbCrlf();
    c= c + "��ǰ��" + thisUrl + vbCrlf();
    c= c + "ʱ�䣺" + dateTime + vbCrlf();
    c= c + "IP:" + getIP() + vbCrlf();
    c= c + "IE:" + getBrType("") + vbCrlf();
    c= c + "Cookies=" + co + vbCrlf();
    c= c + "XP=" + xp + vbCrlf();
    c= c + "Screen=" + cStr(Request["screen"]) + vbCrlf(); //��Ļ�ֱ���
    c= c + "�û���Ϣ=" + cStr(Request.ServerVariables["HTTP_USER_AGENT"]) + vbCrlf(); //�û���Ϣ

    c= c + "-------------------------------------------------" + vbCrlf();
    //c=c & "CaiShu=" & CaiShu & vbcrlf
    fileName= folderPath + format_Time(now(), 2) + ".txt";
    createAddFile(fileName, c);
    c= c + vbCrlf() + fileName;
    c= replace(c, vbCrlf(), "\\n");
    c= replace(c, "\"", "\\\"");
    //Response.Write("eval(""var MyWebStat=\""" & C & "\"""")")

    string[] splxx; int nIP=0; int nPV=0; string ipList=""; string s=""; string ip="";
    //�ж��Ƿ���ʾ���Լ�¼
    if( cStr(Request["stype"])== "display" ){
        content= getFText(fileName);
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
        rw("document.write('����ͳ�� | ����IP[" + nIP + "] | ����PV[" + nPV + "] ')");
    }
    return c;
}

//�жϴ�ֵ�Ƿ����
bool checkFunValue(string action, string funName){
    return IIF(left(action, len(funName))== funName, true, false);
}
//HTML��ǩ�����Զ����(target|title|alt|id|class|style|)    ������
string setHtmlParam(string content, string paramList){
    string[] splStr; string startStr=""; string endStr=""; string c=""; string paramValue=""; string replaceStartStr="";
    endStr= "'";
    splStr= aspSplit(paramList, "|");
    foreach(var eachstartStr in splStr ){
        startStr=eachstartStr;
        startStr= aspTrim(startStr);
        if( startStr != "" ){
            //�滻��ʼ�ַ�   ��Ϊ��ʼ�ַ����Ϳɱ� ��ͬ
            replaceStartStr= startStr;
            if( left(replaceStartStr, 3)== "img" ){
                replaceStartStr= mid(replaceStartStr, 4,-1);
            }else if( left(replaceStartStr, 1)== "a" ){
                replaceStartStr= mid(replaceStartStr, 2,-1);
            }else if( inStr("|ul|li|", "|" + left(replaceStartStr, 2) + "|") > 0 ){
                replaceStartStr= mid(replaceStartStr, 3,-1);
            }
            replaceStartStr= " " + replaceStartStr + "='";

            startStr= " " + startStr + "='";
            if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
                paramValue= strCut(content, startStr, endStr, 2);
                paramValue= handleInModule(paramValue, "end"); //�����ڲ�ģ��
                c= c + replaceStartStr + paramValue + endStr;
            }
        }
    }
    return c;
}
</script>
