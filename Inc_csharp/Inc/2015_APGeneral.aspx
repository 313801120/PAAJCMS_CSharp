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
<script runat="server" language="c#">
//ASP和PHP通用函数


//文章相关标签 组  待改进
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

    c= "<footer class=\"articlefooter\">" + vbCrlf() + "标签： " + c + "</footer>" + vbCrlf();
    return c;
}


//获得随机文章id列表
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
//获得网站栏目排序SQL
string getWebColumnSortSql(string id){
    string sql="";
    tempRs2 = new OleDbCommand("select * from " + db_PREFIX + "webcolumn where id=" + id, conn).ExecuteReader();

    if( tempRs2.Read() ){

        sql= cStr(cStr(tempRs2["sortsql"]));
    }
    return sql;
}

//上一篇文章 这里面的sortrank(排序)也可以改为id,在引用的时候就要用id
string upArticle(string parentid, string lableName, string lableValue, string ascOrDesc){
    return handleUpDownArticle("上一篇：", "uppage", parentid, lableName, lableValue, ascOrDesc);
}
//下一篇文章
string downArticle(string parentid, string lableName, string lableValue, string ascOrDesc){
    return handleUpDownArticle("下一篇：", "downpage", parentid, lableName, lableValue, ascOrDesc);
}
//处理上下页
string handleUpDownArticle(string lableTitle, string sType, string parentid, string lableName, string lableValue, string ascOrDesc){
    string c=""; string url=""; string target=""; string targetStr="";

    string sql="";
    if( lableName== "adddatetime" ){
        lableValue= "#" + lableValue + "#";
    }
    //位置互换
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
        c= lableTitle + "没有";
    }
    return c;
}
//获得RS网址 配置上一页 下一页
string getRsUrl( string fileName, string customAUrl, string defaultFileName){
    string url="";
    //用默认文件名称
    if( fileName== "" ){
        fileName= defaultFileName;
    }
    //网址
    if( fileName != "" ){
        fileName= lCase(fileName); //让文件名称小写20160315
        url= fileName;
        if( inStr(lCase(url), ".html")== 0 && right(url, 1) != "/" ){
            url= url + ".html";
        }
    }
    if( aspTrim(customAUrl) != "" ){
        url= aspTrim(customAUrl);
    }
    //追加这个是为了在生成静态文件时，可以获得首页的文件名称，好让index.html#about  出现 20160728
    if( url=="/" ){
        url="/index.html";
    }
    if( inStr(cfg_flags, "|addwebsite|") > 0 ){
        //url = replaceGlobleVariable(url)   '替换全局变量
        if( inStr(url, "$cfg_websiteurl$")== 0 && inStr(url, "{$GetColumnUrl ")== 0 && inStr(url, "{$GetArticleUrl ")== 0 && inStr(url, "{$GetOnePageUrl ")== 0 ){
            url= urlAddHttpUrl(cfg_webSiteUrl, url);
        }
    }
    return url;
}
//获得处理后RS网址
string getHandleRsUrl(string fileName, string customAUrl, string defaultFileName){
    string url="";
    url= getRsUrl(fileName, customAUrl, defaultFileName);
    //因为URL如果为自定义的则需要处理下全局变量，这样程序运行又会变慢，不就可以使用生成HTML方法解决这个问题，20160308
    url= replaceGlobleVariable(url);
    return url;
}

//获得单页url 20160114
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
//获得文章URL
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
//获得栏目URL 20160114 getColumnUrl("首页","type")
string getColumnUrl(string columnNameOrId, string sType){
    string url=""; string addSql="";

    columnNameOrId= replaceGlobleVariable(columnNameOrId); //处理动作 <a href="{$GetColumnUrl columnname='[$glb_columnName$]' $}" >更多图片</a>

    if( sType== "name" ){
        addSql= " where columnname='" + replace(columnNameOrId, "'", "''") + "'"; //对'号处理，要不然sql查询出错20160716
    }else if( sType== "type" ){
        addSql= " where columntype='"+ columnNameOrId +"'"; //对'号处理，要不然sql查询出错20160716
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

//获得文章标题对应的id
string getArticleId(string title){
    title= replace(title, "'", ""); //注意，这个不能留
    string getArticleId= "-1";
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "ArticleDetail where title='" + title + "'", conn).ExecuteReader();

    if( rsx.Read() ){
        getArticleId= cStr(rsx["id"]);
    }
    return getArticleId;
}

//获得栏目id
string getColumnId(string columnName){
    //columnName = Replace(columnName, "'", "")           '注意，这个不能留  因为sql里已经处理了 20160716 home 程序写得越来越深，逻辑越多
    string getColumnId= "-1";
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "webcolumn where columnName='" + columnName + "'", conn).ExecuteReader();

    if( rsx.Read() ){
        getColumnId= cStr(rsx["id"]);
    }
    return getColumnId;
}

//获得栏目名称
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

//获得栏目类型
string getColumnType(string columnID){
    string getColumnType="";
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "webcolumn where id=" + columnID, conn).ExecuteReader();

    if( rsx.Read() ){
        getColumnType= cStr(rsx["columntype"]);
    }
    return getColumnType;
}

//获得栏目内容
string getColumnBodyContent(string columnID){
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "webcolumn where id=" + columnID, conn).ExecuteReader();

    string getColumnBodyContent="";
    if( rsx.Read() ){
        getColumnBodyContent= cStr(rsx["bodycontent"]);
    }
    return getColumnBodyContent;
}


//获得后台菜单名称
string getListMenuId(string title){
    title= replace(title, "'", ""); //注意，这个不能留
    string getListMenuId= "-1";
    rsx = new OleDbCommand("Select * from " + db_PREFIX + "listmenu where title='" + title + "'", conn).ExecuteReader();

    if( rsx.Read() ){
        getListMenuId= cStr(rsx["id"]);
    }
    return getListMenuId;
}

//获得后台菜单ID
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

//网站统计2014
string webStat(string folderPath){
    System.DateTime dateTime; string content=""; string[] splStr;
    string thisUrl=""; string goToUrl=""; string caiShu=""; string c=""; string fileName=""; string co=""; string ie=""; string xp="";
    createDirFolder(folderPath);		//生成统计指定文件夹
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
    c= "来访" + goToUrl + vbCrlf();
    c= c + "当前：" + thisUrl + vbCrlf();
    c= c + "时间：" + dateTime + vbCrlf();
    c= c + "IP:" + getIP() + vbCrlf();
    c= c + "IE:" + getBrType("") + vbCrlf();
    c= c + "Cookies=" + co + vbCrlf();
    c= c + "XP=" + xp + vbCrlf();
    c= c + "Screen=" + cStr(Request["screen"]) + vbCrlf(); //屏幕分辨率
    c= c + "用户信息=" + cStr(Request.ServerVariables["HTTP_USER_AGENT"]) + vbCrlf(); //用户信息

    c= c + "-------------------------------------------------" + vbCrlf();
    //c=c & "CaiShu=" & CaiShu & vbcrlf
    fileName= folderPath + format_Time(now(), 2) + ".txt";
    createAddFile(fileName, c);
    c= c + vbCrlf() + fileName;
    c= replace(c, vbCrlf(), "\\n");
    c= replace(c, "\"", "\\\"");
    //Response.Write("eval(""var MyWebStat=\""" & C & "\"""")")

    string[] splxx; int nIP=0; int nPV=0; string ipList=""; string s=""; string ip="";
    //判断是否显示回显记录
    if( cStr(Request["stype"])== "display" ){
        content= getFText(fileName);
        splxx= aspSplit(content, vbCrlf() + "-------------------------------------------------" + vbCrlf());
        nIP= 0;
        nPV= 0;
        ipList= "";
        foreach(var eachs in splxx){
            s=eachs;
            if( inStr(s, "当前：") > 0 ){
                s= vbCrlf() + s + vbCrlf();
                ip= ADSql(getStrCut(s, vbCrlf() + "IP:", vbCrlf(), 0));
                nPV= nPV + 1;
                if( inStr(vbCrlf() + ipList + vbCrlf(), vbCrlf() + ip + vbCrlf())== 0 ){
                    ipList= ipList + ip + vbCrlf();
                    nIP= nIP + 1;
                }
            }
        }
        rw("document.write('网长统计 | 今日IP[" + nIP + "] | 今日PV[" + nPV + "] ')");
    }
    return c;
}

//判断传值是否相等
bool checkFunValue(string action, string funName){
    return IIF(left(action, len(funName))== funName, true, false);
}
//HTML标签参数自动添加(target|title|alt|id|class|style|)    辅助类
string setHtmlParam(string content, string paramList){
    string[] splStr; string startStr=""; string endStr=""; string c=""; string paramValue=""; string replaceStartStr="";
    endStr= "'";
    splStr= aspSplit(paramList, "|");
    foreach(var eachstartStr in splStr ){
        startStr=eachstartStr;
        startStr= aspTrim(startStr);
        if( startStr != "" ){
            //替换开始字符   因为开始字符类型可变 不同
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
                paramValue= handleInModule(paramValue, "end"); //处理内部模块
                c= c + replaceStartStr + paramValue + endStr;
            }
        }
    }
    return c;
}
</script>
