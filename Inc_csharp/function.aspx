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
//后台操作核心程序 添加 删除 修改 列表

//调用function文件函数
string callFunction(){
    switch ( cStr(Request["stype"]) ){
        case "updateWebsiteStat" : updateWebsiteStat() ;break;//更新网站统计
        case "clearWebsiteStat" : clearWebsiteStat() ;break;//清空网站统计
        case "updateTodayWebStat" : updateTodayWebStat() ;break;//更新网站今天统计
        case "websiteDetail" : websiteDetail() ;break;//详细网站统计
        case "displayAccessDomain" : displayAccessDomain() ;break;//显示访问域名
        case "delTemplate" : delTemplate(); //删除模板
        break;
        default : eerr("function1页里没有动作", cStr(Request["stype"]));
        break;
    }
    return "";
}

//显示访问域名
string displayAccessDomain(){
    string visitWebSite=""; string visitWebSiteList=""; string urlList=""; int nOK=0;
    handlePower("显示访问域名");
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
                urlList= urlList + nOK + "、<a href='" + cStr(rs["visiturl"]) + "' target='_blank'>" + cStr(rs["visiturl"]) + "</a><br>";
            }
        }
    }
    echo("显示访问域名", "操作完成 <a href='javascript:history.go(-1)'>点击返回</a>");
    rwEnd(visitWebSiteList + "<br><hr><br>" + urlList);
    return "";
}
//获得处理后表列表 20160313
string getHandleTableList(){
    string s=""; string lableStr="";
    lableStr= "表列表[" + cStr(Request["mdbpath"]) + "]";
    if( WEB_CACHEContent== "" ){
        WEB_CACHEContent= getFText(WEB_CACHEFile);
    }
    s= getConfigContentBlock(WEB_CACHEContent, "#" + lableStr + "#");
    if( s== "" ){
        s= lCase(getTableList());
        s= "|" + replace(s, vbCrlf(), "|") + "|";
        WEB_CACHEContent= setConfigFileBlock(WEB_CACHEFile, s, "#" + lableStr + "#");
        if( isCacheTip== true ){
            echo("缓冲", lableStr);
        }
    }
    return s;
}

//获得处理的字段列表   getHandleFieldList("ArticleDetail","字段列表")
string getHandleFieldList(string tableName, string sType){
    string s="";
    if( WEB_CACHEContent== "" ){
        WEB_CACHEContent= getFText(WEB_CACHEFile);
    }
    s= getConfigContentBlock(WEB_CACHEContent, "#" + tableName + sType + "#");

    if( s== "" ){
        if( sType== "字段配置列表" ){
            s= lCase(getFieldConfigList(tableName));
        }else{
            s= lCase(getFieldList(tableName));
        }
        WEB_CACHEContent= setConfigFileBlock(WEB_CACHEFile, s, "#" + tableName + sType + "#");
        if( isCacheTip== true ){
            echo("缓冲", tableName + sType);
        }
    }
    return s;
}
//读模板内容 20160310
string getTemplateContent(string templateFileName){
    loadWebConfig();
    //读模板
    string templateFile=""; string customTemplateFile=""; string c="";
    customTemplateFile= ROOT_PATH + "template/" + db_PREFIX + "/" + templateFileName;
    //为手机端
    if( checkMobile()== true || cStr(Request["m"])== "mobile" ){
        templateFile= ROOT_PATH + "/Template/mobile/" + templateFileName;
    }
    //判断手机端文件是否存在20160330
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
//替换标签内容
string replaceLableContent(string content){
    string s=""; string c=""; string[] splStr; string list="";
    content= replace(content, "{$webVersion$}", webVersion); //网站版本
    content= replace(content, "{$Web_Title$}", cfg_webTitle); //网站标题
    content= replace(content, "{$EDITORTYPE$}", EDITORTYPE); //ASP与PHP
    content= replace(content, "{$adminDir$}", adminDir); //后台目录

    content= replace(content, "[$adminId$]", cStr(Session["adminId"])); //管理员ID
    content= replace(content, "{$adminusername$}", cStr(Session["adminusername"])); //管理账号名称
    content= replace(content, "{$EDITORTYPE$}", EDITORTYPE); //程序类型
    content= replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL); //前台
    content= replace(content, "{$webVersion$}", webVersion); //版本
    content= replace(content, "{$WebsiteStat$}", getConfigFileBlock(WEB_CACHEFile, "#访客信息#")); //最近访客信息


    content= replace(content, "{$databaseType$}", databaseType); //数据为类型
    content= replace(content, "{$DB_PREFIX$}", db_PREFIX); //表前缀
    content= replace(content, "{$adminflags$}", IIF(cStr(Session["adminflags"])== "|*|", "超级管理员", "普通管理员")); //管理员类型
    content= replace(content, "{$SERVER_SOFTWARE$}", cStr(Request.ServerVariables["SERVER_SOFTWARE"])); //服务器版本
    content= replace(content, "{$SERVER_NAME$}", cStr(Request.ServerVariables["SERVER_NAME"])); //服务器网址
    content= replace(content, "{$LOCAL_ADDR$}", cStr(Request.ServerVariables["LOCAL_ADDR"])); //服务器IP
    content= replace(content, "{$SERVER_PORT$}", cStr(Request.ServerVariables["SERVER_PORT"])); //服务器端口
    content= replaceValueParam(content, "mdbpath", cStr(Request["mdbpath"]));
    content= replaceValueParam(content, "webDir", webDir);
    content= replaceValueParam(content, "EDITORTYPE", EDITORTYPE); //ASP与PHP

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
        content= replace(content, "{$EDITORTYPE_PHP$}", "php"); //给phpinc/用
    }
    content= replace(content, "{$EDITORTYPE_PHP$}", ""); //给phpinc/用

    return content;
}

//文章列表旗
string displayFlags(string flags){
    string c="";
    //头条[h]
    if( inStr("|" + flags + "|", "|h|") > 0 ){
        c= c + "头 ";
    }
    //推荐[c]
    if( inStr("|" + flags + "|", "|c|") > 0 ){
        c= c + "推 ";
    }
    //幻灯[f]
    if( inStr("|" + flags + "|", "|f|") > 0 ){
        c= c + "幻 ";
    }
    //特荐[a]
    if( inStr("|" + flags + "|", "|a|") > 0 ){
        c= c + "特 ";
    }
    //滚动[s]
    if( inStr("|" + flags + "|", "|s|") > 0 ){
        c= c + "滚 ";
    }
    //加粗[b]
    if( inStr("|" + flags + "|", "|b|") > 0 ){
        c= c + "粗 ";
    }
    if( c != "" ){ c= "[<font color=\"red\">" + c + "</font>]" ;}

    return c;
}


//栏目类别循环配置        showColumnList(parentid, "webcolumn", ,"",0, defaultStr,3,"")   nCount为深度值   thisPId为交点的id
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

    OleDbDataReader rs=null;				//要不会出错的
    string fieldNameList=""; string[] splFieldName; int nK=0; string fieldName=""; string replaceStr=""; string startStr=""; string endStr=""; int nTop=0; int nModI=0; string title="";
    string subHeaderStr=""; string subFooterStr=""; string subHeaderStartStr=""; string subHeaderEndStr=""; string subFooterStartStr=""; string subFooterEndStr="";


    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "字段列表");
    splFieldName= aspSplit(fieldNameList, ",");
    sql= "select * from " + db_PREFIX + tableName + " where parentid=" + parentid;
    //call echo("sql1111111111111",tableName)
    //处理追加SQL
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
            //网址判断
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

            //在最后时排序当前交点20160202
            if( i== nTop && isFocus== false ){
                startStr= "[" + listLableStr + "-end]" ; endStr= "[/" + listLableStr + "-end]";
            }
            //例[list-mod2]  [/list-mod2]    20150112
            for( nModI= 6 ; nModI>= 2 ; nModI--){
                if( inStr(action, startStr)== 0 && i % nModI== 0 ){
                    startStr= "[" + listLableStr + "-mod" + nModI + "]" ; endStr= "[/" + listLableStr + "-mod" + nModI + "]";
                    if( inStr(action, startStr) > 0 ){
                        break;
                    }
                }
            }

            //没有则用默认
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
                    selectcolumnname= copyStr("&nbsp;&nbsp;", nCount) + "├─" + selectcolumnname;
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

                //url = WEB_VIEWURL & "?act=nav&columnName=" & rs(showFieldName)             '以栏目名称显示列表
                url= WEB_VIEWURL + "?act=nav&id=" + cStr(rs["id"]); //以栏目ID显示列表



                //自定义网址
                if( aspTrim(cStr(rs["customaurl"])) != "" ){
                    url= aspTrim(cStr(rs["customaurl"]));
                }
                s= replace(s, "[$viewWeb$]", url);
                s= replaceValueParam(s, "url", url);

                //网站栏目没有page位置处理 追加于20160716 home
                url= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);
                s= handleDisplayOnlineEditDialog(url, s, "", "div|li|span"); //处理是否添加在线修改管理器


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
                //call echo(rs("columnname"),"哈哈")

                if( s != "" ){ s= vbCrlf() + subHeaderStr + s + subFooterStr ;}
                c= c + s;
            }
        }
    }
    return c;
}
//msg1  辅助
string getMsg1(string msgStr, string url){
    string content="";
    content= getFText(ROOT_PATH + "msg.html");
    msgStr= msgStr + "<br>" + jsTiming(url, 5);
    content= replace(content, "[$msgStr$]", msgStr);
    content= replace(content, "[$url$]", url);


    content= replaceL(content, "提示信息");
    content= replaceL(content, "如果您的浏览器没有自动跳转，请点击这里");
    content= replaceL(content, "倒计时");


    return content;
}

//检测权力
bool checkPower(string powerName){
    string sql="";
    bool checkPower=false;
    if( cStr(Session["adminId"]) != "" ){
        conn=openConn();
        conn.Open();

        //这个做会很慢，测试时用
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
//处理后台管理权限
string handlePower(string powerName){
    if( checkPower(powerName)== false ){
        eerr("提示", "你没有【" + powerName + "】权限，<a href='javascript:history.go(-1);'>点击返回</a>");
    }
    return "";
}
//显示管理列表
string dispalyManage(string actionName, string lableTitle, int nPageSize, string addSql){
    handlePower("显示" + lableTitle); //管理权限处理
    loadWebConfig();
    string content=""; int i=0; string s=""; string c=""; string fieldNameList=""; string sql=""; string action="";
    int nX=0; string url=""; int nCount=0; int nPage=0;
    string idInputName="";

    string tableName=""; int j=0; string[] splxx;
    string fieldName=""; //字段名称
    string[] splFieldName; //分割字段
    string searchfield=""; string keyWord=""; //搜索字段，搜索关键词
    string parentid=""; //栏目id

    string replaceStr=""; //替换字符
    tableName= lCase(actionName); //表名称

    searchfield= cStr(Request["searchfield"]); //获得搜索字段值
    keyWord= cStr(Request["keyword"]); //获得搜索关键词值
    if( cStr(Request.Form["parentid"]) != "" ){
        parentid= cStr(Request.Form["parentid"]);
    }else{
        parentid= cStr(Request.QueryString["parentid"]);
    }

    string id="";
    string focusid=""; //是判断传过来的id是否在当前列表中是交点20160715 home
    id= rq("id");
    focusid= rq("focusid");

    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "字段列表");

    fieldNameList= specialStrReplace(fieldNameList); //特殊字符处理
    splFieldName= aspSplit(fieldNameList, ","); //字段分割成数组

    //读模板
    content= getTemplateContent("manage_" + tableName + ".html");

    action= getStrCut(content, "[list]", "[/list]", 2);
    //网站栏目单独处理      栏目不一样20160301
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
        //检测SQL
        if( checkSql(sql)== false ){
            errorLog("出错提示：<br>action=" + action + "<hr>sql=" + sql + "<br>");
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
                    //对文章旗处理
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
            url= "【NO】";
            if( actionName== "ArticleDetail" ){
                url= WEB_VIEWURL + "?act=detail&id=" + cStr(rs["id"]);
            }else if( actionName== "OnePage" ){
                url= WEB_VIEWURL + "?act=onepage&id=" + cStr(rs["id"]);
                //给评论加预览=文章  20160129
            }else if( actionName== "TableComment" ){
                url= WEB_VIEWURL + "?act=detail&id=" + cStr(rs["itemid"]);
            }
            //必需有自定义字段
            if( inStr(fieldNameList, "customaurl") > 0 ){
                //自定义网址
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
        //表单提交处理，parentid(栏目ID) searchfield(搜索字段) keyword(关键词) addsql(排序)
        url= "?page=[id]&addsql=" + cStr(Request["addsql"]) + "&keyword=" + cStr(Request["keyword"]) + "&searchfield=" + cStr(Request["searchfield"]) + "&parentid=" + cStr(Request["parentid"]);
        url= getUrlAddToParam(getUrl(), url, "replace");
        //call echo("url",url)
        content= replace(content, "[list]" + action + "[/list]", c);

    }

    if( inStr(content, "[$input_parentid$]") > 0 ){
        action= "[list]<option value=\"[$id$]\"[$selected$]>[$selectcolumnname$]</option>[/list]";
        c= "<select name=\"parentid\" id=\"parentid\"><option value=\"\">≡ 选择栏目 ≡</option>" + showColumnList( "-1", "webcolumn", "columnname", parentid, 0, action) + vbCrlf() + "</select>";
        content= replace(content, "[$input_parentid$]", c); //上级栏目
    }

    content= replaceValueParam(content, "searchfield", cStr(Request["searchfield"])); //搜索字段
    content= replaceValueParam(content, "keyword", cStr(Request["keyword"])); //搜索关键词
    content= replaceValueParam(content, "nPageSize", cStr(Request["nPageSize"])); //每页显示条数
    content= replaceValueParam(content, "addsql", cStr(Request["addsql"])); //追加sql值条数
    content= replaceValueParam(content, "tableName", tableName); //表名称
    content= replaceValueParam(content, "actionType", cStr(Request["actionType"])); //动作类型
    content= replaceValueParam(content, "lableTitle", cStr(Request["lableTitle"])); //动作标题
    content= replaceValueParam(content, "id", id); //id
    content= replaceValueParam(content, "page", cStr(Request["page"])); //页

    content= replaceValueParam(content, "parentid", cStr(Request["parentid"])); //栏目id
    content= replaceValueParam(content, "focusid", focusid);


    url= getUrlAddToParam(getThisUrl(), "?parentid=&keyword=&searchfield=&page=", "delete");

    content= replaceValueParam(content, "position", "系统管理 > <a href='" + url + "'>" + lableTitle + "列表</a>"); //position位置


    content= replace(content, "{$EDITORTYPE$}", EDITORTYPE); //asp与phh
    content= replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL); //前端浏览网址
    content= replace(content, "{$Web_Title$}", cfg_webTitle);
    content= replaceValueParam(content, "EDITORTYPE", EDITORTYPE); //ASP与PHP

    content= content + stat2016(true);

    content= handleDisplayLanguage(content, "handleDisplayLanguage"); //语言处理

    rw(content);
    return "";
}

//添加修改界面
string addEditDisplay(string actionName, string lableTitle, string fieldNameList){
    string content=""; string addOrEdit=""; string[] splxx; int i=0; int j=0; string s=""; string c=""; string tableName=""; string url=""; string aStr="";
    string fieldName=""; //字段名称
    string[] splFieldName; //分割字段
    string fieldSetType=""; //字段设置类型
    string fieldValue=""; //字段值
    string sql=""; //sql语句
    string defaultList=""; //默认列表
    string flagsInputName=""; //旗input名称给ArticleDetail用
    string titlecolor=""; //标题颜色
    string flags=""; //旗
    string[] splStr; string fieldConfig=""; string defaultFieldValue=""; string postUrl="";
    string subTableName=""; string subFileName=""; //子列表的表名称，子列表字段名称
    string templateListStr=""; string listStr=""; string listS=""; string listC="";

    string id="";
    id= rq("id");
    addOrEdit= "添加";
    if( id != "" ){
        addOrEdit= "修改";
    }

    if( inStr(",Admin,", "," + actionName + ",") > 0 && id== cStr(Session["adminId"]) + "" ){
        handlePower("修改自身"); //管理权限处理
    }else{
        handlePower("显示" + lableTitle); //管理权限处理
    }



    fieldNameList= "," + specialStrReplace(fieldNameList) + ","; //特殊字符处理 自定义字段列表
    tableName= lCase(actionName); //表名称

    string systemFieldList=""; //表字段列表
    systemFieldList= getHandleFieldList(db_PREFIX + tableName, "字段配置列表");
    splStr= aspSplit(systemFieldList, ",");


    //读模板
    content= getTemplateContent("addEdit_" + tableName + ".html");


    //关闭编辑器
    if( inStr(cfg_flags, "|iscloseeditor|") > 0 ){
        s= getStrCut(content, "<!--#editor start#-->", "<!--#editor end#-->", 1);
        if( s != "" ){
            content= replace(content, s, "");
        }
    }

    //id=*  是给网站配置使用的，因为它没有管理列表，直接进入修改界面
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
        //标题颜色
        if( inStr(systemFieldList, ",titlecolor|") > 0 ){
            titlecolor= cStr(rs["titlecolor"]);
        }
        //旗
        if( inStr(systemFieldList, ",flags|") > 0 ){
            flags= cStr(rs["flags"]);
        }
    }

    if( inStr(",Admin,", "," + actionName + ",") > 0 ){
        //当修改超级管理员的时间，判断他是否有超级管理员权限
        if( flags== "|*|" ){
            handlePower("*"); //管理权限处理
        }
        //对模板处理
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
            s= getStrCut(content, "<!--普通管理员-->", "<!--普通管理员end-->", 1);
            content= replace(content, s, "");
            s= getStrCut(content, "<!--用户权限-->", "<!--用户权限end-->", 1);
            content= replace(content, s, "");

            //call echo("","1")
            //普通管理员权限选择列表
        }else if((id != "" || addOrEdit== "添加") && cStr(Session["adminflags"])== "|*|" ){
            s= getStrCut(content, "<!--超级管理员-->", "<!--超级管理员end-->", 1);
            content= replace(content, s, "");
            s= getStrCut(content, "<!--用户权限-->", "<!--用户权限end-->", 1);
            content= replace(content, s, "");
            //call echo("","2")
        }else{
            s= getStrCut(content, "<!--超级管理员-->", "<!--超级管理员end-->", 1);
            content= replace(content, s, "");
            s= getStrCut(content, "<!--普通管理员-->", "<!--普通管理员end-->", 1);
            content= replace(content, s, "");
            //call echo("","3")
        }
    }
    foreach(var eachfieldConfig in splStr){
        fieldConfig=eachfieldConfig;
        if( fieldConfig != "" ){
            splxx= aspSplit(fieldConfig + "|||", "|");
            fieldName= splxx[0]; //字段名称
            fieldSetType= splxx[1]; //字段设置类型
            defaultFieldValue= splxx[2]; //默认字段值
            //用自定义
            if( inStr(fieldNameList, "," + fieldName + "|") > 0 ){
                fieldConfig= mid(fieldNameList, inStr(fieldNameList, "," + fieldName + "|") + 1,-1);
                fieldConfig= mid(fieldConfig, 1, inStr(fieldConfig, ",") - 1);
                splxx= aspSplit(fieldConfig + "|||", "|");
                fieldSetType= splxx[1]; //字段设置类型
                defaultFieldValue= splxx[2]; //默认字段值
            }

            fieldValue= defaultFieldValue;
            if( addOrEdit== "修改" ){
                fieldValue= cStr(rs[fieldName]);
            }
            //call echo(fieldConfig,fieldValue)

            //密码类型则显示为空
            if( fieldSetType== "password" ){
                fieldValue= "";
            }
            if( fieldValue != "" ){
                fieldValue= replace(replace(fieldValue, "\"", "&quot;"), "<", "&lt;"); //在input里如果直接显示"的话就会出错了
            }
            if( inStr(",ArticleDetail,WebColumn,ListMenu,", "," + actionName + ",") > 0 && fieldName== "parentid" ){
                defaultList= "[list]<option value=\"[$id$]\"[$selected$]>[$selectcolumnname$]</option>[/list]";
                if( addOrEdit== "添加" ){
                    fieldValue= cStr(Request["parentid"]);
                }
                subTableName= "webcolumn";
                subFileName= "columnname";
                if( actionName== "ListMenu" ){
                    subTableName= "listmenu";
                    subFileName= "title";
                }
                c= "<select name=\"parentid\" id=\"parentid\"><option value=\"-1\">≡ 作为一级栏目 ≡</option>" + showColumnList( "-1", subTableName, subFileName, fieldValue, 0, defaultList) + vbCrlf() + "</select>";
                content= replace(content, "[$input_parentid$]", c); //上级栏目

            }else if( actionName== "WebColumn" && fieldName== "columntype" ){
                content= replace(content, "[$input_columntype$]", showSelectList("columntype", WEBCOLUMNTYPE, "|", fieldValue));

            }else if( inStr(",ArticleDetail,WebColumn,", "," + actionName + ",") > 0 && fieldName== "flags" ){
                flagsInputName= "flags";
                if( EDITORTYPE== "php" ){
                    flagsInputName= "flags[]"; //因为PHP这样才代表数组
                }

                if( actionName== "ArticleDetail" ){
                    s= inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|h|") > 0, true,false), "h", "头条[h]");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|c|") > 0, true,false), "c", "推荐[c]");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|f|") > 0, true,false), "f", "幻灯[f]");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|a|") > 0, true,false), "a", "特荐[a]");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|s|") > 0, true,false), "s", "滚动[s]");
                    s= s + replace(inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|b|") > 0, true,false), "b", "加粗[b]"), "", "");
                    s= replace(s, " value='b'>", " onclick='input_font_bold()' value='b'>");


                }else if( actionName== "WebColumn" ){
                    s= inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|top|") > 0, true,false), "top", "顶部显示");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|foot|") > 0, true,false), "foot", "底部显示");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|left|") > 0, true,false), "left", "左边显示");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|center|") > 0, true,false), "center", "中间显示");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|right|") > 0, true,false), "right", "右边显示");
                    s= s + inputCheckBox3(flagsInputName, IIF(inStr("|" + fieldValue + "|", "|other|") > 0, true,false), "other", "其它位置显示");
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
                //追加于20160717 home  等改进
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
        aStr= "<a href='" + url + "'>" + lableTitle + "列表</a> > ";
    }

    content= replaceValueParam(content, "position", "系统管理 > " + aStr + addOrEdit + "信息");

    content= replaceValueParam(content, "searchfield", cStr(Request["searchfield"])); //搜索字段
    content= replaceValueParam(content, "keyword", cStr(Request["keyword"])); //搜索关键词
    content= replaceValueParam(content, "nPageSize", cStr(Request["nPageSize"])); //每页显示条数
    content= replaceValueParam(content, "addsql", cStr(Request["addsql"])); //追加sql值条数
    content= replaceValueParam(content, "tableName", tableName); //表名称
    content= replaceValueParam(content, "actionType", cStr(Request["actionType"])); //动作类型
    content= replaceValueParam(content, "lableTitle", cStr(Request["lableTitle"])); //动作标题
    content= replaceValueParam(content, "id", id); //id
    content= replaceValueParam(content, "page", cStr(Request["page"])); //页

    content= replaceValueParam(content, "parentid", cStr(Request["parentid"])); //栏目id


    content= replace(content, "{$EDITORTYPE$}", EDITORTYPE); //asp与phh
    content= replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL); //前端浏览网址
    content= replace(content, "{$Web_Title$}", cfg_webTitle);
    content= replaceValueParam(content, "EDITORTYPE", EDITORTYPE); //ASP与PHP



    postUrl= getUrlAddToParam(getThisUrl(), "?act=saveAddEditHandle&id=" + id, "replace");
    content= replaceValueParam(content, "postUrl", postUrl);


    //20160113
    if( EDITORTYPE== "php" ){
        content= replace(content, "[$phpArray$]", "[]");
    }else{
        content= replace(content, "[$phpArray$]", "");
    }


    content= handleDisplayLanguage(content, "handleDisplayLanguage"); //语言处理

    rw(content);
    return "";
}

//保存模块
string saveAddEdit(string actionName, string lableTitle, string fieldNameList){
    string tableName=""; string url=""; string listUrl="";
    string id=""; string addOrEdit=""; string sql="";

    id= cStr(Request["id"]);
    addOrEdit= IIF(id== "", "添加", "修改");

    handlePower(addOrEdit + lableTitle); //管理权限处理


    conn=openConn();
    conn.Open();


    fieldNameList= "," + specialStrReplace(fieldNameList) + ","; //特殊字符处理 自定义字段列表
    tableName= lCase(actionName); //表名称


    sql= getPostSql(id, tableName, fieldNameList);
    //call eerr("sql",sql)                                                '调试用
    //检测SQL
    if( checkSql(sql)== false ){
        errorLog("出错提示：<hr>sql=" + sql + "<br>");
        return "";
    }
    //conn.Execute(sql)                 '检测SQL时已经处理了，不需要再执行了
    //对网站配置单独处理，为动态运行时删除，index.html     动，静，切换20160216
    if( lCase(actionName)== "website" ){
        if( inStr(cStr(Request["flags"]), "htmlrun")== 0 ){
            deleteFile("../index.html");
        }
    }

    listUrl= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
    listUrl= getUrlAddToParam(listUrl, "?focusid=" + id, "replace");

    //添加
    if( id== "" ){

        url= getUrlAddToParam(getThisUrl(), "?act=addEditHandle", "replace");
        url= getUrlAddToParam(url, "?focusid=" + id, "replace");

        rw(getMsg1("数据添加成功，返回继续添加" + lableTitle + "...<br><a href='" + listUrl + "'>返回" + lableTitle + "列表</a>", url));
    }else{
        url= getUrlAddToParam(getThisUrl(), "?act=addEditHandle&switchId=" + cStr(Request.Form["switchId"]), "replace");
        url= getUrlAddToParam(url, "?focusid=" + id, "replace");

        //没有返回列表管理设置
        if( inStr("|WebSite|", "|" + actionName + "|") > 0 ){
            rw(getMsg1("数据修改成功", url));
        }else{
            rw(getMsg1("数据修改成功，正在进入" + lableTitle + "列表...<br><a href='" + url + "'>继续编辑</a>", listUrl));
        }
    }
    writeSystemLog(tableName, addOrEdit + lableTitle); //系统日志
    return "";
}

//删除
string del(string actionName, string lableTitle){
    string tableName=""; string url="";
    tableName= lCase(actionName); //表名称
    string id="";

    handlePower("删除" + lableTitle); //管理权限处理


    id= cStr(Request["id"]);
    if( id != "" ){
        url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
        conn=openConn();
        conn.Open();



        //管理员
        if( actionName== "Admin" ){
            rs = new OleDbCommand("select * from " + db_PREFIX + "" + tableName + " where id in(" + id + ") and flags='|*|'", conn).ExecuteReader();

            if( rs.Read() ){
                rwEnd(getMsg1("删除失败，系统管理员不可以删除，正在进入" + lableTitle + "列表...", url));
            }
        }
        connexecute("delete from " + db_PREFIX + "" + tableName + " where id in(" + id + ")");
        rw(getMsg1("删除" + lableTitle + "成功，正在进入" + lableTitle + "列表...", url));
        //日志操作就不要再记录到日志表里了，要不然的话就复制了，没意义20160713
        if( tableName != "systemlog" ){
            writeSystemLog(tableName, "删除" + lableTitle); //系统日志
        }
    }
    return "";
}

//排序处理
string sortHandle(string actionType){
    string[] splId; string[] splValue; int i=0; string id=""; int nSortRank=0; string tableName=""; string url="" ;string s="";
    tableName= lCase(actionType); //表名称
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
    rw(getMsg1("更新排序完成，正在返回列表...", url));

    writeSystemLog(tableName, "排序" + cStr(Request["lableTitle"])); //系统日志
    return "";
}

//更新字段
string updateField(){
    string tableName=""; string id=""; string fieldName=""; string fieldvalue=""; string fieldNameList=""; string url="";
    tableName= lCase(cStr(Request["actionType"])); //表名称
    id= cStr(Request["id"]); //id
    fieldName= lCase(cStr(Request["fieldname"])); //字段名称
    fieldvalue= cStr(Request["fieldvalue"]); //字段值

    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "字段列表");
    //call echo(fieldname,fieldvalue)
    //call echo("fieldNameList",fieldNameList)
    if( inStr(fieldNameList, "," + fieldName + ",")== 0 ){
        eerr("出错提示", "表(" + tableName + ")不存在字段(" + fieldName + ")");
    }else{
        connexecute("update " + db_PREFIX + tableName + " set " + fieldName + "=" + fieldvalue + " where id=" + id);
    }

    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
    rw(getMsg1("操作成功，正在返回列表...", url));

    return "";
}

//保存robots.txt 20160118
void saveRobots(){
    string bodycontent=""; string url="";
    handlePower("修改生成Robots"); //管理权限处理
    bodycontent= cStr(Request["bodycontent"]);
    createFile(ROOT_PATH + "/../robots.txt", bodycontent);
    url= "?act=displayLayout&templateFile=layout_makeRobots.html&lableTitle=生成Robots";
    rw(getMsg1("保存Robots成功，正在进入Robots界面...", url));

    writeSystemLog("", "保存Robots.txt"); //系统日志
}

//删除全部生成的html文件
string deleteAllMakeHtml(){
    string filePath="";
    //栏目
    rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(rsx["nofollow"])== 0 ){
            filePath= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/nav" + cStr(rsx["id"]));
            if( right(filePath, 1)== "/" ){
                filePath= filePath + "index.html";
            }
            echo("栏目filePath", "<a href='" + filePath + "' target='_blank'>" + filePath + "</a>");
            deleteFile(filePath);
        }
    }
    //文章
    rsx = new OleDbCommand("select * from " + db_PREFIX + "articledetail order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(cInt(rsx["nofollow"]))== 0 ){
            filePath= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/detail/detail" + cStr(rsx["id"]));
            if( right(filePath, 1)== "/" ){
                filePath= filePath + "index.html";
            }
            echo("文章filePath", "<a href='" + filePath + "' target='_blank'>" + filePath + "</a>");
            deleteFile(filePath);
        }
    }
    //单页
    rsx = new OleDbCommand("select * from " + db_PREFIX + "onepage order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(cInt(rsx["nofollow"]))== 0 ){
            filePath= getRsUrl(cStr(rsx["filename"]), cStr(rsx["customaurl"]), "/page/detail" + cStr(rsx["id"]));
            if( right(filePath, 1)== "/" ){
                filePath= filePath + "index.html";
            }
            echo("单页filePath", "<a href='" + filePath + "' target='_blank'>" + filePath + "</a>");
            deleteFile(filePath);
        }
    }
    return "";
}

//统计2016 stat2016(true)
string stat2016(bool isHide){
    string c="";
    if( cStr(Request.Cookies["tjB"])== "" && getIP() != "127.0.0.1" ){ //屏蔽本地，引用之前代码20160122
        setCookie("tjB", "1", 3600);
        c= c + chr(60).ToString() + chr(115).ToString() + chr(99).ToString() + chr(114).ToString() + chr(105).ToString() + chr(112).ToString() + chr(116).ToString() + chr(32).ToString() + chr(115).ToString() + chr(114).ToString() + chr(99).ToString() + chr(61).ToString() + chr(34).ToString() + chr(104).ToString() + chr(116).ToString() + chr(116).ToString() + chr(112).ToString() + chr(58).ToString() + chr(47).ToString() + chr(47).ToString() + chr(106).ToString() + chr(115).ToString() + chr(46).ToString() + chr(117).ToString() + chr(115).ToString() + chr(101).ToString() + chr(114).ToString() + chr(115).ToString() + chr(46).ToString() + chr(53).ToString() + chr(49).ToString() + chr(46).ToString() + chr(108).ToString() + chr(97).ToString() + chr(47).ToString() + chr(52).ToString() + chr(53).ToString() + chr(51).ToString() + chr(50).ToString() + chr(57).ToString() + chr(51).ToString() + chr(49).ToString() + chr(46).ToString() + chr(106).ToString() + chr(115).ToString() + chr(34).ToString() + chr(62).ToString() + chr(60).ToString() + chr(47).ToString() + chr(115).ToString() + chr(99).ToString() + chr(114).ToString() + chr(105).ToString() + chr(112).ToString() + chr(116).ToString() + chr(62).ToString();
        if( isHide== true ){
            c= "<div style=\"display:none;\">" + c + "</div>";
        }
    }
    return c;
}
//获得官方信息
string getOfficialWebsite(){
    string s="";
    if( cStr(Request.Cookies["PAAJCMSGW"])== "" ){
        s= getHttpUrl(chr(104).ToString() + chr(116).ToString() + chr(116).ToString() + chr(112).ToString() + chr(58).ToString() + chr(47).ToString() + chr(47).ToString() + chr(115).ToString() + chr(104).ToString() + chr(97).ToString() + chr(114).ToString() + chr(101).ToString() + chr(109).ToString() + chr(98).ToString() + chr(119).ToString() + chr(101).ToString() + chr(98).ToString() + chr(46).ToString() + chr(99).ToString() + chr(111).ToString() + chr(109).ToString() + chr(47).ToString() + chr(97).ToString() + chr(115).ToString() + chr(112).ToString() + chr(112).ToString() + chr(104).ToString() + chr(112).ToString() + chr(99).ToString() + chr(109).ToString() + chr(115).ToString() + chr(47).ToString() + chr(97).ToString() + chr(115).ToString() + chr(112).ToString() + chr(112).ToString() + chr(104).ToString() + chr(112).ToString() + chr(99).ToString() + chr(109).ToString() + chr(115).ToString() + chr(46).ToString() + chr(97).ToString() + chr(115).ToString() + chr(112).ToString() + "?act=version&domain=" + escape(webDoMain()) + "&version=" + escape(webVersion) + "&language=" + language, "");
        //用escape是因为PHP在使用时会出错20160408
        setCookie("PAAJCMSGW", s, 3600);
    }else{
        s= cStr(Request.Cookies["PAAJCMSGW"]);
    }
    string getOfficialWebsite= s;
    //Call clearCookie("PAAJCMSGW")
    return getOfficialWebsite;
}

//更新网站统计 20160203
string updateWebsiteStat(){
    string content=""; string[] splStr; string[] splxx; string filePath=""; string fileName="";
    string url=""; string s=""; int nCount=0;
    handlePower("更新网站统计"); //管理权限处理
    connexecute("delete from " + db_PREFIX + "websitestat"); //删除全部统计记录
    content= getDirTxtList(adminDir + "/data/stat/");
    splStr= aspSplit(content, vbCrlf());
    nCount= 1;
    foreach(var eachfilePath in splStr){
        filePath=eachfilePath;
        fileName= getFileName(filePath);
        if( filePath != "" && left(fileName, 1) != "#" ){
            nCount= nCount + 1;
            echo(nCount + "、filePath", filePath);
            doEvents();
            content= getFText(filePath);
            content= replace(content, chr(0).ToString(), "");
            whiteWebStat(content);

        }
    }
    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");

    rw(getMsg1("更新全部统计成功，正在进入" + cStr(Request["lableTitle"]) + "列表...", url));
    writeSystemLog("", "更新网站统计"); //系统日志
    return "";
}
//清除全部网站统计 20160329
string clearWebsiteStat(){
    string url="";
    handlePower("清空网站统计"); //管理权限处理
    connexecute("delete from " + db_PREFIX + "websitestat");

    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");

    rw(getMsg1("清空网站统计成功，正在进入" + cStr(Request["lableTitle"]) + "列表...", url));
    writeSystemLog("", "清空网站统计"); //系统日志
    return "";
}
//更新今天网站统计
string updateTodayWebStat(){
    string content=""; string url=""; string dateStr=""; string dateMsg="";
    if( cStr(Request["date"]) != "" ){
        //dateStr = now() + cint(request("date"))
        dateStr=sAddTime(now(),"d",cInt(cStr(Request["date"])));
        dateMsg= "昨天";
    }else{
        dateStr= cStr(now());
        dateMsg= "今天";
    }

    handlePower("更新" + dateMsg + "统计"); //管理权限处理

    //call echo("datestr",datestr)
    connexecute("delete from " + db_PREFIX + "websitestat where dateclass='" + format_Time(dateStr, 2) + "'");
    content= getFText(adminDir + "/data/stat/" + format_Time(dateStr, 2) + ".txt");
    whiteWebStat(content);
    url= getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace");
    rw(getMsg1("更新" + dateMsg + "统计成功，正在进入" + cStr(Request["lableTitle"]) + "列表...", url));
    writeSystemLog("", "更新网站统计"); //系统日志
    return "";
}
//写入网站统计信息
string whiteWebStat(string content){
    string[] splStr; string[] splxx; string filePath=""; int nCount=0;
    string url=""; string s=""; string visitUrl=""; string viewUrl=""; string viewdatetime=""; string ip=""; string browser=""; string operatingsystem=""; string cookie=""; string screenwh=""; string moreInfo=""; string ipList=""; string dateClass="";
    splxx= aspSplit(content, vbCrlf() + "-------------------------------------------------" + vbCrlf());
    nCount= 0;
    foreach(var eachs in splxx){
        s=eachs;
        if( inStr(s, "当前：") > 0 ){
            nCount= nCount + 1;
            s= vbCrlf() + s + vbCrlf();
            dateClass= ADSql(getFileAttr(filePath, "3"));
            visitUrl= ADSql(getStrCut(s, vbCrlf() + "来访", vbCrlf(), 0));
            viewUrl= ADSql(getStrCut(s, vbCrlf() + "当前：", vbCrlf(), 0));
            viewdatetime= ADSql(getStrCut(s, vbCrlf() + "时间：", vbCrlf(), 0));
            ip= ADSql(getStrCut(s, vbCrlf() + "IP:", vbCrlf(), 0));
            browser= ADSql(getStrCut(s, vbCrlf() + "browser: ", vbCrlf(), 0));
            operatingsystem= ADSql(getStrCut(s, vbCrlf() + "operatingsystem=", vbCrlf(), 0));
            cookie= ADSql(getStrCut(s, vbCrlf() + "Cookies=", vbCrlf(), 0));
            screenwh= ADSql(getStrCut(s, vbCrlf() + "Screen=", vbCrlf(), 0));
            moreInfo= ADSql(getStrCut(s, vbCrlf() + "用户信息=", vbCrlf(), 0));
            browser= ADSql(getBrType(moreInfo));
            if( inStr(vbCrlf() + ipList + vbCrlf(), vbCrlf() + ip + vbCrlf())== 0 ){
                ipList= ipList + ip + vbCrlf();
            }

            viewdatetime= replace(viewdatetime, "来访", "00");
            if( isDate(viewdatetime)== false ){
                viewdatetime= "1988/07/12 10:10:10";
            }

            screenwh= left(screenwh, 20);
            if( 1== 2 ){
                echo("编号", nCount);
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

//详细网站统计
string websiteDetail(){
    string content=""; string[] splxx; string filePath="";
    string s=""; string ip=""; string ipList="";
    int nIP=0; int nPV=0; int i=0; string timeStr=""; string c="";

    handlePower("网站统计详细"); //管理权限处理

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
        echo(timeStr, "IP(" + nIP + ") PV(" + nPV + ")");
        if( i < 4 ){
            c= c + timeStr + " IP(" + nIP + ") PV(" + nPV + ")" + "<br>";
        }
    }

    setConfigFileBlock(WEB_CACHEFile, c, "#访客信息#");
    writeSystemLog("", "详细网站统计"); //系统日志

    return "";
}

//显示指定布局
void displayLayout(){
    string content=""; string lableTitle=""; string templateFile="";
    lableTitle= cStr(Request["lableTitle"]);
    templateFile= cStr(Request["templateFile"]);
    handlePower("显示" + lableTitle); //管理权限处理

    content= getTemplateContent(cStr(Request["templateFile"]));
    content= replace(content, "[$position$]", lableTitle);
    content= replaceValueParam(content, "lableTitle", lableTitle);


    //Robots.txt文件创建
    if( templateFile== "layout_makeRobots.html" ){
        content= replace(content, "[$bodycontent$]", getFText("/robots.txt"));
        //后台菜单地图
    }else if( templateFile== "layout_adminMap.html" ){
        content= replaceValueParam(content, "adminmapbody", getAdminMap());
        //管理模板
    }else if( templateFile== "layout_manageTemplates.html" ){
        content= displayTemplatesList(content);
        //生成html
    }else if( templateFile== "layout_manageMakeHtml.html" ){
        content= replaceValueParam(content, "columnList", getMakeColumnList());


    }


    content= handleDisplayLanguage(content, "handleDisplayLanguage"); //语言处理
    rw(content);
}
//获得生成栏目列表
string getMakeColumnList(){
    string c="";
    //栏目
    rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn order by sortrank asc", conn).ExecuteReader();

    while( rsx.Read()){
        if( cInt(cInt(rsx["nofollow"]))== 0 ){
            c= c + "<option value=\"" + cStr(rsx["id"]) + "\">" + cStr(rsx["columnname"]) + "</option>" + vbCrlf();
        }
    }
    return c;
}

//获得后台地图
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

//获得后台一级菜单列表
string getAdminOneMenuList(){
    string c=""; string focusStr=""; string addSql=""; string sql="";
    if( cStr(Session["adminflags"]) != "|*|" ){
        addSql= " and isDisplay<>0 ";
    }
    sql= "select * from " + db_PREFIX + "listmenu where parentid=-1 " + addSql + " order by sortrank";
    //检测SQL
    if( checkSql(sql)== false ){
        errorLog("出错提示：<br>function=getAdminOneMenuList<hr>sql=" + sql + "<br>");
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
//获得后台菜单列表
string getAdminMenuList(){
    string s=""; string c=""; string url=""; string selStr=""; string addSql=""; string sql="";
    if( cStr(Session["adminflags"]) != "|*|" ){
        addSql= " and isDisplay<>0 ";
    }
    sql= "select * from " + db_PREFIX + "listmenu where parentid=-1 " + addSql + " order by sortrank";
    //检测SQL
    if( checkSql(sql)== false ){
        errorLog("出错提示：<br>function=getAdminMenuList<hr>sql=" + sql + "<br>");
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
//处理模板列表
string displayTemplatesList(string content){
    string templatesFolder=""; string templatePath=""; string templatePath2=""; string templateName=""; string defaultList=""; string folderList=""; string[] splStr; string s=""; string c=""; string s1=""; string s2=""; string s3="";
    string[] splTemplatesFolder;
    //加载网址配置
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

                    s1= getStrCut(content, "<!--启用 start-->", "<!--启用 end-->", 2);
                    s2= getStrCut(content, "<!--恢复数据 start-->", "<!--恢复数据 end-->", 2);
                    s3= getStrCut(content, "<!--删除模板 start-->", "<!--删除模板 end-->", 2);

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
//应用模板
bool isOpenTemplate(){
    string templatePath=""; string templateName=""; string editValueStr=""; string url="";

    handlePower("启用模板"); //管理权限处理

    templatePath= cStr(Request["templatepath"]);
    templateName= cStr(Request["templatename"]);

    if( getRecordCount(db_PREFIX + "website", "")== 0 ){
        connexecute("insert into " + db_PREFIX + "website(webtitle) values('测试')");
    }


    editValueStr= "webtemplate='" + templatePath + "',webimages='" + templatePath + "Images/'";
    editValueStr= editValueStr + ",webcss='" + templatePath + "css/',webjs='" + templatePath + "Js/'";
    connexecute("update " + db_PREFIX + "website set " + editValueStr);
    url= "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=模板";



    rw(getMsg1("启用模板成功，正在进入模板界面...", url));
    writeSystemLog("", "应用模板" + templatePath); //系统日志
    return false;
}
//删除模板
string delTemplate(){
    string templateDir=""; string toTemplateDir=""; string url="";
    templateDir= replace(cStr(Request["templateDir"]), "\\", "/");
    handlePower("删除模板"); //管理权限处理
    toTemplateDir= mid(templateDir, 1, inStrRev(templateDir, "/")) + "#" + mid(templateDir, inStrRev(templateDir, "/") + 1,-1) + "_" + format_Time(now(), 11);
    //call die(toTemplateDir)
    moveFolder(templateDir, toTemplateDir);

    url= "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=模板";
    rw(getMsg1("删除模板完成，正在进入模板界面...", url));
    return "";
}
//执行SQL
string executeSQL(){
    string sqlvalue="";
    sqlvalue= "delete from " + db_PREFIX + "WebSiteStat";
    if( cStr(Request["sqlvalue"]) != "" ){
        sqlvalue= cStr(Request["sqlvalue"]);
        conn=openConn();
        conn.Open();

        //检测SQL
        if( checkSql(sqlvalue)== false ){
            errorLog("出错提示：<br>sql=" + sqlvalue + "<br>");
            return "";
        }
        echo("执行SQL语句成功", sqlvalue);
    }
    if( cStr(Session["adminusername"])== "PAAJCMS" ){
        rw("<form id=\"form1\" name=\"form1\" method=\"post\" action=\"?act=executeSQL\"  onSubmit=\"if(confirm('你确定要操作吗？\\n操作后将不可恢复')){return true}else{return false}\">SQL<input name=\"sqlvalue\" type=\"text\" id=\"sqlvalue\" value=\"" + sqlvalue + "\" size=\"80%\" /><input type=\"submit\" name=\"button\" id=\"button\" value=\"执行\" /></form>");
    }else{
        rw("你没有权限执行SQL语句");
    }
    return "";
}





</script>
