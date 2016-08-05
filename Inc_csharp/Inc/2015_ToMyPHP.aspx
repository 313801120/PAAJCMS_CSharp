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
//与php通用   我的后台

//获得网站底部内容aa
string XY_AP_WebSiteBottom(string action){
    string s="";string url="";
    if( inStr(cfg_webSiteBottom, "[$aoutadd$]") > 0 ){
        cfg_webSiteBottom= getDefaultValue(action); //获得默认内容
        connexecute("update " + db_PREFIX + "website set websitebottom='" + ADSql(cfg_webSiteBottom) + "'");
    }

    s=cfg_webSiteBottom;
    //网站底部
    if( cStr(Request["gl"])== "edit" ){
        s= "<span>" + s + "</span>";
    }
    url= WEB_ADMINURL + "?act=addEditHandle&switchId=2&id=*&actionType=WebSite&lableTitle=站点配置&n=" + getRnd(11);
    s= handleDisplayOnlineEditDialog(url, s, "", "span");

    return s;
}

//asp与php版本
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


//加载文件
string XY_Include(string action){
    string templateFilePath=""; string Block=""; string startStr=""; string endStr=""; string content="";
    templateFilePath= lCase(rParam(action, "File"));
    Block= lCase(rParam(action, "Block"));

    string findstr=""; string replaceStr=""; //查找字符，替换字符
    findstr= moduleFindContent(action, "findstr"); //先找块
    replaceStr= moduleFindContent(action, "replacestr"); //先找块

    templateFilePath= handleFileUrl(templateFilePath); //处理文件路径
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
    //替换读出来的内容
    if( findstr != "" ){
        content= replace(content, findstr, replaceStr);
    }

    return content;
}

//栏目菜单
string XY_AP_ColumnMenu(string action){
    string defaultStr=""; string thisId=""; string parentid="";string c="";string s="";
    parentid= aspTrim(rParam(action, "parentid"));
    parentid=getColumnId(parentid);

    if( parentid== "" ){ parentid= "-1" ;}

    thisId= glb_columnId;
    if( thisId== "" ){ thisId= "-1" ;}
    defaultStr= getDefaultValue(action); //获得默认内容

    defaultStr=defaultStr + "[topnav]"+ parentid +"[/topnav]";
    string XY_AP_ColumnMenu= showColumnList( parentid, "webcolumn", "columnname",thisId , 0, defaultStr);

    return XY_AP_ColumnMenu;
}



//显示栏目列表
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
    //追加sql
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    string XY_AP_ColumnList= XY_AP_GeneralList(action, "WebColumn", sql);

    return XY_AP_ColumnList;
}

//显示文章列表
string XY_AP_ArticleList(string action){
    string sql=""; string addSql=""; string columnName=""; string columnId=""; string topNumb=""; string idRand=""; string[] splStr; string s=""; string columnIdList="";
    action= replaceGlobleVariable(action); //处理下替换标签
    sql= rParam(action, "sql");
    topNumb= rParam(action, "topNumb");


    //id随机
    idRand= lCase(rParam(action, "rand"));
    if( idRand== "true" || idRand== "1" ){
        sql= sql + " where id in(" + getRandArticleId("", topNumb) + ")";
    }

    //栏目名称 对栏目数组处理如 模板分享下载[Array]CSS3[Array]HTML5
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
    //追加sql
    addSql= rParam(action, "addSql");
    //call echo("addsql",addsql)
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    sql= replaceGlobleVariable(sql);
    //call echo(RParam(action, "columnName") ,sql)
    return XY_AP_GeneralList(action, "ArticleDetail", sql);
}

//显示评论列表
string XY_AP_CommentList(string action){
    string itemID=""; string sql=""; string addSql="";
    addSql= rParam(action, "addsql");
    itemID= rParam(action, "itemID");
    itemID= replaceGlobleVariable(itemID);

    if( itemID != "" ){
        sql= " where itemID=" + itemID;
    }
    //追加sql
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    return XY_AP_GeneralList(action, "TableComment", sql);
}

//显示搜索统计
string XY_AP_SearchStatList(string action){
    string addSql="";
    addSql= rParam(action, "addSql");
    return XY_AP_GeneralList(action, "SearchStat", addSql);
}
//显示友情链接
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
//通用信息列表
string XY_AP_GeneralList(string action, string tableName, string addSql){
    string title=""; string topNumb=""; int nTop=0; bool isB; string sql="";
    string columnName=""; string columnEnName=""; string aboutcontent=""; string bodyContent=""; string showTitle="";
    string bannerImage=""; string smallImage=""; string bigImage=""; string id="";
    string defaultStr=""; int i=0; int j=0; string s=""; string c=""; string startStr=""; string endStr=""; string url="";
    string sNoFollow=""; //不追踪 20141222
    defaultStr= getDefaultValue(action); //获得默认内容
    int nModI=0; //余循环20150112
    sNoFollow= aspTrim(lCase(rParam(action, "noFollow"))); //不追踪
    string lableTitle=""; //标题标题
    string target=""; //a链接打开目标方式
    string adddatetime=""; //添加时间
    bool isFocus;
    string fieldNameList=""; //字段列表
    string abcolorStr=""; //A加粗和颜色
    string atargetStr=""; //A链接打开方式
    string atitleStr=""; //A链接的title20160407
    string anofollowStr=""; //A链接的nofollow
    string[] splFieldName; string fieldName=""; string replaceStr=""; int nI2=0;string idPage="";


    tableName=lCase(tableName);		//转小写
    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "字段列表");
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
    //追加sql
    if( addSql != "" ){
        sql= getWhereAnd(sql, addSql);
    }
    sql= replaceGlobleVariable(sql); //替换全局变量

    //检测SQL
    if( checkSql(sql)== false ){
        errorLog("出错提示：<br>action=" + action + "<hr>sql=" + sql + "<br>");
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
        //#【PHP】$rs=mysql_fetch_array($rsObj);                                            //给PHP用，因为在 asptophp转换不完善
        rs.Read();
        isFocus= false; //交点为假
        id= cStr(rs["id"]);
        //【导航】
        if( tableName== "webcolumn" ){
            if( isMakeHtml== true ){
                url= getRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "/nav" + cStr(rs["id"]));
            }else{
                url= handleWebUrl("?act=nav&columnName=" + cStr(rs["columnname"])); //会追加gl等参数
                if( cStr(rs["customaurl"]) != "" ){
                    url= cStr(rs["customaurl"]);
                    url= replaceGlobleVariable(url);
                }
            }
            //全局栏目名称为空则为自动定位首页 追加(20160128)
            if( glb_columnName== "" && cStr(rs["columntype"])== "首页" ){
                glb_columnName= cStr(rs["columnname"]);
            }
            if( cStr(rs["columnname"])== glb_columnName ){
                isFocus= true;
            }
            //【文章】
        }else if( tableName== "articledetail" ){
            if( isMakeHtml== true ){
                url= getRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "detail/detail" + cStr(rs["id"]));
            }else{
                url= handleWebUrl("?act=detail&id=" + cStr(rs["id"])); //会追加gl等参数
                if( cStr(rs["customaurl"]) != "" ){
                    url= cStr(rs["customaurl"]);
                }
            }
            //评论
        }else if( tableName== "tablecomment" ){

        }

        //A链接添加颜色
        abcolorStr= "";
        if( inStr(fieldNameList, ",titlecolor,") > 0 ){
            //A链接颜色
            if( cStr(rs["titlecolor"]) != "" ){
                abcolorStr= "color:" + cStr(rs["titlecolor"]) + ";";
            }
        }
        if( inStr(fieldNameList, ",flags,") > 0 ){
            //A链接加粗
            if( inStr(cStr(rs["flags"]), "|b|") > 0 ){
                abcolorStr= abcolorStr + "font-weight:bold;";
            }
        }
        if( abcolorStr != "" ){
            abcolorStr= " style=\"" + abcolorStr + "\"";
        }

        //打开方式2016
        if( inStr(fieldNameList, ",target,") > 0 ){
            atargetStr= IIF(cStr(rs["target"]) != "", " target=\"" + cStr(rs["target"]) + "\"", "");
        }

        //A的title
        if( inStr(fieldNameList, ",title,") > 0 ){
            atitleStr= IIF(cStr(rs["title"]) != "", " title=\"" + cStr(rs["title"]) + "\"", "");
        }

        //A的nofollow
        if( inStr(fieldNameList, ",nofollow,") > 0 ){
            anofollowStr= IIF(cInt(rs["nofollow"]) != 0, " rel=\"nofollow\"", "");
        }

        //交点判断(给栏目导航用的)
        if( isFocus== true ){
            startStr= "[list-focus]" ; endStr= "[/list-focus]";
        }else if( inStr(defaultStr, "[list-start]")>0 && inStr(defaultStr, "[/list-start]")>0 ){
            startStr= "[list-start]" ; endStr= "[/list-start]";
        }else{
            startStr= "[list-" + i + "]" ; endStr= "[/list-" + i + "]";
        }

        //在最后时排序当前交点20160202
        if( i== nTop && isFocus== false ){
            if( inStr(defaultStr, "[list-end]")>0 && inStr(defaultStr, "[/list-end]")>0 ){
                startStr= "[list-end]" ; endStr= "[/list-end]";
            }
        }

        //例[list-mod2]  [/list-mod2]    20150112
        for( nModI= 6 ; nModI>= 2 ; nModI--){
            if( inStr(defaultStr, startStr)== 0 && i % nModI== 0 ){
                startStr= "[list-mod" + nModI + "]" ; endStr= "[/list-mod" + nModI + "]";
                if( inStr(defaultStr, startStr) > 0 ){
                    break;
                }
            }
        }

        //没有则用默认
        if( inStr(defaultStr, startStr)== 0 ){
            startStr= "[list]" ; endStr= "[/list]";
        }

        if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){

            s= strCut(defaultStr, startStr, endStr, 2);

            s= replaceValueParam(s, "i", i); //循环编号
            s= replaceValueParam(s, "编号", i); //循环编号
            s= replaceValueParam(s, "id", cStr(rs["id"])); //id编号 因为获得字段他不获得id
            s= replaceValueParam(s, "url", url); //网址
            s= replaceValueParam(s, "aurl", "href=\"" + url + "\""); //网址
            s= replaceValueParam(s, "abcolor", abcolorStr); //A链接加颜色与加粗
            s= replaceValueParam(s, "atitle", atitleStr); //A链接title
            s= replaceValueParam(s, "anofollow", anofollowStr); //A链接nofollow
            s= replaceValueParam(s, "atarget", atargetStr); //A链接打开方式



            for( nI2= 0 ; nI2<= uBound(splFieldName); nI2++){
                if( splFieldName[nI2] != "" ){
                    fieldName= splFieldName[nI2];
                    replaceStr= cStr(rs[fieldName]) + "";
                    s= replaceValueParam(s, fieldName, replaceStr);
                }
            }


            //开始位置加Dialog内容
            startStr= "[list-" + i + " startdialog]" ; endStr= "[/list-" + i + " startdialog]";
            if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
                s= strCut(defaultStr, startStr, endStr, 2) + s;
            }
            //结束位置加Dialog内容
            startStr= "[list-" + i + " enddialog]" ; endStr= "[/list-" + i + " enddialog]";
            if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
                s= s + strCut(defaultStr, startStr, endStr, 2);
            }

            //加控制
            //【导航】
            if( tableName== "webcolumn" ){
                url= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);

                //【评论】
            }else if( tableName=="tablecomment" ){
                idPage=getThisIdPage(db_PREFIX + tableName ,cStr(rs["id"]),10);
                url= WEB_ADMINURL + "?act=addEditHandle&actionType=TableComment&lableTitle=评论&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page="+ idPage +"&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);

                //【文章】
            }else if( tableName== "articledetail" ){
                idPage=getThisIdPage(db_PREFIX + tableName ,cStr(rs["id"]),10);
                url= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page="+ idPage +"&parentid="+ cStr(rs["parentid"]) +"&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);


                s= replaceValueParam(s, "columnurl", getColumnUrl(cStr(rs["parentid"]), "")); //文章对应栏目URL 20160304
                s= replaceValueParam(s, "columnname", getColumnName(cStr(rs["parentid"]))); //文章对应栏目名称 20160304
            }
            s= handleDisplayOnlineEditDialog(url, s, "", "div|li|span"); //处理是否添加在线修改管理器
            c= c + s;
        }
    }

    //开始内容加Dialog内容
    startStr= "[dialog start]" ; endStr= "[/dialog start]";
    if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
        c= strCut(defaultStr, startStr, endStr, 2) + c;
    }
    //结束内容加Dialog内容
    startStr= "[dialog end]" ; endStr= "[/dialog end]";
    if( inStr(defaultStr, startStr) > 0 && inStr(defaultStr, endStr) > 0 ){
        c= c + strCut(defaultStr, startStr, endStr, 2);
    }
    return c;
}


//处理获得表内容
string XY_handleGetTableBody(string action, string tableName, string fieldParamName, string defaultFileName, string adminUrl){
    string url=""; string content=""; string id=""; string sql=""; string addSql=""; string fieldName=""; string fieldParamValue=""; string fieldNameList=""; string sLen=""; string delHtmlYes=""; string trimYes="";string defaultStr="";
    string snoisonhtml="";string intoFieldStr="";string valuesStr="";
    fieldName= rParam(action, "fieldname"); //字段名称
    snoisonhtml= rParam(action, "noisonhtml");					 //不生成html
    //dim nonull nonull=RParam(action, "noisonhtml")							'内容不能为空20160716 home

    if( snoisonhtml=="true" ){
        intoFieldStr=",isonhtml";
        valuesStr=",0";
    }

    fieldNameList= getHandleFieldList(db_PREFIX + tableName, "字段列表");
    //字段名称不为空，并且要在表字段里
    if( fieldName== "" || inStr(fieldNameList, "," + fieldName + ",")== 0 ){
        fieldName= defaultFileName;
    }
    fieldName= lCase(fieldName); //转为小写，因为在PHP里是全小写的

    fieldParamValue= rParam(action, fieldParamName); //截取字段内容
    id= handleNumber(rParam(action, "id")); //获得ID
    addSql= " where " + fieldParamName + "='" + fieldParamValue + "'";
    if( id != "" ){
        addSql= " where id=" + id;
    }

    content= getDefaultValue(action) ;defaultStr=content; //获得默认内容
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
        //自动添加 20160113
        if( rParam(action, "autoadd")== "true" ){
            connexecute("insert into " + db_PREFIX + tableName + " (" + fieldParamName + "," + fieldName + intoFieldStr + ") values('" + fieldParamValue + "','" + ADSql(content) + "'"+ valuesStr +")");
        }
    }

    //删除Html
    delHtmlYes= rParam(action, "delHtml"); //是否删除Html
    if( delHtmlYes== "true" ){ content= replace(delHtml(content), "<", "&lt;") ;}//HTML处理
    //删除两边空格
    trimYes= rParam(action, "trim"); //是否删除两边空格
    if( trimYes== "true" ){ content= trimVbCrlf(content) ;}

    //截取字符处理
    sLen= rParam(action, "len"); //字符长度值
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

//获得单页内容
string XY_AP_GetOnePageBody(string action){
    string adminUrl="";
    adminUrl= WEB_ADMINURL + "?act=addEditHandle&actionType=OnePage&lableTitle=单页&nPageSize=10&page=&switchId=2";
    return XY_handleGetTableBody(action, "onepage", "title", "bodycontent", adminUrl);
}

//获得导航内容
string XY_AP_GetColumnBody(string action){
    string adminUrl="";
    adminUrl= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&switchId=2";
    return XY_handleGetTableBody(action, "webcolumn", "columnname", "bodycontent", adminUrl);
}

//显示文章内容
string XY_AP_GetArticleBody(string action){
    string adminUrl="";
    adminUrl= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&switchId=2";
    return XY_handleGetTableBody(action, "articledetail", "title", "bodycontent", adminUrl);
}


//获得栏目URL
string XY_GetColumnUrl(string action){
    string columnName=""; string url="";
    columnName= rParam(action, "columnName");
    url= getColumnUrl(columnName, "name");
    //handleWebUrl  有对gl处理

    //If Request("gl") <> "" Then
    //    url = url & "&gl=" & Request("gl")
    //End If
    return url;
}
//获得导航url 辅助之前网站开发时用的 20160722  不可删除
string XY_GetNavUrl(string action){
    string columnName=""; string url="";string did="";string sid="";
    did= rParam(action, "did");
    sid= rParam(action, "sid");
    columnName=did;
    if( sid!="" ){
        columnName=sid;
    }
    url= getColumnUrl(columnName, "name");
    //handleWebUrl  有对gl处理

    //If Request("gl") <> "" Then
    //    url = url & "&gl=" & Request("gl")
    //End If
    return url;
}

//获得文章URL
string XY_GetArticleUrl(string action){
    string title=""; string url="";
    title= rParam(action, "title");
    url= getArticleUrl(title);
    //If Request("gl") <> "" Then
    //    url = url & "&gl=" & Request("gl")
    //End If
    return url;
}

//获得单页URL
string XY_GetOnePageUrl(string action){
    string title=""; string url="";
    title= rParam(action, "title");
    url= getOnePageUrl(title);
    //If Request("gl") <> "" Then
    //    url = url & "&gl=" & Request("gl")
    //End If
    return url;
}


//获得单个字段内容
string XY_AP_GetFieldValue(string action, string sql, string fieldName){
    string title=""; string content="";
    rs = new OleDbCommand(sql, conn).ExecuteReader();

    if( rs.Read() ){
        content= cStr(rs[fieldName]);
    }
    return content;
}


//Js版网站统计
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
    s= s + "+\"&co=\"+escape(document.cookie)"; //收集cookie 不需要则屏蔽掉
    s= s + "+\" \\'><\\/script>\");</"+"script>";
    return s;
}



//普通链接A
string XY_HrefA(string action){
    string content=""; string Href=""; string c=""; string AContent=""; string AType=""; string url=""; string title="";
    action= handleInModule(action, "start");
    content= rParam(action, "Content");
    AType= rParam(action, "Type");
    if( AType== "收藏" ){
        //第一种方法
        //Url = "window.external.addFavorite('"& WebUrl &"','"& WebTitle &"')"
        url= "shoucang(document.title,window.location)";
        c= "<a href='javascript:;' onClick=\"" + url + "\" " + setHtmlParam(action, "target|title|alt|id|class|style") + ">" + content + "</a>";
    }else if( AType== "设为首页" ){
        //第一种方法
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



//布局20151231
string XY_Layout(string action){
    string layoutName=""; string s=""; string c=""; string sourceStr=""; string replaceStr=""; string[] splSource; string[] splReplace; int i=0;

    layoutName= rParam(action, "layoutname");
    rs = new OleDbCommand("select * from " + db_PREFIX + "weblayout where layoutname='" + layoutName + "'", conn).ExecuteReader();

    if( rs.Read() ){
        c= cStr(rs["bodycontent"]);

        sourceStr= cStr(rs["sourcestr"]); //源内容 被替换内容
        replaceStr= cStr(rs["replacestr"]); //替换内容
        splSource= aspSplit(sourceStr, "[Array]"); //源内容数组
        splReplace= aspSplit(replaceStr, "[Array]"); //替换内容数组

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

//模块20151231
string XY_Module(string action){
    string moduleName=""; string s=""; string c=""; string sourceStr=""; string replaceStr=""; string[] splSource; string[] splReplace; int i=0;
    moduleName= rParam(action, "modulename");
    rs = new OleDbCommand("select * from " + db_PREFIX + "webmodule where modulename='" + moduleName + "'", conn).ExecuteReader();

    if( rs.Read() ){
        c= cStr(rs["bodycontent"]);

        sourceStr= rParam(action, "sourceStr"); //源内容 被替换内容
        replaceStr= rParam(action, "replaceStr"); //替换内容

        splSource= aspSplit(sourceStr, "[Array]"); //源内容数组
        splReplace= aspSplit(replaceStr, "[Array]"); //替换内容数组

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

//显示包裹块20160127
string XY_DisplayWrap( string action){
    string content="";
    content= getDefaultValue(action);
    return content;
}




//嵌套标题 测试
string XY_getLableValue(string action){
    string title=""; string content=""; string c="";
    //call echo("Action",Action)
    title= rParam(action, "title");
    content= rParam(action, "content");
    c= c + "title=" + getContentRunStr(title) + "<hr>";
    c= c + "content=" + getContentRunStr(content) + "<hr>";
    string XY_getLableValue= c;
    echo("title", title);
    return "【title=】【" + title + "】";
}
//标题在搜索引擎中搜索列表
string XY_TitleInSearchEngineList(string action){
    string title=""; string sType="";string divclass="";string spanclass="";string s="";string c="";

    title= rParam(action, "title");
    sType= rParam(action, "sType");
    divclass= rParam(action, "divclass");
    spanclass= rParam(action, "spanclass");

    s="<strong>更多关于《" + title + "》</strong>";
    if( divclass!="" ){
        s="<div class=\""+ divclass +"\">"+ s +"</div>";
    }else if( spanclass!="" ){
        s="<span class=\""+ spanclass +"\">"+ s +"</span>" + "<br>";
    }else{
        s=s + "<br>";
    }
    c= c + s + vbCrlf();
    c= c + "<ul class=\"list\"> " + vbCrlf();
    c= c + "<li><a href=\"https://www.baidu.com/s?ie=gb2312&word=" + title + "\" rel=\"nofollow\" target=\"_blank\">【baidu搜索】在百度里搜索(" + title + ")</a></li>" + vbCrlf();
    c= c + "<li><a href=\"http://www.haosou.com/s?ie=gb2312&q=" + title + "\" rel=\"nofollow\" target=\"_blank\">【haosou搜索】在好搜里搜索(" + title + ")</a></li>" + vbCrlf();
    c= c + "<li><a href=\"https://search.yahoo.com/search;_ylt=A86.JmbkJatWH5YARmebvZx4?toggle=1&cop=mss&ei=gb2312&fr=yfp-t-901&fp=1&p=" + title + "\" rel=\"nofollow\" target=\"_blank\">【yahoo搜索】在雅虎里搜索(" + title + ")</a></li>" + vbCrlf();

    c= c + "<li><a href=\"https://www.sogou.com/sogou?ie=utf8&query=" + toUTFChar(title) + "\" rel=\"nofollow\" target=\"_blank\">【sogou搜索】在搜狗里搜索(" + title + ")</a></li>" + vbCrlf();
    c= c + "<li><a href=\"http://www.youdao.com/search?ue=utf8&q=" + toUTFChar(title) + "\" rel=\"nofollow\" target=\"_blank\">【youdao搜索】在有道里搜索(" + title + ")</a></li>" + vbCrlf();
    c= c + "<li><a href=\"http://search.yam.com/Search/Web/DefaultKSA.aspx?SearchType=web&l=0&p=0&k=" + toUTFChar(title) + "\" rel=\"nofollow\" target=\"_blank\">【yam搜索(google提供技术)】在蕃薯藤里搜索(" + title + ")</a></li>" + vbCrlf();


    c= c + "<li><a href=\"http://cn.bing.com/search?q=" + toUTFChar(title) + "\" rel=\"nofollow\" target=\"_blank\">【bing搜索】在必应里搜索(" + title + ")</a></li>" + vbCrlf();
    c= c + "</ul>" + vbCrlf();

    return c;
}

//URL加密
string XY_escape(string action){
    string content="";
    content= rParam(action, "content");
    return escape(content);
}
//URL解密
string XY_unescape(string action){
    string content="";
    content= rParam(action, "content");
    return escape(content);
}
//获得网址
string XY_getUrl(string action){
    string stype="";
    stype= rParam(action, "stype");
    return getThisUrlNoParam();
}



//For循环处理
string XY_ForArray(string action){
    string sListArray=""; string sSplitStr=""; string defaultStr=""; string[] splStr; int nForI=0; string title=""; string s=""; string c=""; string sNLoop="";
    sListArray= atRParam(action, "arraylist"); //atRParam获得结果处理动作，但不替换动作内容
    sSplitStr= rParam(action, "splitstr");
    sNLoop= rParam(action, "nloop"); //循环数

    if( sListArray== "" && sNLoop!="" ){
        sListArray= copyStr("循环" + sSplitStr, cInt(sNLoop));
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

//文章位置显示信息{}为有动作的
string XY_detailPosition(string action){
    string c="";string stype="";
    stype= rParam(action, "stype");
    stype="|"+ stype +"|";
    c= "<a href=\""+ getColumnUrl("首页","type") +"\">首页</a>";
    //call echo("type",getColumnUrl("首页","type"))
    if( glb_columnName != "" ){
        c= c + " >> <a href=\"" + getColumnUrl(glb_columnName, "name") + "\">" + glb_columnName + "</a>";
    }
    //20160330
    if( glb_locationType== "detail" ){
        if( inStr(stype,"|显示文章标题|")>0 ){
            c= c + " >> " + glb_detailTitle;
        }else if( inStr(stype,"|隐藏文章标题|")>0 ){
        }else{
            c= c + " >> 查看内容";
        }
    }
    //尾部追加内容

    return c;
}

</script>
