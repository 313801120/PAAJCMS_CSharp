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
	//关闭数据库
	if (conn.State != ConnectionState.Closed){
		conn.Close(); 
	}
	
} 



//=========


//处理动作   ReplaceValueParam为控制字符显示方式
string handleAction(string content){
    string startStr=""; string endStr=""; string actionList=""; string[] splStr; string action=""; string s=""; bool isHnadle;
    startStr= "{$" ; endStr= "$}";
    actionList= getArrayNew(content, startStr, endStr, true, true);
    //Call echo("ActionList ", ActionList)
    splStr= aspSplit(actionList, "$Array$");
    foreach(var eachs in splStr){
        s=eachs;
        action= aspTrim(s);
        action= handleInModule(action, "start"); //处理\'替换掉
        if( action != "" ){
            action= aspTrim(mid(action, 3, len(action) - 4)) + " ";
            //call echo("s",s)
            isHnadle= true; //处理为真
            //{VB #} 这种是放在图片路径里，目的是为了在VB里不处理这个路径
            if( checkFunValue(action, "# ")== true ){
                action= "";
                //测试
            }else if( checkFunValue(action, "GetLableValue ")== true ){
                action= XY_getLableValue(action);
                //标题在搜索引擎里列表
            }else if( checkFunValue(action, "TitleInSearchEngineList ")== true ){
                action= XY_TitleInSearchEngineList(action);

                //加载文件
            }else if( checkFunValue(action, "Include ")== true ){
                action= XY_Include(action);
                //栏目列表
            }else if( checkFunValue(action, "ColumnList ")== true ){
                action= XY_AP_ColumnList(action);
                //文章列表
            }else if( checkFunValue(action, "ArticleList ")== true || checkFunValue(action, "CustomInfoList ")== true ){
                action= XY_AP_ArticleList(action);
                //评论列表
            }else if( checkFunValue(action, "CommentList ")== true ){
                action= XY_AP_CommentList(action);
                //搜索统计列表
            }else if( checkFunValue(action, "SearchStatList ")== true ){
                action= XY_AP_SearchStatList(action);
                //友情链接列表
            }else if( checkFunValue(action, "Links ")== true ){
                action= XY_AP_Links(action);

                //显示单页内容
            }else if( checkFunValue(action, "GetOnePageBody ")== true || checkFunValue(action, "MainInfo ")== true ){
                action= XY_AP_GetOnePageBody(action);
                //显示文章内容
            }else if( checkFunValue(action, "GetArticleBody ")== true ){
                action= XY_AP_GetArticleBody(action);
                //显示栏目内容
            }else if( checkFunValue(action, "GetColumnBody ")== true ){
                action= XY_AP_GetColumnBody(action);

                //获得栏目URL
            }else if( checkFunValue(action, "GetColumnUrl ")== true ){
                action= XY_GetColumnUrl(action);
                //获得栏目ID
            }else if( checkFunValue(action, "GetColumnId ")== true ){
                action= XY_GetColumnId(action);
                //获得文章URL
            }else if( checkFunValue(action, "GetArticleUrl ")== true ){
                action= XY_GetArticleUrl(action);
                //获得单页URL
            }else if( checkFunValue(action, "GetOnePageUrl ")== true ){
                action= XY_GetOnePageUrl(action);
                //获得导航URL
            }else if( checkFunValue(action, "GetNavUrl ")== true ){
                action= XY_GetNavUrl(action);



                //------------------- 模板模块区 -----------------------
                //显示包裹块 作用不大
            }else if( checkFunValue(action, "DisplayWrap ")== true ){
                action= XY_DisplayWrap(action);
                //显示布局
            }else if( checkFunValue(action, "Layout ")== true ){
                action= XY_Layout(action);
                //显示模块
            }else if( checkFunValue(action, "Module ")== true ){
                action= XY_Module(action);
                //读模块内容
            }else if( checkFunValue(action, "ReadTemplateModule ")== true ){
                action= XY_ReadTemplateModule(action);
                //获得内容模块 20150108
            }else if( checkFunValue(action, "GetContentModule ")== true ){
                action= XY_ReadTemplateModule(action);
                //读模板样式并设置标题与内容   软件里有个栏目Style进行设置
            }else if( checkFunValue(action, "ReadColumeSetTitle ")== true ){
                action= XY_ReadColumeSetTitle(action);


                //------------------- 其它区 -----------------------
                //显示JS渲染ASP/PHP/VB等程序的编辑器
            }else if( checkFunValue(action, "displayEditor ")== true ){
                action= displayEditor(action);
                //Js版网站统计
            }else if( checkFunValue(action, "JsWebStat ")== true ){
                action= XY_JsWebStat(action);

                //------------------- 链接区 -----------------------
                //普通链接A
            }else if( checkFunValue(action, "HrefA ")== true ){
                action= XY_HrefA(action);
                //栏目菜单(引用后台栏目程序)
            }else if( checkFunValue(action, "ColumnMenu ")== true ){
                action= XY_AP_ColumnMenu(action);

                //------------------- 循环处理 -----------------------
                //For循环处理
            }else if( checkFunValue(action, "ForArray ")== true ){
                action= XY_ForArray(action);

                //------------------- 待分区 -----------------------
                //网站底部
            }else if( checkFunValue(action, "WebSiteBottom ")== true || checkFunValue(action, "WebBottom ")== true ){
                action= XY_AP_WebSiteBottom(action);
                //显示网站栏目 20160331
            }else if( checkFunValue(action, "DisplayWebColumn ")== true ){
                action= XY_DisplayWebColumn(action);
                //URL加密
            }else if( checkFunValue(action, "escape ")== true ){
                action= XY_escape(action);
                //URL解密
            }else if( checkFunValue(action, "unescape ")== true ){
                action= XY_unescape(action);
                //asp与php版本
            }else if( checkFunValue(action, "EDITORTYPE ")== true ){
                action= XY_EDITORTYPE(action);

                //获得网址
            }else if( checkFunValue(action, "getUrl ")== true ){
                action= XY_getUrl(action);

                //文章位置显示信息{}为有动作的
            }else if( checkFunValue(action, "detailPosition ")== true ){
                action= XY_detailPosition(action);




                //暂时不屏蔽
            }else if( checkFunValue(action, "copyTemplateMaterial ")== true ){
                action= "";
            }else if( checkFunValue(action, "clearCache ")== true ){
                action= "";
            }else{
                isHnadle= false; //处理为假
            }
            //注意这样，有的则不显示 晕 And IsNul(action)=False
            if( isNul(action)== true ){ action= "" ;}
            if( isHnadle== true ){
                content= replace(content, s, action);
            }
        }
    }
    return content;
}

//显示网站栏目 新版 把之前网站 导航程序改进过来的
string XY_DisplayWebColumn(string action){
    int i=0; string c=""; string s=""; string url=""; string sql=""; string dropDownMenu=""; string focusType=""; string addSql="";
    bool isConcise; //简洁显示20150212
    string styleId=""; string styleValue=""; //样式ID与样式内容
    string cssNameAddId="";
    string shopnavidwrap=""; //是否显示栏目ID包

    styleId= phpTrim(rParam(action, "styleID"));
    styleValue= phpTrim(rParam(action, "styleValue"));
    addSql= phpTrim(rParam(action, "addSql"));
    shopnavidwrap= phpTrim(rParam(action, "shopnavidwrap"));
    //If styleId <> "" Then
    //Call ReadNavCSS(styleId, styleValue)
    //End If

    //为数字类型 则自动提取样式内容  20150615
    if( checkStrIsNumberType(styleValue) ){
        cssNameAddId= "_" + styleValue; //Css名称追加Id编号
    }
    sql= "select * from " + db_PREFIX + "webcolumn";
    //追加sql
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
        //网站栏目没有page位置处理
        url= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);
        s= handleDisplayOnlineEditDialog(url, s, "", "div|li|span"); //处理是否添加在线修改管理器

        c= c + s;

        //小类

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

//替换全局变量 {$cfg_websiteurl$}
string replaceGlobleVariable( string content){
    content= handleRGV(content, "{$cfg_webSiteUrl$}", cfg_webSiteUrl); //网址
    content= handleRGV(content, "{$cfg_webTemplate$}", cfg_webTemplate); //模板
    content= handleRGV(content, "{$cfg_webImages$}", cfg_webImages); //图片路径
    content= handleRGV(content, "{$cfg_webCss$}", cfg_webCss); //css路径
    content= handleRGV(content, "{$cfg_webJs$}", cfg_webJs); //js路径
    content= handleRGV(content, "{$cfg_webTitle$}", cfg_webTitle); //网站标题
    content= handleRGV(content, "{$cfg_webKeywords$}", cfg_webKeywords); //网站关键词
    content= handleRGV(content, "{$cfg_webDescription$}", cfg_webDescription); //网站描述

    content= handleRGV(content, "{$cfg_webSiteBottom$}", cfg_webSiteBottom); //网站底部内容

    content= handleRGV(content, "{$glb_columnId$}", glb_columnId); //栏目Id
    content= handleRGV(content, "{$glb_columnName$}", glb_columnName); //栏目名称
    content= handleRGV(content, "{$glb_columnType$}", glb_columnType); //栏目类型
    content= handleRGV(content, "{$glb_columnENType$}", glb_columnENType); //栏目英文类型

    content= handleRGV(content, "{$glb_Table$}", glb_table); //表
    content= handleRGV(content, "{$glb_Id$}", glb_id); //id

    content= handleRGV(content, "[$模块目录$]", "Module/"); //Module


    //兼容旧版本 渐渐把它去掉
    content= handleRGV(content, "{$WebImages$}", cfg_webImages); //图片路径
    content= handleRGV(content, "{$WebCss$}", cfg_webCss); //css路径
    content= handleRGV(content, "{$WebJs$}", cfg_webJs); //js路径
    content= handleRGV(content, "{$Web_Title$}", cfg_webTitle);
    content= handleRGV(content, "{$Web_KeyWords$}", cfg_webKeywords);
    content= handleRGV(content, "{$Web_Description$}", cfg_webDescription);


    content= handleRGV(content, "{$EDITORTYPE$}", EDITORTYPE); //后缀
    content= handleRGV(content, "{$WEB_VIEWURL$}", WEB_VIEWURL); //首页显示网址
    //文章用到
    content= handleRGV(content, "{$glb_articleAuthor$}", glb_articleAuthor); //文章作者
    content= handleRGV(content, "{$glb_articleAdddatetime$}", glb_articleAdddatetime); //文章添加时间
    content= handleRGV(content, "{$glb_articlehits$}", glb_articlehits); //文章添加时间

    content= handleRGV(content, "{$glb_upArticle$}", glb_upArticle); //上一篇文章
    content= handleRGV(content, "{$glb_downArticle$}", glb_downArticle); //下一篇文章
    content= handleRGV(content, "{$glb_aritcleRelatedTags$}", glb_aritcleRelatedTags); //文章标签组
    content= handleRGV(content, "{$glb_aritcleBigImage$}", glb_aritcleBigImage); //文章大图
    content= handleRGV(content, "{$glb_aritcleSmallImage$}", glb_aritcleSmallImage); //文章小图
    content= handleRGV(content, "{$glb_searchKeyWord$}", glb_searchKeyWord); //首页显示网址


    return content;
}

//处理替换
string handleRGV( string content, string findStr, string replaceStr){
    string lableName="";
    //对[$$]处理
    lableName= mid(findStr, 3, len(findStr) - 4) + " ";
    lableName= mid(lableName, 1, inStr(lableName, " ") - 1);
    content= replaceValueParam(content, lableName, replaceStr);
    content= replaceValueParam(content, lCase(lableName), replaceStr);
    //直接替换{$$}这种方式，兼容之前网站
    content= replace(content, findStr, replaceStr);
    content= replace(content, lCase(findStr), replaceStr);
    return content;
}

//加载网站配置信息
void loadWebConfig(){
    string templatedir="";
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "website", conn).ExecuteReader();

    if( rs.Read() ){
        cfg_webSiteUrl= phpTrim(cStr(rs["websiteurl"])); //网址
        cfg_webTemplate= webDir + phpTrim(cStr(rs["webtemplate"])); //模板路径
        cfg_webImages= webDir + phpTrim(cStr(rs["webimages"])); //图片路径
        cfg_webCss= webDir + phpTrim(cStr(rs["webcss"])); //css路径
        cfg_webJs= webDir + phpTrim(cStr(rs["webjs"])); //js路径
        cfg_webTitle= cStr(rs["webtitle"]); //网址标题
        cfg_webKeywords= cStr(rs["webkeywords"]); //网站关键词
        cfg_webDescription= cStr(rs["webdescription"]); //网站描述
        cfg_webSiteBottom= cStr(rs["websitebottom"]); //网站地底
        cfg_flags= cStr(rs["flags"]); //旗

        //改换模板20160202
        if( cStr(Request["templatedir"]) != "" ){
            //删除绝对目录前面的目录，不需要那个东西20160414
            templatedir= replace(handlePath(cStr(Request["templatedir"])), handlePath("/"), "/");
            //call eerr("templatedir",templatedir)

            if((inStr(templatedir, ":") > 0 || inStr(templatedir, "..") > 0) && getIP() != "127.0.0.1" ){
                eerr("提示", "模板目录有非法字符");
            }
            templatedir= handleHttpUrl(replace(templatedir, handlePath("/"), "/"));

            cfg_webImages= replace(cfg_webImages, cfg_webTemplate, templatedir);
            cfg_webCss= replace(cfg_webCss, cfg_webTemplate, templatedir);
            cfg_webJs= replace(cfg_webJs, cfg_webTemplate, templatedir);
            cfg_webTemplate= templatedir;
        }
    }
}

//网站位置 待完善
string thisPosition(string content){
    string c="";
    c= "<a href=\"" + getColumnUrl("首页", "type") + "\">首页</a>";
    if( glb_columnName != "" ){
        c= c + " >> <a href=\"" + getColumnUrl(glb_columnName, "name") + "\">" + glb_columnName + "</a>";
    }
    //20160330
    if( glb_locationType== "detail" ){
        c= c + " >> 查看内容";
    }
    //尾部追加内容
    c= c + positionEndStr;

    //call echo("glb_locationType",glb_locationType)

    content= replace(content, "[$detailPosition$]", c);
    content= replace(content, "[$detailTitle$]", glb_detailTitle);
    content= replace(content, "[$detailContent$]", glb_bodyContent);

    return content;
}

//显示管理列表
string getDetailList(string action, string content, string actionName, string lableTitle, string fieldNameList, int nPageSize, string sPage, string addSql){
    conn=openConn();
    conn.Open();

    string defaultStr=""; int i=0; string s=""; string c=""; string tableName=""; int j=0; string[] splxx; string sql=""; int nPage=0;
    int nX=0; string url=""; int nCount=0;
    string pageInfo=""; int nModI=0; string startStr=""; string endStr="";

    string fieldName=""; //字段名称
    string[] splFieldName; //分割字段

    string replaceStr=""; //替换字符
    tableName= lCase(actionName); //表名称
    string listFileName=""; //列表文件名称
    listFileName= rParam(action, "listFileName");
    string abcolorStr=""; //A加粗和颜色
    string atargetStr=""; //A链接打开方式
    string atitleStr=""; //A链接的title20160407
    string anofollowStr=""; //A链接的nofollow

    string id=""; string idPage="";
    id= rq("id");
    checkIDSQL(cStr(Request["id"]));

    if( fieldNameList== "*" ){
        fieldNameList= getHandleFieldList(db_PREFIX + tableName, "字段列表");
    }

    fieldNameList= specialStrReplace(fieldNameList); //特殊字符处理
    splFieldName= aspSplit(fieldNameList, ","); //字段分割成数组


    defaultStr= getStrCut(content, "<!--#body start#-->", "<!--#body end#-->", 2);



    pageInfo= getStrCut(content, "[page]", "[/page]", 1);
    if( pageInfo != "" ){
        content= replace(content, pageInfo, "");
    }
    //call eerr("pageInfo",pageInfo)

    sql= "select * from " + db_PREFIX + tableName + " " + addSql;
    //检测SQL
    if( checkSql(sql)== false ){
        errorLog("出错提示：<br>sql=" + sql + "<br>");
        return "";
    }
    rs = new OleDbCommand(sql, conn).ExecuteReader();


    nCount= rsRecordcount(sql);

    //为动态翻页网址
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

        //在最后时排序当前交点20160202
        if( i== nX ){
            startStr= "[list-end]" ; endStr= "[/list-end]";
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



                s= replaceValueParam(s, "url", url);
                s= replaceValueParam(s, "abcolor", abcolorStr); //A链接加颜色与加粗
                s= replaceValueParam(s, "atitle", atitleStr); //A链接title
                s= replaceValueParam(s, "anofollow", anofollowStr); //A链接nofollow
                s= replaceValueParam(s, "atarget", atargetStr); //A链接打开方式


            }
        }
        //call echo("tableName",tableName)
        idPage= getThisIdPage(db_PREFIX + tableName, cStr(rs["id"]), 10);
        //【留言】
        if( tableName== "guestbook" ){
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=GuestBook&lableTitle=留言&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page=" + idPage + "&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);

            //【默认显示文章】
        }else{
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=" + idPage + "&parentid=" + cStr(rs["parentid"]) + "&id=" + cStr(rs["id"]) + "&n=" + getRnd(11);
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
//默认列表模板
string defaultListTemplate(string sType, string sName){
    string c=""; string templateHtml=""; string listTemplate=""; string startStr=""; string endStr=""; string lableName="";

    templateHtml= getFText(cfg_webTemplate + "/" + templateName);
    //从栏目名称搜索，到栏目类型，到默认20160630
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


//加载就运行
void loadRun(){
    //这是为了给.net使用的，因为在.net里面全局变量不能有变量
    WEB_CACHEFile=replace(replace(WEB_CACHEFile,"[adminDir]",adminDir),"[EDITORTYPE]",EDITORTYPE);

    //缓冲处理20160622
    cacheHtmlFilePath= "/cache/html/" + setFileName(getThisUrlFileParam()) + ".html";
    //启用缓冲
    if( cStr(Request["cache"]) != "false" && isOnCacheHtml== true ){
        if( checkFile(cacheHtmlFilePath)== true ){
            //call echo("读取缓冲文件","OK")
            rwEnd(getFText(cacheHtmlFilePath));
        }
    }

    //记录表前缀
    if( cStr(Request["db_PREFIX"]) != "" ){
        db_PREFIX= cStr(Request["db_PREFIX"]);
    }else if( cStr(Session["db_PREFIX"]) != "" ){
        db_PREFIX= cStr(Session["db_PREFIX"]);
    }
    //加载网址配置
    loadWebConfig();
    isMakeHtml= false; //默认生成HTML为关闭
    if( cStr(Request["isMakeHtml"])== "1" || cStr(Request["isMakeHtml"])== "true" ){
        isMakeHtml= true;
    }
    templateName= cStr(Request["templateName"]); //模板名称

    //保存数据处理页
    switch ( cStr(Request["act"]) ){
        case "savedata" : saveData(cStr(Request["stype"])) ; Response.End(); //保存数据
        break;//'站长统计 | 今日IP[653] | 今日PV[9865] | 当前在线[65]')
        case "webstat" : webStat(adminDir + "/Data/Stat/") ; Response.End(); //网站统计
        break;
        case "saveSiteMap" : isMakeHtml= true ; saveSiteMap() ; Response.End(); //保存sitemap.xml
        break;
        case "handleAction":
        if( cStr(Request["ishtml"])== "1" ){
            isMakeHtml= true;
        }
        rwEnd(handleAction(cStr(Request["content"]))); //处理动作
        break;
    }


    //生成html
    if( cStr(Request["act"])== "makehtml" ){
        echo("makehtml", "makehtml");
        isMakeHtml= true;
        makeWebHtml(" action actionType='" + cStr(Request["act"]) + "' columnName='" + cStr(Request["columnName"]) + "' id='" + cStr(Request["id"]) + "' ");
        createFileGBK("index.html", code);

        //复制Html到网站
    }else if( cStr(Request["act"])== "copyHtmlToWeb" ){
        copyHtmlToWeb();
        //全部生成
    }else if( cStr(Request["act"])== "makeallhtml" ){
        makeAllHtml("", "", cStr(Request["id"]));

        //生成当前页面
    }else if( cStr(Request["isMakeHtml"]) != "" && cStr(Request["isSave"]) != "" ){

        handlePower("生成当前HTML页面"); //管理权限处理
        writeSystemLog("", "生成当前HTML页面"); //系统日志

        isMakeHtml= true;


        checkIDSQL(cStr(Request["id"]));
        rw(makeWebHtml(" action actionType='" + cStr(Request["act"]) + "' columnName='" + cStr(Request["columnName"]) + "' columnType='" + cStr(Request["columnType"]) + "' id='" + cStr(Request["id"]) + "' npage='" + cStr(Request["page"]) + "' "));
        glb_filePath= replace(glb_url, cfg_webSiteUrl, "");
        if( right(glb_filePath, 1)== "/" ){
            glb_filePath= glb_filePath + "index.html";
        }else if( glb_filePath== "" && glb_columnType== "首页" ){
            glb_filePath= "index.html";
        }
        //文件不为空  并且开启生成html
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
            echo("生成文件路径", "<a href=\"" + glb_filePath + "\" target='_blank'>" + glb_filePath + "</a>");

            //新闻则批量生成 20160216
            if( glb_columnType== "新闻" ){
                makeAllHtml("", "", glb_columnId);
            }

        }

        //全部生成
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
    //开启缓冲html
    if( isOnCacheHtml== true ){
        createFile(cacheHtmlFilePath, code); //保存到缓冲文件里20160622
    }
}
//检测ID是否SQL安全
bool checkIDSQL(string id){
    if( checkNumber(id)== false && id != "" ){
        eerr("提示", "id中有非法字符");
    }
    return false;
}
//http://127.0.0.1/aspweb.asp?act=nav&columnName=ASP
//http://127.0.0.1/aspweb.asp?act=detail&id=75
//生成html静态页
string makeWebHtml(string action){
    string actionType=""; int npagesize=0; int npage=0; string s=""; string url=""; string addSql=""; string sortSql=""; string sortFieldName=""; string ascOrDesc="";
    string serchKeyWordName=""; string parentid=""; //追加于20160716 home
    actionType= rParam(action, "actionType");
    s= rParam(action, "npage");
    s= getNumber(s);
    if( s== "" ){
        npage= 1;
    }else{
        npage= cInt(s);
    }
    //导航
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
            npagesize= cInt(rs["npagesize"]); //每页显示条数
            glb_isonhtml= IIF(cBool(rs["isonhtml"])== true, true, false); //是否生成静态网页
            sortSql= " " + cStr(rs["sortsql"]); //排序SQL

            if( cStr(rs["webtitle"]) != "" ){
                cfg_webTitle= cStr(rs["webtitle"]); //网址标题
            }
            if( cStr(rs["webkeywords"]) != "" ){
                cfg_webKeywords= cStr(rs["webkeywords"]); //网站关键词
            }
            if( cStr(rs["webdescription"]) != "" ){
                cfg_webDescription= cStr(rs["webdescription"]); //网站描述
            }
            if( templateName== "" ){
                if( aspTrim(cStr(rs["templatepath"])) != "" ){
                    templateName= cStr(rs["templatepath"]);
                }else if( cStr(rs["columntype"]) != "首页" ){
                    templateName= getDateilTemplate(cStr(rs["id"]), "List");
                }
            }
        }
        glb_columnENType= handleColumnType(glb_columnType);
        glb_url= getColumnUrl(glb_columnName, "name");

        //文章类列表
        if( inStr("|产品|新闻|视频|下载|案例|", "|" + glb_columnType + "|") > 0 ){
            glb_bodyContent= getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "栏目列表", "*", npagesize, cStr(npage), "where parentid=" + glb_columnId + sortSql);
            //留言类列表
        }else if( inStr("|留言|", "|" + glb_columnType + "|") > 0 ){
            glb_bodyContent= getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "GuestBook", "留言列表", "*", npagesize, cStr(npage), " where isthrough<>0 " + sortSql);
        }else if( glb_columnType== "文本" ){
            //航行栏目加管理
            if( cStr(Request["gl"])== "edit" ){
                glb_bodyContent= "<span>" + glb_bodyContent + "</span>";
            }
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" + glb_columnId + "&n=" + getRnd(11);
            glb_bodyContent= handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span");

        }
        //细节
    }else if( actionType== "detail" ){
        glb_locationType= "detail";
        rs = new OleDbCommand("Select * from " + db_PREFIX + "articledetail where id=" + rParam(action, "id"), conn).ExecuteReader();

        if( rs.Read() ){
            glb_columnName= getColumnName(cStr(rs["parentid"]));
            glb_detailTitle= cStr(rs["title"]);
            glb_flags= cStr(rs["flags"]);
            glb_isonhtml= cBool(rs["isonhtml"]); //是否生成静态网页
            glb_id= cStr(rs["id"]); //文章ID
            if( isMakeHtml== true ){
                glb_url= getHandleRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "/detail/detail" + cStr(rs["id"]));
            }else{
                glb_url= handleWebUrl("?act=detail&id=" + cStr(rs["id"]));
            }

            if( cStr(rs["webtitle"]) != "" ){
                cfg_webTitle= cStr(rs["webtitle"]); //网址标题
            }
            if( cStr(rs["webkeywords"]) != "" ){
                cfg_webKeywords= cStr(rs["webkeywords"]); //网站关键词
            }
            if( cStr(rs["webdescription"]) != "" ){
                cfg_webDescription= cStr(rs["webdescription"]); //网站描述
            }

            //改进20160628
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
            connexecute("update " + db_PREFIX + "articledetail set hits=hits+1 where id=" + cStr(rs["id"])); //更新点击数
            //文章内容
            //glb_bodyContent = "<div class=""articleinfowrap"">[$articleinfowrap$]</div>" & rs("bodycontent") & "[$relatedtags$]<ul class=""updownarticlewrap"">[$updownArticle$]</ul>"
            //上一篇文章，下一篇文章
            //glb_bodyContent = Replace(glb_bodyContent, "[$updownArticle$]", upArticle(rs("parentid"), "sortrank", rs("sortrank")) & downArticle(rs("parentid"), "sortrank", rs("sortrank")))
            //glb_bodyContent = Replace(glb_bodyContent, "[$articleinfowrap$]", "来源：" & rs("author") & " &nbsp; 发布时间：" & format_Time(rs("adddatetime"), 1))
            //glb_bodyContent = Replace(glb_bodyContent, "[$relatedtags$]", aritcleRelatedTags(rs("relatedtags")))

            glb_bodyContent= cStr(rs["bodycontent"]);

            //文章详细加控制
            if( cStr(Request["gl"])== "edit" ){
                glb_bodyContent= "<span>" + glb_bodyContent + "</span>";
            }
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=" + rParam(action, "id") + "&n=" + getRnd(11);
            glb_bodyContent= handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span");

            if( templateName== "" ){
                if( aspTrim(cStr(rs["templatepath"])) != "" ){
                    templateName= cStr(rs["templatepath"]);
                }else{
                    templateName= getDateilTemplate(cStr(rs["parentid"]), "Detail");
                }
            }

        }

        //单页
    }else if( actionType== "onepage" ){
        rs = new OleDbCommand("Select * from " + db_PREFIX + "onepage where id=" + rParam(action, "id"), conn).ExecuteReader();

        if( rs.Read() ){
            glb_detailTitle= cStr(rs["title"]);
            glb_isonhtml= cBool(rs["isonhtml"]); //是否生成静态网页
            if( isMakeHtml== true ){
                glb_url= getHandleRsUrl(cStr(rs["filename"]), cStr(rs["customaurl"]), "/page/page" + cStr(rs["id"]));
            }else{
                glb_url= handleWebUrl("?act=detail&id=" + cStr(rs["id"]));
            }

            if( cStr(rs["webtitle"]) != "" ){
                cfg_webTitle= cStr(rs["webtitle"]); //网址标题
            }
            if( cStr(rs["webkeywords"]) != "" ){
                cfg_webKeywords= cStr(rs["webkeywords"]); //网站关键词
            }
            if( cStr(rs["webdescription"]) != "" ){
                cfg_webDescription= cStr(rs["webdescription"]); //网站描述
            }
            //内容
            glb_bodyContent= cStr(rs["bodycontent"]);


            //文章详细加控制
            if( cStr(Request["gl"])== "edit" ){
                glb_bodyContent= "<span>" + glb_bodyContent + "</span>";
            }
            url= WEB_ADMINURL + "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=" + rParam(action, "id") + "&n=" + getRnd(11);
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

        //搜索
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
        glb_bodyContent= getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "网站栏目", "*", npagesize, cStr(npage), addSql);
        positionEndStr= " >> 搜索内容”" + glb_searchKeyWord + "“";
        //加载等待
    }else if( actionType== "loading" ){
        rwEnd("页面正在加载中。。。");
    }
    //模板为空，则用默认首页模板
    if( templateName== "" ){
        templateName= "Index_Model.html"; //默认模板
    }
    //检测当前路径是否有模板
    if( inStr(templateName, "/")== 0 ){
        templateName= cfg_webTemplate + "/" + templateName;
    }
    //call echo("templateName",templateName)
    if( checkFile(templateName)== false ){
        eerr("未找到模板文件", templateName);
    }
    code= getFText(templateName);

    code= handleAction(code); //处理动作
    code= thisPosition(code); //位置
    code= replaceGlobleVariable(code); //替换全局标签
    code= handleAction(code); //处理动作    '再来一次，处理数据内容里动作

    //call die(code)
    code= handleAction(code); //处理动作
    code= handleAction(code); //处理动作
    code= thisPosition(code); //位置
    code= replaceGlobleVariable(code); //替换全局标签
    code= delTemplateMyNote(code); //删除无用内容
    code= handleAction(code); //处理动作

    //格式化HTML
    if( inStr(cfg_flags, "|formattinghtml|") > 0 ){
        //code = HtmlFormatting(code)        '简单
        code= handleHtmlFormatting(code, false, 0, "删除空行"); //自定义
        //格式化HTML第二种
    }else if( inStr(cfg_flags, "|formattinghtmltow|") > 0 ){
        code= htmlFormatting(code); //简单
        code= handleHtmlFormatting(code, false, 0, "删除空行"); //自定义
        //压缩HTML
    }else if( inStr(cfg_flags, "|ziphtml|") > 0 ){
        code= zipHTML(code);

    }
    //闭合标签
    if( inStr(cfg_flags, "|labelclose|") > 0 ){
        code= handleCloseHtml(code, true, ""); //图片自动加alt  "|*|",
    }

    //在线编辑20160127
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

//获得默认细节模板页
string getDateilTemplate(string parentid, string templateType){
    string templateName="";
    templateName= "Main_Model.html";
    rsx = new OleDbCommand("select * from " + db_PREFIX + "webcolumn where id=" + parentid, conn).ExecuteReader();

    if( rsx.Read() ){
        //call echo("columntype",rsx("columntype"))
        if( cStr(rsx["columntype"])== "新闻" ){
            //新闻细节页
            if( checkFile(cfg_webTemplate + "/News_" + templateType + ".html")== true ){
                templateName= "News_" + templateType + ".html";
            }
        }else if( cStr(rsx["columntype"])== "产品" ){
            //产品细节页
            if( checkFile(cfg_webTemplate + "/Product_" + templateType + ".html")== true ){
                templateName= "Product_" + templateType + ".html";
            }
        }else if( cStr(rsx["columntype"])== "下载" ){
            //下载细节页
            if( checkFile(cfg_webTemplate + "/Down_" + templateType + ".html")== true ){
                templateName= "Down_" + templateType + ".html";
            }

        }else if( cStr(rsx["columntype"])== "视频" ){
            //视频细节页
            if( checkFile(cfg_webTemplate + "/Video_" + templateType + ".html")== true ){
                templateName= "Video_" + templateType + ".html";
            }
        }else if( cStr(rsx["columntype"])== "留言" ){
            //视频细节页
            if( checkFile(cfg_webTemplate + "/GuestBook_" + templateType + ".html")== true ){
                templateName= "Video_" + templateType + ".html";
            }
        }else if( cStr(rsx["columntype"])== "文本" ){
            //视频细节页
            if( checkFile(cfg_webTemplate + "/Page_" + templateType + ".html")== true ){
                templateName= "Page_" + templateType + ".html";
            }
        }
    }
    //call echo(templateType,templateName)
    string getDateilTemplate= templateName;

    return getDateilTemplate;
}


//生成全部html页面
void makeAllHtml(string columnType, string columnName, string columnId){
    string action=""; string s=""; int i=0; int nPageSize=0; int nCountSize=0; int nPage=0; string addSql=""; string url=""; string articleSql="";
    handlePower("生成全部HTML页面"); //管理权限处理
    writeSystemLog("", "生成全部HTML页面"); //系统日志

    isMakeHtml= true;
    //栏目
    echo("栏目", "");
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
        //开启生成html
        if( cBool(rss["isonhtml"])== true ){
            if( inStr("|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单|", "|" + cStr(rss["columntype"]) + "|") > 0 ){
                if( cStr(rss["columntype"])== "留言" ){
                    nCountSize= getRecordCount(db_PREFIX + "guestbook", ""); //记录数
                }else{
                    nCountSize= getRecordCount(db_PREFIX + "articledetail", " where parentid=" + cStr(rss["id"])); //记录数
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
                    templateName= ""; //清空模板文件名称
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
            connexecute("update " + db_PREFIX + "WebColumn set ishtml=true where id=" + cStr(rss["id"])); //更新导航为生成状态
        }
    }

    //单独处理指定栏目对应文章
    if( columnId != "" ){
        articleSql= "select * from " + db_PREFIX + "articledetail where parentid=" + columnId + " order by sortrank asc";
        //批量处理文章
    }else if( addSql== "" ){
        articleSql= "select * from " + db_PREFIX + "articledetail order by sortrank asc";
    }
    if( articleSql != "" ){
        //文章
        echo("文章", "");
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
            //文件不为空  并且开启生成html
            if( glb_filePath != "" && cBool(rss["isonhtml"])== true ){
                createDirFolder(getFileAttr(glb_filePath, "1"));
                createFileGBK(glb_filePath, code);
                connexecute("update " + db_PREFIX + "ArticleDetail set ishtml=true where id=" + cStr(rss["id"])); //更新文章为生成状态
            }
            templateName= ""; //清空模板文件名称
        }
    }

    if( addSql== "" ){
        //单页
        echo("单页", "");
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
            //文件不为空  并且开启生成html
            if( glb_filePath != "" && cBool(rss["isonhtml"])== true ){
                createDirFolder(getFileAttr(glb_filePath, "1"));
                createFileGBK(glb_filePath, code);
                connexecute("update " + db_PREFIX + "onepage set ishtml=true where id=" + cStr(rss["id"])); //更新单页为生成状态
            }
            templateName= ""; //清空模板文件名称
        }

    }
}

//复制html到网站
void copyHtmlToWeb(){
    string webDir=""; string toWebDir=""; string toFilePath=""; string filePath=""; string fileName=""; string fileList=""; string[] splStr; string content=""; string s=""; string s1=""; string c=""; string webImages=""; string webCss=""; string webJs=""; string[] splJs;
    string webFolderName=""; string jsFileList=""; string setFileCode=""; int nErrLevel=0; string jsFilePath=""; string url="";

    setFileCode= cStr(Request["setcode"]); //设置文件保存编码

    handlePower("复制生成HTML页面"); //管理权限处理
    writeSystemLog("", "复制生成HTML页面"); //系统日志

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

    deleteFolder(toWebDir); //删除
    createFolder("/htmlweb/web"); //创建文件夹 防止web文件夹不存在20160504
    deleteFolder(webDir);
    createDirFolder(webDir);
    webImages= webDir + "Images/";
    webCss= webDir + "Css/";
    webJs= webDir + "Js/";
    copyFolder(cfg_webImages, webImages);
    copyFolder(cfg_webCss, webCss);
    createFolder(webJs); //创建Js文件夹


    //处理Js文件夹
    splJs= aspSplit(getDirJsList(webJs), vbCrlf());
    foreach(var eachfilePath in splJs){
        filePath=eachfilePath;
        if( filePath != "" ){
            toFilePath= webJs + getFileName(filePath);
            echo("js", filePath);
            moveFile(filePath, toFilePath);
        }
    }
    //处理Css文件夹
    splStr= aspSplit(getDirCssList(webCss), vbCrlf());
    foreach(var eachfilePath in splStr){
        filePath=eachfilePath;
        if( filePath != "" ){
            content= getFText(filePath);
            content= replace(content, cfg_webImages, "../images/");

            content= deleteCssNote(content);
            content= phpTrim(content);
            //设置为utf-8编码 20160527
            if( lCase(setFileCode)== "utf-8" ){
                content= replace(content, "gb2312", "utf-8");
            }
            writeToFile(filePath, content, setFileCode);
            echo("css", cfg_webImages);
        }
    }
    //复制栏目HTML
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
            echo("导航", glb_filePath);
        }
    }
    //复制文章HTML
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
            echo("文章" + cStr(rss["title"]), glb_filePath);
        }
    }
    //复制单面HTML
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
            echo("单页" + cStr(rss["title"]), glb_filePath);
        }
    }
    //批量处理html文件列表
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
                //Call echo(sourceUrl, replaceUrl)                             '屏蔽  否则大量显示20160613
                content= replace(content, sourceUrl, replaceUrl);
            }
            content= replace(content, cfg_webSiteUrl, ""); //删除网址
            content= replace(content, cfg_webTemplate + "/", ""); //删除模板路径 记

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
            content= replace(content, "<a href=\"\" ", "<a href=\"index.html\" "); //让首页加index.html

            createFileGBK(filePath, content);
        }
    }

    //把复制网站夹下的images/文件夹下的js移到js/文件夹下  20160315
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
        content= handleHtmlFormatting(content, false, nErrLevel, "|删除空行|"); //|删除空行|
        content= handleCloseHtml(content, true, ""); //闭合标签
        //nErrLevel = checkHtmlFormatting(content)
        if( checkHtmlFormatting(content)== false ){
            echoRed(htmlFilePath + "(格式化错误)", nErrLevel); //注意
        }
        //设置为utf-8编码
        if( lCase(setFileCode)== "utf-8" ){
            content= replace(content, "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\" />", "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />");
        }
        content= phpTrim(content);
        writeToFile(htmlFilePath, content, setFileCode);
    }
    //images下js移动到js下
    foreach(var eachjsFileName in splJsFile){
        jsFileName=eachjsFileName;
        jsFilePath= webImages + jsFileName;
        content= getFText(jsFilePath);
        content= phpTrim(content);
        writeToFile(webJs + jsFileName, content, setFileCode);
        deleteFile(jsFilePath);
    }

    copyFolder(webDir, toWebDir);
    //使htmlWeb文件夹用php压缩
    if( cStr(Request["isMakeZip"])== "1" ){
        makeHtmlWebToZip(webDir);
    }
    //使网站用xml打包20160612
    if( cStr(Request["isMakeXml"])== "1" ){
        makeHtmlWebToXmlZip("/htmladmin/", webFolderName);
    }
    //浏览地址
    url= "http://10.10.10.57/" + toWebDir;
    echo("浏览", "<a href='" + url + "' target='_blank'>" + url + "</a>");
}
//使htmlWeb文件夹用php压缩
string makeHtmlWebToZip(string webDir){
    string content=""; string[] splStr; string filePath=""; string c=""; string[] arrayFile; string fileName=""; string fileType=""; bool isTrue;
    string webFolderName="";
    string cleanFileList="";
    splStr= aspSplit(webDir, "/");
    webFolderName= splStr[2];
    //content = getFileFolderList(webDir, true, "全部", "", "全部文件夹", "", "") 		'屏蔽这种
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
    //先判断这个文件存在20160309
    if( checkFile("/myZIP.php")== true ){
        echo("", xmlPost(getHost() + "/myZIP.php?webFolderName=" + webFolderName, "content=" + escape(c)));
    }

    return "";
}
//使网站用xml打包20160612
string makeHtmlWebToXmlZip(string sNewWebDir,string rootDir){//留空函数
    return "";
}


//生成更新sitemap.xml 20160118
void saveSiteMap(){
    bool isWebRunHtml; //是否为html方式显示网站
    string changefreg=""; //更新频率
    string priority=""; //优先级
    string s=""; string c=""; string url="";string sql="";
    handlePower("修改生成SiteMap"); //管理权限处理

    changefreg= cStr(Request["changefreg"]);
    priority= cStr(Request["priority"]);
    loadWebConfig(); //加载配置
    //call eerr("cfg_flags",cfg_flags)
    if( inStr(cfg_flags, "|htmlrun|") > 0 ){
        isWebRunHtml= true;
    }else{
        isWebRunHtml= false;
    }

    c= c + "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + vbCrlf();
    c= c + vbTab() + "<urlset xmlns=\"http://www.sitemaps.org/schemas/sitemap/0.9\">" + vbCrlf();

    //栏目
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
            echo("栏目", "<a href=\"" + url + "\" target='_blank'>" + url + "</a>");
        }
    }

    //文章
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
            echo("文章", "<a href=\"" + url + "\">" + url + "</a>");
        }
    }

    //单页
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
            echo("单页", "<a href=\"" + url + "\">" + url + "</a>");
        }
    }
    c= c + vbTab() + "</urlset>" + vbCrlf();
    loadWebConfig();
    createFile("sitemap.xml", c);
    echo("生成sitemap.xml文件成功", "<a href='/sitemap.xml' target='_blank'>点击预览sitemap.xml</a>");


    //判断是否生成sitemap.html
    if( cStr(Request["issitemaphtml"])== "1" ){
        c= "";
        //第二种
        //栏目
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

                //判断是否生成html
                if( cBool(rsx["isonhtml"])== true ){
                    s= "<a href=\"" + url + "\">" + cStr(rsx["columnname"]) + "</a>";
                }else{
                    s= "<span>" + cStr(rsx["columnname"]) + "</span>";
                }
                c= c + "<li style=\"width:20%;\">" + s + vbCrlf() + "<ul>" + vbCrlf();

                //文章
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
                        //判断是否生成html
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

        //单面
        c= c + "<li style=\"width:20%;\"><a href=\"javascript:;\">单面列表</a>" + vbCrlf() + "<ul>" + vbCrlf();
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
                //判断是否生成html
                if( cBool(rsx["isonhtml"])== true ){
                    s= "<a href=\"" + url + "\">" + cStr(rsx["title"]) + "</a>";
                }else{
                    s= "<span>" + cStr(rsx["title"]) + "</span>";
                }

                c= c + "<li style=\"width:20%;\">" + s + "</li>" + vbCrlf(); //target=""_blank""  去掉
            }
        }
        c= c + "</ul>" + vbCrlf() + "</li>" + vbCrlf();

        string templateContent="";
        templateContent= getFText(adminDir + "/template_SiteMap.html");


        templateContent= replace(templateContent, "{$content$}", c);
        templateContent= replace(templateContent, "{$Web_Title$}", cfg_webTitle);


        createFile("sitemap.html", templateContent);
        echo("生成sitemap.html文件成功", "<a href='/sitemap.html' target='_blank'>点击预览sitemap.html</a>");
    }
    writeSystemLog("", "保存sitemap.xml"); //系统日志
}
</script>
