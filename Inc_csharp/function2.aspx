<%
/************************************************************
作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
版权：源代码免费公开，各种用途均可使用。 
创建：2016-08-04
联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
*                                    Powered by PAAJCMS 
************************************************************/
%>
<script runat="server" language="c#">


//调用function2文件函数
string callFunction2(){
    switch ( cStr(Request["stype"]) ){
        case "runScanWebUrl" : runScanWebUrl() ;break;//运行扫描网址
        case "scanCheckDomain" : scanCheckDomain() ;break;//检测域名有效
        case "bantchImportDomain" : bantchImportDomain() ;break;//批量导入域名
        case "scanDomainHomePage" : scanDomainHomePage()								;break;//扫描域名首页
        case "scanDomainHomePageSize" : scanDomainHomePageSize()								;break;//扫描域名首页大小与标题
        case "isthroughTrue" : isthroughTrue()											;break;//让审核全部为真
        case "printOKWebSite" : printOKWebSite()										;break;//打印有效网址
        case "printAspServerWebSite" : printAspServerWebSite();										//打印asp类型网站
        break;
        case "clearAllData" : fun2_clearAllData();										//清除全部数据
        break;
        case "function2test" : function2test()											;break;//测试
        default : eerr("function2页里没有动作", cStr(Request["stype"]));
        break;
    }
    return "";
}

//测试
string function2test(){
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where isdomain=true", conn).ExecuteReader();

    echo("共",rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where isdomain=true"));
    while( rs.Read()){
        echo(cStr(rs["isdomain"]),cStr(rs["website"]));
    }
    return "";
}
//清除全部数据
string fun2_clearAllData(){
    conn=openConn();
    conn.Open();

    connexecute("delete from " + db_PREFIX + "webdomain");
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK</a>");
    return "";
}
//打印有效网址
string printOKWebSite(){
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where isdomain=true", conn).ExecuteReader();

    echo("共",rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where isdomain=true"));
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK</a>");
    while( rs.Read()){
        //call echo(rs("isdomain"),rs("website"))
        rw(cStr(rs["website"]) + "<br>");
    }
    return "";
}
//打印asp类型网站
string printAspServerWebSite(){
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where isasp=true and (isaspx=false and isphp=false)", conn).ExecuteReader();

    echo("共",rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where isasp=true and (isaspx=false and isphp=false)"));
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK</a>");
    while( rs.Read()){
        //call echo(rs("isdomain"),rs("website"))
        rw(cStr(rs["website"]) + "<br>");
    }
    return "";
}

//让审核全部为真
bool isthroughTrue(){
    conn=openConn();
    conn.Open();

    connexecute("update " + db_PREFIX + "webdomain set isthrough=true");
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK</a>");
    return false;
}

//扫描首页大小
string scanDomainHomePageSize(){
    string url=""; int nSetTime=0; bool isdomain; string htmlDir=""; string txtFilePath="";string homePageList="";int nThis=0;int nCount=0;
    string[] splstr;string s="";string c="";string website="";int nState=0;int nWebSize=0;string content="";System.DateTime startTime;string webtitle="";string webkeywords="";string webdescription="";

    if( cStr(Request["nThis"])=="" ){
        nThis=0;
    }else{
        nThis=cInt(Request["nThis"]);
    }

    nSetTime= 3;
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where website<>'' and websize=0 and isdomain=true", conn).ExecuteReader();


    if( cStr(Request["nCount"])=="" ){
        nCount=rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where website<>'' and websize=0 and isdomain=true");
    }else{
        nCount=cInt(Request["nCount"]);
    }
    while( rs.Read()){
        nThis=nThis+1;
        echo(nThis + "/" + nCount, cStr(rs["website"]));
        doEvents();
        htmlDir= "/../网站UrlScan/域名首页大小/";
        createDirFolder(htmlDir);
        txtFilePath= htmlDir + "/" + setFileName(cStr(rs["website"])) + ".txt";
        if( checkFile(txtFilePath)== true ){
            echo("类型", "本地");
            nSetTime=1;
        }else{
            website=getWebSite(cStr(rs["website"]));
            if( website=="" ){
                eerr("域名为空",website);
            }
            content=getHttpPage(website,cStr(rs["charset"]));

            if( content=="" ){
                content=" ";
            }

            createFile(txtFilePath, content);
            echo("类型", "网络");
        }
        content=getFText(txtFilePath);
        webtitle=getHtmlValue(content,"webtitle");
        webkeywords=getHtmlValue(content,"webkeywords");
        webdescription=getHtmlValue(content,"webdescription");


        nWebSize=getFSize(txtFilePath);
        echo("webtitle",webtitle);
        //这样写是给转PHP时方便
        connexecute("update " + db_PREFIX + "webdomain  set webtitle='"+ ADSql(webtitle) +"',webkeywords='"+ webkeywords +"',webdescription='"+ webdescription +"',websize="+ nWebSize +",isthrough=false,updatetime='" + now() + "'  where id=" + cStr(rs["id"]) + "");


        if( cStr(Request["startTime"])=="" ){
            startTime=now();
        }else{
            startTime=cTime(Request["startTime"]);
        }

        rw(vbRunTimer(startTime) + "<hr>");
        url= getUrlAddToParam(getThisUrl(), "?nThis="+ nThis +"&nCount="+ nCount +"&startTime="+ startTime +"&N=" + getRnd(11), "replace");

        rw(jsTiming(url, nSetTime));
        Response.End();
    }
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK，共("+ nThis +")条</a>");
    return "";
}

//扫描域名首页
string scanDomainHomePage(){
    string url=""; int nSetTime=0; bool isdomain; string htmlDir=""; string txtFilePath="";string homePageList="";int nThis=0;int nCount=0;
    string[] splstr;string s="";string c="";string website="";int nState=0;System.DateTime startTime;
    int nIsAsp=0;int nIsAspx=0;int nIsPhp=0;int nIsJsp=0;string c2="";
    nIsAsp=0;nIsAspx=0;nIsPhp=0;nIsJsp=0;

    if( cStr(Request["nThis"])=="" ){
        nThis=0;
    }else{
        nThis=cInt(Request["nThis"]);
    }

    nSetTime= 3;
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where website<>'' and homepagelist='' and isdomain=true", conn).ExecuteReader();


    if( cStr(Request["nCount"])=="" ){
        nCount=rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where website<>'' and homepagelist='' and isdomain=true");
    }else{
        nCount=cInt(Request["nCount"]);
    }
    while( rs.Read()){
        nThis=nThis+1;
        echo(nThis + "/" + nCount, cStr(rs["website"]));
        doEvents();
        htmlDir= "/../网站UrlScan/域名首页/";
        createDirFolder(htmlDir);
        txtFilePath= htmlDir + "/" + setFileName(cStr(rs["website"])) + ".txt";
        if( checkFile(txtFilePath)== true ){
            c= phpTrim(getFText(txtFilePath));
            nIsAsp=cInt(getStrCut(c,"isAsp=",vbCrlf(),1));
            nIsAspx=cInt(getStrCut(c,"isAspx=",vbCrlf(),1));
            nIsPhp=cInt(getStrCut(c,"isPhp=",vbCrlf(),1));
            nIsJsp=cInt(getStrCut(c,"isJsp=",vbCrlf(),1));
            echo("类型", "本地");
            nSetTime=1;
        }else{
            website=getWebSite(cStr(rs["website"]));
            if( website=="" ){
                eerr("域名为空",website);
            }
            splstr=aspSplit("index.asp|index.aspx|index.php|index.jsp|index.htm|index.html|default.asp|default.aspx|default.jsp|default.htm|default.html","|");
            c2="";
            homePageList="";
            foreach(var eachs in splstr){
                s=eachs;
                url=website + s;
                nState=getHttpUrlState(url);
                echo(url,nState + "   ("+ getHttpUrlStateAbout(nState) +")");
                doEvents();
                if( (s=="index.asp" || s=="default.asp") && (nState==200 || nState==302) ){
                    nIsAsp=1;
                }else if( (s=="index.aspx" || s=="default.aspx") && (nState==200 || nState==302) ){
                    nIsAspx=1;
                }else if( (s=="index.php" || s=="default.php") && (nState==200 || nState==302) ){
                    nIsPhp=1;
                }else if( (s=="index.jsp" || s=="default.jsp") && (nState==200 || nState==302) ){
                    nIsJsp=1;
                }
                if( nState==200 || nState==302 ){
                    homePageList=homePageList + s + "|";
                }
                c2=c2 + s + "=" + nState + vbCrlf();
            }
            c= "isAsp=" + nIsAsp + vbCrlf();
            c= c + "isAspx=" + nIsAspx + vbCrlf();
            c= c + "isPhp=" + nIsPhp + vbCrlf();
            c= c + "isJsp=" + nIsJsp + vbCrlf() + c2;

            if( homePageList=="" ){
                homePageList="无";
            }

            createFile(txtFilePath, c);
            echo("类型", "网络");
        }
        //这样写是给转PHP时方便
        connexecute("update " + db_PREFIX + "webdomain  set isasp="+ nIsAsp +",isaspx="+ nIsAspx +",isphp="+ nIsPhp +",isjsp="+ nIsJsp +",isthrough=false,homepagelist='"+ homePageList +"',updatetime='" + now() + "'  where id=" + cStr(rs["id"]) + "");

        if( cStr(Request["startTime"])=="" ){
            startTime=now();
        }else{
            startTime=cTime(Request["startTime"]);
        }

        rw(vbRunTimer(startTime) + "<hr>");
        url= getUrlAddToParam(getThisUrl(), "?nThis="+ nThis +"&nCount="+ nCount +"&startTime="+ startTime +"&N=" + getRnd(11), "replace");

        rw(jsTiming(url, nSetTime));
        Response.End();
    }
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK，共("+ nThis +")条</a>");
    return "";
}

//批量导入域名
string bantchImportDomain(){
    string content=""; string[] splStr; string url=""; string webSite=""; int nOK=0;
    content= lCase(cStr(Request.Form["bodycontent"]));
    splStr= aspSplit(content, vbCrlf());
    nOK= 0;
    conn=openConn();
    conn.Open();

    foreach(var eachurl in splStr){
        url=eachurl;
        webSite= getWebSite(url);
        if( webSite != "" ){
            rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where website='" + webSite + "'", conn).ExecuteReader();

            if( rs.Read() ){
                connexecute("insert into " + db_PREFIX + "webdomain(website,isthrough,isdomain) values('" + webSite + "',true,false)");
                echo("添加成功", webSite);
                nOK= nOK + 1;
            }else{
                echo("website", webSite);
            }
        }
    }
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK 共(" + nOK + ")条</a>");
    return "";
}

//检测域名有效
string scanCheckDomain(){
    string url=""; int nSetTime=0; int nIsDomain=0; string htmlDir=""; string txtFilePath=""; int nThis=0;int nCount=0;System.DateTime startTime;
    nSetTime= 3;
    if( cStr(Request["nThis"])=="" ){
        nThis=0;
    }else{
        nThis=cInt(Request["nThis"]);
    }
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where isthrough=true", conn).ExecuteReader();


    if( cStr(Request["nCount"])=="" ){
        nCount=rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where isthrough=true");
    }else{
        nCount=cInt(Request["nCount"]);
    }
    while( rs.Read()){
        nThis=nThis+1;
        echo(nThis + "/" + nCount, cStr(rs["website"]));
        doEvents();
        htmlDir= "/../网站UrlScan/域名/";
        createDirFolder(htmlDir);
        txtFilePath= htmlDir + "/" + setFileName(cStr(rs["website"])) + ".txt";
        if( checkFile(txtFilePath)== true ){
            nIsDomain= cInt(phpTrim(getFText(txtFilePath)));
            echo("类型", "本地");
            nSetTime=1;
        }else{
            nIsDomain= IIF(checkDomainName(cStr(rs["website"])), 1, 0);
            createFile(txtFilePath, nIsDomain + " ");			 //防止PHP版写入不进去 0 这个内容
            echo("类型", "网络" + txtFilePath + "("+ checkFile(txtFilePath) +")=" + nIsDomain);
        }
        //这样写是给转PHP时方便
        connexecute("update " + db_PREFIX + "webdomain  set isthrough=false,isdomain=" + nIsDomain + ",updatetime='" + now() + "'  where id=" + cStr(rs["id"]) + "");

        if( cStr(Request["startTime"])=="" ){
            startTime=now();
        }else{
            startTime=cTime(Request["startTime"]);
        }

        rw(vbRunTimer(startTime) + "<hr>");
        url= getUrlAddToParam(getThisUrl(), "?nThis="+ nThis +"&nCount="+ nCount +"&startTime="+ startTime +"&N=" + getRnd(11), "replace");

        rw(jsTiming(url, nSetTime));
        Response.End();
    }
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=网站域名'>OK，共("+ nThis +")条</a>");
    return "";
}

//扫描网址
string runScanWebUrl(){
    int nSetTime=0; string setCharSet=""; string httpUrl=""; string url=""; string selectWeb="" ;int nThis=0;int nCount=0;System.DateTime startTime;
    setCharSet= "gb2312"; //gb2312
    //http://www.dfz9.com/
    //http://www.maiside.net/
    //http://sharembweb.com/
    //http://www.ufoer.com/
    httpUrl= "http://sharembweb.com/";
    //selectWeb="ufoer"
    if( selectWeb== "ufoer" ){
        httpUrl= "http://www.ufoer.com/";
        setCharSet= "utf-8";
    }

    if( cStr(Request["nThis"])=="" ){
        nThis=0;
    }else{
        nThis=cInt(Request["nThis"]);
    }

    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "weburlscan", conn).ExecuteReader();


    if( cStr(Request["nCount"])=="" ){
        nCount=rsRecordcount("select count(*) from " + db_PREFIX + "weburlscan");
    }else{
        nCount=cInt(Request["nCount"]);
    }
    if( rs.Read() ){
        connexecute("insert into " + db_PREFIX + "weburlscan(httpurl,title,isthrough,charset) values('" + httpUrl + "','home',true,'" + setCharSet + "')");
    }
    //循环
    rsx = new OleDbCommand("select * from " + db_PREFIX + "weburlscan where isThrough=true", conn).ExecuteReader();

    if( rsx.Read() ){
        nThis=nThis+1;
        echo(nThis, cStr(rsx["httpurl"]));
        doEvents();
        nSetTime= scanUrl(cStr(rsx["httpurl"]), cStr(rsx["title"]), cStr(rsx["charset"]));
        //这样写是给转PHP时方便
        connexecute("update " + db_PREFIX + "weburlscan  set isthrough=false  where id=" + cStr(rsx["id"]) + "");


        if( cStr(Request["startTime"])=="" ){
            startTime=now();
        }else{
            startTime=cTime(Request["startTime"]);
        }

        vbRunTimer(startTime);
        url= getUrlAddToParam(getThisUrl(), "?nThis="+ nThis +"&nCount="+ nCount +"&startTime="+ startTime +"&N=" + getRnd(11), "replace");

        rw(jsTiming(url, nSetTime));
        Response.End();
    }
    echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=WebUrlScan&addsql=order by id desc&lableTitle=网址扫描'>OK，共("+ nThis +")条</a>");
    //输入报告
    rs = new OleDbCommand("select * from " + db_PREFIX + "weburlscan where webstate=404", conn).ExecuteReader();

    while( rs.Read()){
        echo("<a href='" + cStr(rs["httpurl"]) + "' target='_blank'>" + cStr(rs["httpurl"]) + "</a>", "<a href='" + cStr(rs["tohttpurl"]) + "' target='_blank'>" + cStr(rs["tohttpurl"]) + "</a>");
    }
    return "";
}
//扫描网址
int scanUrl(string httpUrl, string toTitle, string codeset){
    string[] splStr; int i=0; string s=""; string content=""; string PubAHrefList=""; string PubATitleList=""; string[] splUrl; string[] spltitle; string title=""; string url=""; string htmlDir=""; string htmlFilePath=""; int nOK=0; string[] arrayData; int nWebState=0; string u=""; string iniDir=""; string iniFilePath="" ;int nWebSize=0;
    int nSetTime=0; System.DateTime startTime; int nOpenSpeed=0; bool isLocal; int nIsThrough=0;
    htmlDir= "/../网站UrlScan/" + setFileName(getWebSite(httpUrl));
    createDirFolder(htmlDir);
    htmlFilePath= htmlDir + "/" + setFileName(httpUrl) + ".html";
    iniDir= htmlDir + "/conifg";
    createFolder(iniDir);
    iniFilePath= iniDir + "/" + setFileName(httpUrl) + ".txt";

    //httpurl="http://maiside.net/"

    nWebState= 0;
    nSetTime= 1;
    nOpenSpeed= 0;
    if( checkFile(htmlFilePath)== false ){
        startTime= now();
        echo("codeset", codeset);
        arrayData= handleXmlGet(httpUrl, codeset);
        content= arrayData[0];


        nWebState= cInt(arrayData[1]);
        nOpenSpeed= dateDiff("s", startTime, now());
        //content=gethttpurl(httpurl,codeset)
        //call createfile(htmlFilePath,content)
        writeToFile(htmlFilePath, content, codeset);
        createFile(iniFilePath, nWebState + vbCrlf() + nOpenSpeed);
        nSetTime= 3;
        isLocal= false;
    }else{
        //content=getftext(htmlFilePath)
        content= reaFile(htmlFilePath, codeset);

        splStr= aspSplit(getFText(iniFilePath), vbCrlf());
        nWebState= cInt(splStr[0]);
        nOpenSpeed= cInt(splStr[0]);
        isLocal= true;
    }
    nWebSize=getFSize(htmlFilePath);
    echo("isLocal", isLocal);
    rs = new OleDbCommand("select * from " + db_PREFIX + "weburlscan where httpurl='" + httpUrl + "'", conn).ExecuteReader();

    if( rs.Read() ){
        connexecute("insert into " + db_PREFIX + "weburlscan(httpurl,title,charset) values('" + httpUrl + "','" + toTitle + "','" + codeset + "')");
    }
    connexecute("update " + db_PREFIX + "weburlscan  set webstate=" + nWebState + ",websize=" + nWebSize + ",openspeed=" + nOpenSpeed + ",charset='" + codeset + "'  where httpurl='" + httpUrl + "'");

    //strLen(content)  不用这个，不精准

    s= getContentAHref("", content, PubAHrefList, PubATitleList);
    s= handleScanUrlList(httpUrl, s);

    //call echo("httpurl",httpurl)
    //call echo("s",s)
    //call echo("PubATitleList",PubATitleList)
    nOK= 0;
    splUrl= aspSplit(PubAHrefList, vbCrlf());
    spltitle= aspSplit(PubATitleList, vbCrlf());
    for( i= 1 ; i<= uBound(splUrl); i++){
        title= spltitle[i];
        url= splUrl[i];
        //去掉#号后台的字符20160506
        if( inStr(url, "#") > 0 ){
            url= mid(url, 1, inStr(url, "#") - 1);
        }
        if( url== "" ){
            if( title != "" ){
                echo("网址为空", title);
            }
        }else{
            url= handleScanUrlList(httpUrl, url);
            url= handleWithWebSiteList(httpUrl, url);
            if( url != "" ){
                rs = new OleDbCommand("select * from " + db_PREFIX + "weburlscan where httpurl='" + url + "'", conn).ExecuteReader();

                if( rs.Read() ){
                    u= lCase(url);
                    if( inStr(u, "tools/downfile.asp") > 0 || inStr(u, "/url.asp?") > 0 || inStr(u, "/aspweb.asp?") > 0 || inStr(u, "/phpweb.php?") > 0 || u== "http://www.maiside.net/qq/" || inStr(u, "mailto:") > 0 || inStr(u, "tel:") > 0 || inStr(u, ".html?replytocom") > 0 ){//.html?replytocom  王通网站
                        nIsThrough= 0;
                    }else{
                        nIsThrough= 1; //不用true 因为写入数据会有问题
                    }
                    connexecute("insert into " + db_PREFIX + "weburlscan(tohttpurl,totitle,httpurl,title,isthrough,charset) values('" + httpUrl + "','" + toTitle + "','" + url + "','" + left(title, 255) + "'," + nIsThrough + ",'" + codeset + "')");
                    nOK= nOK + 1;
                    echo(i, url);
                }else{
                    echo(title, url);
                }
            }
        }
    }

    return nSetTime;
}


</script>
