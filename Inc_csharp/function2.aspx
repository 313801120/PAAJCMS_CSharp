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
<script runat="server" language="c#">


//����function2�ļ�����
string callFunction2(){
    switch ( cStr(Request["stype"]) ){
        case "runScanWebUrl" : runScanWebUrl() ;break;//����ɨ����ַ
        case "scanCheckDomain" : scanCheckDomain() ;break;//���������Ч
        case "bantchImportDomain" : bantchImportDomain() ;break;//������������
        case "scanDomainHomePage" : scanDomainHomePage()								;break;//ɨ��������ҳ
        case "scanDomainHomePageSize" : scanDomainHomePageSize()								;break;//ɨ��������ҳ��С�����
        case "isthroughTrue" : isthroughTrue()											;break;//�����ȫ��Ϊ��
        case "printOKWebSite" : printOKWebSite()										;break;//��ӡ��Ч��ַ
        case "printAspServerWebSite" : printAspServerWebSite();										//��ӡasp������վ
        break;
        case "clearAllData" : fun2_clearAllData();										//���ȫ������
        break;
        case "function2test" : function2test()											;break;//����
        default : eerr("function2ҳ��û�ж���", cStr(Request["stype"]));
        break;
    }
    return "";
}

//����
string function2test(){
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where isdomain=true", conn).ExecuteReader();

    echo("��",rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where isdomain=true"));
    while( rs.Read()){
        echo(cStr(rs["isdomain"]),cStr(rs["website"]));
    }
    return "";
}
//���ȫ������
string fun2_clearAllData(){
    conn=openConn();
    conn.Open();

    connexecute("delete from " + db_PREFIX + "webdomain");
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����'>OK</a>");
    return "";
}
//��ӡ��Ч��ַ
string printOKWebSite(){
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where isdomain=true", conn).ExecuteReader();

    echo("��",rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where isdomain=true"));
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����'>OK</a>");
    while( rs.Read()){
        //call echo(rs("isdomain"),rs("website"))
        rw(cStr(rs["website"]) + "<br>");
    }
    return "";
}
//��ӡasp������վ
string printAspServerWebSite(){
    conn=openConn();
    conn.Open();

    rs = new OleDbCommand("select * from " + db_PREFIX + "webdomain where isasp=true and (isaspx=false and isphp=false)", conn).ExecuteReader();

    echo("��",rsRecordcount("select count(*) from " + db_PREFIX + "webdomain where isasp=true and (isaspx=false and isphp=false)"));
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����'>OK</a>");
    while( rs.Read()){
        //call echo(rs("isdomain"),rs("website"))
        rw(cStr(rs["website"]) + "<br>");
    }
    return "";
}

//�����ȫ��Ϊ��
bool isthroughTrue(){
    conn=openConn();
    conn.Open();

    connexecute("update " + db_PREFIX + "webdomain set isthrough=true");
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����'>OK</a>");
    return false;
}

//ɨ����ҳ��С
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
        htmlDir= "/../��վUrlScan/������ҳ��С/";
        createDirFolder(htmlDir);
        txtFilePath= htmlDir + "/" + setFileName(cStr(rs["website"])) + ".txt";
        if( checkFile(txtFilePath)== true ){
            echo("����", "����");
            nSetTime=1;
        }else{
            website=getWebSite(cStr(rs["website"]));
            if( website=="" ){
                eerr("����Ϊ��",website);
            }
            content=getHttpPage(website,cStr(rs["charset"]));

            if( content=="" ){
                content=" ";
            }

            createFile(txtFilePath, content);
            echo("����", "����");
        }
        content=getFText(txtFilePath);
        webtitle=getHtmlValue(content,"webtitle");
        webkeywords=getHtmlValue(content,"webkeywords");
        webdescription=getHtmlValue(content,"webdescription");


        nWebSize=getFSize(txtFilePath);
        echo("webtitle",webtitle);
        //����д�Ǹ�תPHPʱ����
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
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����'>OK����("+ nThis +")��</a>");
    return "";
}

//ɨ��������ҳ
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
        htmlDir= "/../��վUrlScan/������ҳ/";
        createDirFolder(htmlDir);
        txtFilePath= htmlDir + "/" + setFileName(cStr(rs["website"])) + ".txt";
        if( checkFile(txtFilePath)== true ){
            c= phpTrim(getFText(txtFilePath));
            nIsAsp=cInt(getStrCut(c,"isAsp=",vbCrlf(),1));
            nIsAspx=cInt(getStrCut(c,"isAspx=",vbCrlf(),1));
            nIsPhp=cInt(getStrCut(c,"isPhp=",vbCrlf(),1));
            nIsJsp=cInt(getStrCut(c,"isJsp=",vbCrlf(),1));
            echo("����", "����");
            nSetTime=1;
        }else{
            website=getWebSite(cStr(rs["website"]));
            if( website=="" ){
                eerr("����Ϊ��",website);
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
                homePageList="��";
            }

            createFile(txtFilePath, c);
            echo("����", "����");
        }
        //����д�Ǹ�תPHPʱ����
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
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����'>OK����("+ nThis +")��</a>");
    return "";
}

//������������
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
                echo("��ӳɹ�", webSite);
                nOK= nOK + 1;
            }else{
                echo("website", webSite);
            }
        }
    }
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����'>OK ��(" + nOK + ")��</a>");
    return "";
}

//���������Ч
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
        htmlDir= "/../��վUrlScan/����/";
        createDirFolder(htmlDir);
        txtFilePath= htmlDir + "/" + setFileName(cStr(rs["website"])) + ".txt";
        if( checkFile(txtFilePath)== true ){
            nIsDomain= cInt(phpTrim(getFText(txtFilePath)));
            echo("����", "����");
            nSetTime=1;
        }else{
            nIsDomain= IIF(checkDomainName(cStr(rs["website"])), 1, 0);
            createFile(txtFilePath, nIsDomain + " ");			 //��ֹPHP��д�벻��ȥ 0 �������
            echo("����", "����" + txtFilePath + "("+ checkFile(txtFilePath) +")=" + nIsDomain);
        }
        //����д�Ǹ�תPHPʱ����
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
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebDomain&addsql=order by id desc&lableTitle=��վ����'>OK����("+ nThis +")��</a>");
    return "";
}

//ɨ����ַ
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
    //ѭ��
    rsx = new OleDbCommand("select * from " + db_PREFIX + "weburlscan where isThrough=true", conn).ExecuteReader();

    if( rsx.Read() ){
        nThis=nThis+1;
        echo(nThis, cStr(rsx["httpurl"]));
        doEvents();
        nSetTime= scanUrl(cStr(rsx["httpurl"]), cStr(rsx["title"]), cStr(rsx["charset"]));
        //����д�Ǹ�תPHPʱ����
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
    echo("�������", "<a href='?act=dispalyManageHandle&actionType=WebUrlScan&addsql=order by id desc&lableTitle=��ַɨ��'>OK����("+ nThis +")��</a>");
    //���뱨��
    rs = new OleDbCommand("select * from " + db_PREFIX + "weburlscan where webstate=404", conn).ExecuteReader();

    while( rs.Read()){
        echo("<a href='" + cStr(rs["httpurl"]) + "' target='_blank'>" + cStr(rs["httpurl"]) + "</a>", "<a href='" + cStr(rs["tohttpurl"]) + "' target='_blank'>" + cStr(rs["tohttpurl"]) + "</a>");
    }
    return "";
}
//ɨ����ַ
int scanUrl(string httpUrl, string toTitle, string codeset){
    string[] splStr; int i=0; string s=""; string content=""; string PubAHrefList=""; string PubATitleList=""; string[] splUrl; string[] spltitle; string title=""; string url=""; string htmlDir=""; string htmlFilePath=""; int nOK=0; string[] arrayData; int nWebState=0; string u=""; string iniDir=""; string iniFilePath="" ;int nWebSize=0;
    int nSetTime=0; System.DateTime startTime; int nOpenSpeed=0; bool isLocal; int nIsThrough=0;
    htmlDir= "/../��վUrlScan/" + setFileName(getWebSite(httpUrl));
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

    //strLen(content)  �������������׼

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
        //ȥ��#�ź�̨���ַ�20160506
        if( inStr(url, "#") > 0 ){
            url= mid(url, 1, inStr(url, "#") - 1);
        }
        if( url== "" ){
            if( title != "" ){
                echo("��ַΪ��", title);
            }
        }else{
            url= handleScanUrlList(httpUrl, url);
            url= handleWithWebSiteList(httpUrl, url);
            if( url != "" ){
                rs = new OleDbCommand("select * from " + db_PREFIX + "weburlscan where httpurl='" + url + "'", conn).ExecuteReader();

                if( rs.Read() ){
                    u= lCase(url);
                    if( inStr(u, "tools/downfile.asp") > 0 || inStr(u, "/url.asp?") > 0 || inStr(u, "/aspweb.asp?") > 0 || inStr(u, "/phpweb.php?") > 0 || u== "http://www.maiside.net/qq/" || inStr(u, "mailto:") > 0 || inStr(u, "tel:") > 0 || inStr(u, ".html?replytocom") > 0 ){//.html?replytocom  ��ͨ��վ
                        nIsThrough= 0;
                    }else{
                        nIsThrough= 1; //����true ��Ϊд�����ݻ�������
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
