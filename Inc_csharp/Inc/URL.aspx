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
//URL ��ַ���� (2013,9,27)

//byref PubAHrefList, byref PubATitleList   ΪΪʲôҪ����byref  ����Ϊ��תPHPʱ��Ҫ����жϣ��������ǵ�ַ���� �ǹ�������

//NOVBNet start
//ʹ���ֲ�
//call echo("'��õ�ǰ��ַ��һ�֣�getUrl��",getUrl())
//call echo("'��õ�ǰ��ַ�ڶ��֣�getThisUrl��",getThisUrl())
//call echo("'��õ�ǰ��ַ�޲�����getThisUrlNoParam��",getThisUrlNoParam())
//call echo("'���������ַ��getGoToUrl��",getGoToUrl())
//call echo("'���������ַ�޲�����getGoToUrlNoParam��",getGoToUrlNoParam())
//call echo("'���������ַ���ļ����ƣ�getGoToUrlNoFileName��",getGoToUrlNoFileName())
//call echo("'������webDoMain��",webDoMain())
//call echo("'������host��",host())
//call echo("'��õ�ǰ��ַ�Ӳ�����getUrlAddToParam��",getUrlAddToParam(GetUrl(), "PageSize=20", "replace"))


//��õ�ǰ��ַ��һ�֣�getUrl����http://127.0.0.1/atemp.asp?act=1
//��õ�ǰ��ַ�ڶ��֣�getThisUrl����http://127.0.0.1/atemp.asp?act=1
//��õ�ǰ��ַ�޲�����getThisUrlNoParam����http://127.0.0.1/atemp.asp
//���������ַ��getGoToUrl����http://127.0.0.1/5.asp?act=5&aa=aa&bb=bb
//���������ַ�޲�����getGoToUrlNoParam����http://127.0.0.1/5.asp
//���������ַ���ļ����ƣ�getGoToUrlNoFileName����http://127.0.0.1/
//������webDoMain����http://127.0.0.1
//������host����http://127.0.0.1/
//��õ�ǰ��ַ�Ӳ�����GetUrlAddToParam����http://127.0.0.1/atemp.asp?PageSize=20&act=1
//getThisUrlFileName()    : 4.asp
//getThisUrlFileParam()    : 4.asp?act=11

//Url = getUrlAddToParam("http://www.baidu.com/?a=1&b=2&c=3","?a=11&b=22&c=333","")        'http://www.baidu.com/?a=1&b=2&c=3
//Url = getUrlAddToParam("http://www.baidu.com/?a=1&b=2&c=3","?a=11&b=22&c=333","replace")        'http://www.baidu.com/?a=11&b=22&c=333

//��õ�ǰ������ļ��������
string getThisUrlFileParam(){
    string url="";
    url= getUrl();
    url= mid(url, inStrRev(url, "/") + 1,-1);
    return url;
}

//��õ�ǰ��ַ�޲���
string getThisUrlNoParam(){
    string httpType="";
    if( lCase(cStr(Request.ServerVariables["HTTPS"]))== "off" ){
        httpType= "http://";
    }else{
        httpType= "https://";
    }
    return httpType + cStr(Request.ServerVariables["HTTP_HOST"]) + cStr(Request.ServerVariables["SCRIPT_NAME"]);
}

//��ô�������ַ
string getGoToUrl(){
    return cStr(Request.ServerVariables["HTTP_REFERER"]);
}

//��ô�������ַ �޲���
string getGoToUrlNoParam(){
    string url="";
    url= getGoToUrl();
    if( inStr(url, "?") > 0 ){
        url= mid(url, 1, inStr(url, "?") - 1);
    }
    return url;
}

//��ô�������ַ ���ļ�����
string getGoToUrlNoFileName(){
    string url="";
    url= getGoToUrl();
    if( right(url, 1) != "/" ){
        if( inStrRev(url, "/") > 0 ){
            url= mid(getGoToUrlNoParam(), 1, inStrRev(getGoToUrlNoParam(), "/"));
        }
    }
    if( right(url, 1) != "/" ){ url= url + "/" ;}
    return url;
}
//�Ƴ���ַ�ļ�������
string remoteUrlFileName( string url){
    if( right(url, 1) != "/" ){
        if( inStrRev(url, "/") > 0 ){
            url= mid(url, 1, inStrRev(url, "/"));
        }
    }
    return url;
}
//�Ƴ���ַ��������
string remoteUrlParam( string url){
    if( right(url, 1) != "?" ){
        if( inStrRev(url, "?") > 0 ){
            url= mid(url, 1, inStrRev(url, "?") - 1);
        }
    }
    return url;
}
//�����ַĿ¼����
string getUrlDir( string url){
    return getHandleUrlValue(url, "��ַĿ¼");
}
//��ô����urlֵ 20160701
string getHandleUrlValue( string url, string sType){
    sType= "|" + sType + "|";
    if( inStr(url, "://") > 0 ){
        url= mid(url, inStr(url, "://") + 3,-1);
    }
    //ȥ������
    if( inStr(url, "/") > 0 ){
        url= mid(url, inStr(url, "/") + 1,-1);
    }

    if( inStr(sType, "|��ַĿ¼|") > 0 ){
        if( inStr(url, "/") > 0 ){
            url= mid(url, 1, inStrRev(url, "/") - 1);
        }
    }else{
        if( inStr(url, "/") > 0 ){
            url= mid(url, inStrRev(url, "/") + 1,-1);
        }
        if( inStr(url, "?") > 0 ){
            url= mid(url, 1, inStr(url, "?") - 1);
        }
    }
    if( inStr(sType, "|����|") > 0 || inStr(sType, "|name|") > 0 ){
        if( inStr(url, ".") > 0 ){
            url= mid(url, 1, inStrRev(url, ".") - 1);
        }
    }
    return url;
}

//��ȡ�ͻ���IP��ַ�ڶ���
string getIP2(){

    string x=""; string y=""; string addr="";
    x= cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]);
    y= cStr(Request.ServerVariables["REMOTE_ADDR"]);
    addr= IIF(isNul(x) || lCase(x)== "unknown", y, x);
    if( inStr(addr, ".")== 0 ){ addr= "0.0.0.0" ;}
    return addr;
}

//��ȡIP��ַ ����д�ú����רҵһ�� ��ȫ
string getIP(){

    string strIPAddr="";
    if( cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"])== "" || inStr(cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]), "unknown") > 0 ){
        strIPAddr= cStr(Request.ServerVariables["REMOTE_ADDR"]);
    }else if( inStr(cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]), ",") > 0 ){
        strIPAddr= mid(cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]), 1, inStr(cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]), ",") - 1);
    }else if( inStr(cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]), ";") > 0 ){
        strIPAddr= mid(cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]), 1, inStr(cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]), ";") - 1);
    }else{
        strIPAddr= cStr(Request.ServerVariables["HTTP_X_FORWARDED_FOR"]);
    }
    return aspTrim(mid(strIPAddr, 1, 30));
}

//��÷�����IP
string getServicerIP(){
    return cStr(Request.ServerVariables["LOCAL_ADDR"]);
}

//��÷�����IP   ����
string getRemoteIP(){
    return getServicerIP();
}


//NOVBNet end

//������ַ\ת/  '������
string handleHttpUrl(string httpurl){
    string headStr=""; string url="";
    if( isNul(httpurl) ){ return ""; }//Ϊ�����˳�
    httpurl= replace(aspTrim(httpurl), "\\", "/");
    headStr= mid(httpurl, 1, inStr(httpurl, ":") + 2);
    httpurl= mid(httpurl, inStr(httpurl, ":") + 3,-1);
    httpurl= replace(httpurl, "http://", "��|http|��");
    httpurl= replace(httpurl, "https://", "��|https|��");
    httpurl= replace(httpurl, "ftp://", "��|ftp|��");

    while( inStr(httpurl, "//") > 0){
        httpurl= replace(httpurl, "//", "/");
    }
    httpurl= replace(httpurl, "��|http|��", "http://");
    httpurl= replace(httpurl, "��|https|��", "https://");
    httpurl= replace(httpurl, "��|ftp|��", "ftp://");

    url= headStr + httpurl;
    ///"http://www.qibosoft.com/images/qibosoft/loading.gif/"
    if( left(url, 2)== "/\"" ){
        url= mid(url, 3,-1);
    }
    if( right(url, 2)== "/\"" ){
        url= mid(url, 1, len(url) - 2);
    }
    return url;
}

//�����ļ�/ת\   ʹ����While�жϣ�������
string handleFileUrl(string fileUrl){
    fileUrl= replace(fileUrl, "/", "\\");
    int i=0;
    for( i= 1 ; i<= 99; i++){
        fileUrl= replace(fileUrl, "\\\\", "\\");

        if( inStr(fileUrl, "\\\\")== 0 ){
            break;
        }
    }
    return fileUrl;
}

//������ַ������
string handleUrlComplete(string httpurl){
    string lastStr="";
    string handleUrlComplete= httpurl;
    if( inStr(httpurl, "?") > 0 ){ return handleUrlComplete; }//��?�������˳�
    //��ַ���û��/  �ж����Ϊ���� ����������/�˳�
    if( right(httpurl, 1) != "/" ){
        if( httpurl + "/"== getWebSite(httpurl) ){
            handleUrlComplete= httpurl + "/";
            return handleUrlComplete;
        }
    }
    lastStr= mid(httpurl, inStrRev(httpurl, "/") + 1,-1);
    if( lastStr != "" && inStr(lastStr, ".")== 0 ){
        handleUrlComplete= httpurl + "/";
    }
    return handleUrlComplete;
}

//����ַ�������
string urlAddHttpUrl(string httpurl, string url){
    httpurl= replace(httpurl, "\\", "/");
    url= handleHttpUrl(url);
    if( inStr(lCase(url), "http://")== 0 && inStr(lCase(url), "www.")== 0 ){
        if( right(httpurl, 1)== "/" && left(url, 1)== "/" ){
            url= httpurl + mid(url, 2,-1);
        }else if( right(httpurl, 1) != "/" && left(url, 1) != "/" ){
            url= httpurl + "/" + url;
        }else{
            url= httpurl + url;
        }
    }
    return url;
}
//��������˿ں�
string getPort(){
    string port="";
    port= cStr(cStr(Request.ServerVariables["SERVER_PORT"]));
    if( port != "80" && port != "8080" ){
        port= ":" + port;
    }else{
        port= "";
    }
    return port;
}

//�������
string webDoMain(){
    return "http://" + cStr(Request.ServerVariables["SERVER_NAME"]) + getPort();
}

//��õ�ǰ����
string host(){
    return "http://" + cStr(Request.ServerVariables["HTTP_HOST"]) + "/";
}

//��õ�ǰ���� (����)


//��õ�ǰ���� (����)


//���õ�ǰ��ַ
string getUrl(){
    string getUrl= getThisUrl();
    //GetUrl = WebDoMain() & Request.ServerVariables("SCRIPT_NAME") & Request.ServerVariables("QUERY_STRING")
    return getUrl;
}

//��õ�ǰ��������ַ
string getThisUrl(){
    string url="";
    //vbdel start
    url= cStr(Request.ServerVariables["server_name"]);
    //PHP������ֱ�ӻ�ö˿�
    if( inStr(url, ":")== 0 ){
        url= url + getPort();
    }
    url= url + cStr(Request.ServerVariables["script_name"]);
    if( cStr(Request.ServerVariables["QUERY_STRING"]) != "" ){ url= url + "?" + cStr(Request.ServerVariables["QUERY_STRING"]) ;}
    //vbdel end
    return "http://" + url;
}
//��õ�ǰ���ļ�������ַ
string getThisUrlNoFileName(){
    string url="";
    url= getThisUrl();
    if( right(url, 1) != "/" ){
        if( inStrRev(url, "/") > 0 ){
            url= mid(url, 1, inStrRev(url, "/"));
        }
    }
    return url;
}

//�����ַ�������� http://www.aaa.bb.mywebname.com/   mywebname
string getWebSiteName(string httpurl){
    string url=""; string[] splStr; string s=""; string domainName="";
    url= getWebSite(httpurl);
    url= replace(url, "://", "://.");
    splStr= aspSplit(url, ".");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "/")== 0 && s != "" ){
            if( len(s) >= 4 ){
                domainName= s;
            }
        }
    }
    return domainName;
}

//������ַ �������
string getWebSite( string httpurl){
    //On Error Resume Next
    //�°�����������
    string url=""; string tempHttpUrl=""; bool is_WebSite; string httpHead="";
    string getWebSite="";
    tempHttpUrl= httpurl;
    url= aspTrim(lCase(replace(httpurl, "?", "/")));
    url= replace(replace(url, "\\", "/"), "http://", "");
    if( inStr(url, "/") > 0 ){ url= mid(url, 1, inStr(url, "/") - 1) ;}
    url= "http://" + url + "/";
    string[] splStr; string s=""; string c="";
    httpurl= replace(lCase(httpurl), "http://", "");
    httpHead= "http://";
    //������https://���ְ�ȫ����ʽ  20160526
    if( inStr(lCase(httpurl), "https://") > 0 ){
        httpurl= replace(lCase(httpurl), "https://", "");
        httpHead= "https://";
    }
    //ɾ��/��̨��ֵ20160526
    if( inStr(httpurl, "/") > 0 ){
        httpurl= mid(httpurl, 1, inStr(httpurl, "/") - 1);
    }
    if( inStr(httpurl, "?") > 0 ){ httpurl= mid(httpurl, 1, inStr(httpurl, "?") - 1) ;}
    if( left(httpurl, 9)== "localhost" ){
        if( inStr(httpurl, "/") > 0 ){
            httpurl= mid(httpurl, 1, inStr(httpurl, "/") - 1);
        }else{
            httpurl= "localhost";
        }
    }else if( left(httpurl, 8)== "192.168." || left(httpurl, 9)== "127.0.0.1" ){
        httpurl= httpurl + "/";
        httpurl= "http://" + mid(httpurl, 1, inStr(httpurl, "/") - 1) + "/";
        return httpurl ;
    }else{
        splStr= aspSplit(httpurl, ".");
        if((inStr(httpurl, "www.") > 0 && uBound(splStr) >= 2) || uBound(splStr) >= 1 ){
            if( inStr(httpurl, "/") > 0 ){
                s= mid(httpurl, 1, inStr(httpurl, "/") - 1);
                if( s== getDianNumb(s) ){
                    httpurl= s;
                }
            }
        }else{
            httpurl= ""; //û����Ϊ��
        }
    }
    is_WebSite= false; //������Ϊ��
    c= ".com.hk|.sh.cn|.com.cn|.net.cn|.org.cn";
    c= c + "|.com|.net|.org|.tv|.cc|.info|.cn|.tw|:81|:99|.biz|.mobi|.hk|.us|.la|.gl|.in|.top|.win|.vip|.pw|.me|.wiki|.co";
    splStr= aspSplit(c, "|");
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            if( inStr(httpurl, s) > 0 ){
                httpurl= httpHead + left(httpurl, inStr(httpurl, s) + len(s) - 1) + "/" ; is_WebSite= true ; break;
            }
        }
    }
    getWebSite= left(httpurl, 255); //���������ڣ���ֻ��ȡ255���ַ�
    //GetWebSite = ""                        '������������Ϊ�� 20150104
    if( getWebSite== "http:///" ){ getWebSite= "" ;}//û���ҵ�����
    if( is_WebSite== false ){
        getWebSite= "";
    }

    return getWebSite;
}
//����Ƿ�Ϊ��ַ
bool checkUrl(string url){
    return IIF(getWebSite(url)== "", false, true);
}
//����Ƿ�Ϊ��ַ
bool checkHttpUrl(string url){
    return checkUrl(url);
}
//����Ƿ�Ϊ��ַ
bool isUrl(string url){
    return checkUrl(url);
}
//����Ƿ�Ϊ��ַ
bool isHttpUrl(string url){
    return checkUrl(url);
}

//������ַ
//��������� style="background: url(20130510162636_96168.jpg) no-repeat scroll center top; height: 426px;cursor:pointer; width: 100%; margin:0 auto;" 20160701
string fullHttpUrl( string httpurl, string url){
    //On Error Resume Next:    err.clear                                                                       '������� Ҫ��Ȼ�ᱨ��
    string fullHttpUrl="";
    string rootUrl=""; string thisUrl=""; string[] splStr; int i=0; string s=""; string c=""; string parentUrl=""; string parentParentUrl=""; string parentParentParentUrl=""; string rootWebSite=""; string thisWebSite=""; bool isHandle; string lCaseUrl="";
    string styleStart=""; string styleEnd="";
    httpurl= phpTrim(httpurl); //������߿ո�
    url= phpTrim(url); //������߿ո�
    lCaseUrl= lCase(url);
    if( url== "" ){ return "" ; }//��ַΪ���˳�(20150805)
    if( aspTrim(httpurl)== "" ){ return url ; }//����ַΪ���˳� ������ַ
    httpurl= replace(httpurl, "\\", "/"); //����ַ\ת/����
    url= replace(url, "\\", "/"); //����ַ\ת/����

    //��ַǰ�����ַ�Ϊ//���˳�
    if( left(url, 2)== "//" ){
        fullHttpUrl= "http:" + url;
        return fullHttpUrl;
    }
    //����style��ʽ�ﱳ��ͼƬ
    url= hanldeStyleBackgroundUrl(url, styleStart, styleEnd);
    rootUrl= getWebSite(httpurl); //��������Ҳ��������ַ
    rootWebSite= rootUrl;
    thisWebSite= getWebSite(url);
    if( right(rootUrl, 1)== "/" ){ rootUrl= left(rootUrl, len(rootUrl) - 1); }
    thisUrl= left(httpurl, inStrRev(httpurl, "/")); //��ǰ��ַ
    splStr= aspSplit(httpurl, "/");
    for( i= 0 ; i<= uBound(splStr); i++){
        if( i + 1== uBound(splStr) ){ parentUrl= c ;}
        if( i + 2== uBound(splStr) ){ parentParentUrl= c ;}
        if( i + 3== uBound(splStr) ){ parentParentParentUrl= c ;}
        s= splStr[i];
        c= c + s + "/";
    }
    url= aspTrim(url); //ȥ����ַ���ҿո�
    isHandle= false; //����Ϊ��
    if( url != "" && inStr(left(url, 10), "www.")== 0 && inStr(left(url, 10), "http://")== 0 && inStr(left(url, 10), "https://")== 0 ){
        isHandle= true;
        if( rootWebSite != thisWebSite ){
            if( rootWebSite== replace(thisWebSite, "http://", "http://www.") ){
                isHandle= false;
                if( inStr(lCase(url), "http://") > 0 ){
                    url= "http://www." + right(url, len(url) - 7);
                }else{
                    url= "http://www." + url;
                }
            }
        }
    }
    //�����Ƿ�Ϊ��
    if( isHandle== true ){
        if( left(url, 1)== "/" ){
            url= rootUrl + url;
        }else if( left(url, 9)== "../../../" ){
            url= parentParentParentUrl + right(url, len(url) - 9);
        }else if( left(url, 6)== "../../" ){
            url= parentParentUrl + right(url, len(url) - 6);
        }else if( left(url, 3)== "../" ){
            url= parentUrl + right(url, len(url) - 3);
        }else if( left(url, 2)== "./" ){
            url= thisUrl + mid(url, 3,-1);
        }else{
            url= thisUrl + url;
        }
    }
    if( inStr(lCase(url), "http://")== 0 && inStr(lCase(url), "https://")== 0 ){
        if( inStr(lCase(httpurl), "http://") > 0 && inStr(lCase(url), "http://")== 0 ){
            url= "http://" + url;
        }else if( inStr(lCase(httpurl), "https://") > 0 && inStr(lCase(url), "https://")== 0 ){
            url= "https://" + url;
        }
    }
    fullHttpUrl= styleStart + url + styleEnd;

    return fullHttpUrl;
}
//������ʽ�ﱳ��ͼƬ20160701
string hanldeStyleBackgroundUrl( string url, string styleStart, string styleEnd){
    string lCaseUrl="";
    url= phpTrim(url); //������߿ո�
    lCaseUrl= lCase(url);
    url= replace(url, "\\", "/"); //����ַ\ת/����
    if((inStr(url, "background") > 0 || inStr(url, "background-image") > 0) && inStr(url, "(") > 0 && inStr(url, ")") > 0 ){
        styleStart= mid(url, 1, inStr(url, "("));
        styleEnd= mid(url, inStr(url, ")"),-1);
        url= mid(url, len(styleStart) + 1,-1);
        url= phpTrim(mid(url, 1, inStr(url, ")") - 1));
        url= replace(replace(url, "'", ""), "\"", "");
    }
    return url;
}

//����������ַ����20150728
string batchFullHttpUrl(string webSite, string urlList){
    string[] splStr; string url=""; string c="";
    splStr= aspSplit(urlList, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        if( len(url) > 3 ){
            if( c != "" ){ c= c + vbCrlf() ;}
            c= c + fullHttpUrl(webSite, url);
        }
    }
    return c;
}


//��ַ�����ַ� ��ദ��
string uRLJianJieHandle( string url){
    url= replace(url, "&amp;", "&");
    return url;
}

//URL���� �������С�������
string urlToAsc(string url){
    int i=0;
    string urlToAsc="";
    for( i= 1 ; i<= len(url); i++){
        urlToAsc= urlToAsc + "%" + hex(ascW(mid(url, i, 1)));
    }
    return urlToAsc;
}


//�����վ����
string getWebTitle(string content){
    return getStrCut(content, "<title>", "</title>", 0);
}

//���������ַ�б� (ȱ������ַȫ��Сд��20150728)
string getContentAHref(string httpurl, string content, string pubAHrefList, string pubATitleList){
    int i=0; string s=""; string tempS=""; string lalType=""; int nLen=0; string lalStr=""; string c="";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( s== "<" ){
            tempS= lCase(mid(content, i,-1));
            lalType= lCase(mid(tempS, 1, inStr(tempS, " ")));
            if( lalType== "<a " ){
                lalStr= mid(tempS, 1, inStr(tempS, "</") + 2);
                nLen= len(lalStr) - 1;
                c= c + handleLink(httpurl, lalStr, "href", "", "url", pubAHrefList, pubATitleList) + vbCrlf();
                i= i + nLen;
            }
        }
        doEvents();
    }
    if( c != "" ){ c= left(c, len(c) - 2); }
    return c;
}

//���������ͼƬ�б�
string getContentImgSrc(string httpurl, string content, string pubAHrefList, string pubATitleList){
    int i=0; string s=""; string tempS=""; string lalType=""; int nLen=0; string lalStr=""; string c="";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( s== "<" ){
            tempS= lCase(mid(content, i,-1));
            lalType= lCase(mid(tempS, 1, inStr(tempS, " ")));
            if( lalType== "<img " ){
                lalStr= mid(tempS, 1, inStr(tempS, ">"));
                nLen= len(lalStr) - 1;
                //Call Echo(I,LalStr)
                c= c + handleLink(httpurl, lalStr, "src", "", "url", pubAHrefList, pubATitleList) + vbCrlf();
                i= i + nLen;
            }
        }
        doEvents();
    }
    if( c != "" ){ c= left(c, len(c) - 2); }
    return c;
}

//����������ַ���� sType=|*|link|img|a|script|embed|param|meta|
string handleConentUrl(string httpurl, string content, string sType, string pubAHrefList, string pubATitleList){
    int i=0; string s=""; string yuanStr=""; string tempS=""; string lalType=""; int nLen=0; string lalStr=""; string c="";
    sType= "|" + sType + "|";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( s== "<" ){
            yuanStr= mid(content, i,-1);
            tempS= lCase(yuanStr);
            tempS= replace(replace(tempS, chr(10).ToString(), " " + vbCrlf()), chr(13).ToString(), " " + vbCrlf()); //�ô���ͼƬ�زĸ�����  ����  <img����src=""  Ҳ���Ի�� 20150714
            lalStr= mid(yuanStr, 1, inStr(yuanStr, ">"));
            lalType= lCase(mid(tempS, 1, inStr(tempS, " ")));
            if( lalType== "<link " &&(inStr(sType, "|link|") > 0 || inStr(sType, "|*|") > 0) ){
                nLen= len(lalStr) - 1;
                c= c + handleLink(httpurl, lalStr, "href", "", "", pubAHrefList, pubATitleList);
                i= i + nLen;
            }else if( lalType== "<img " &&(inStr(sType, "|img|") > 0 || inStr(sType, "|*|") > 0) ){
                nLen= len(lalStr) - 1;
                c= c + handleLink(httpurl, lalStr, "src", "", "", pubAHrefList, pubATitleList);
                i= i + nLen;
            }else if( lalType== "<a " &&(inStr(sType, "|a|") > 0 || inStr(sType, "|*|") > 0) ){
                nLen= len(lalStr) - 1;
                //û��javascript�����У����ǻ����в���֮��
                if( inStr(lCase(lalStr), "javascript:")== 0 ){
                    c= c + handleLink(httpurl, lalStr, "href", "", "", pubAHrefList, pubATitleList);
                }else{
                    c= c + lalStr;
                }
                i= i + nLen;
            }else if( lalType== "<script " &&(inStr(sType, "|script|") > 0 || inStr(sType, "|*|") > 0) ){
                nLen= len(lalStr) - 1;
                if( inStr(lCase(lalStr), "src") > 0 ){
                    c= c + handleLink(httpurl, lalStr, "src", "", "", pubAHrefList, pubATitleList);
                }else{
                    c= c + lalStr;
                }
                i= i + nLen;
            }else if( lalType== "<embed " &&(inStr(sType, "|embed|") > 0 || inStr(sType, "|*|") > 0) ){
                nLen= len(lalStr) - 1;
                c= c + handleLink(httpurl, lalStr, "src", "", "", pubAHrefList, pubATitleList);
                i= i + nLen;
            }else if( lalType== "<param " &&(inStr(sType, "|param|") > 0 || inStr(sType, "|*|") > 0) ){
                nLen= len(lalStr) - 1;
                if( inStr(lCase(lalStr), "movie") > 0 ){
                    c= c + handleLink(httpurl, lalStr, "value", "", "", pubAHrefList, pubATitleList);
                }else{
                    c= c + lalStr;
                }
                i= i + nLen;
            }else if( lalType== "<meta " &&(inStr(sType, "|meta|") > 0 || inStr(sType, "|*|") > 0) ){
                nLen= len(lalStr) - 1;
                //�滻�ؼ���
                if( inStr(lCase(lalStr), "keywords") > 0 ){
                    c= c + handleLink(httpurl, lalStr, "content", webKeywords, "", pubAHrefList, pubATitleList);
                    //�滻��վ����
                }else if( inStr(lCase(lalStr), "description") > 0 ){
                    c= c + handleLink(httpurl, lalStr, "content", webDescription, "", pubAHrefList, pubATitleList);
                }else{
                    c= c + lalStr;
                }
                i= i + nLen;

            }else if( lalType== "<input " &&(inStr(sType, "|src|") > 0 || inStr(sType, "|*|") > 0) && inStr(lCase(lalStr), " src=") > 0 ){
                nLen= len(lalStr) - 1;
                c= c + handleLink(httpurl, lalStr, "src", "", "", pubAHrefList, pubATitleList);
                i= i + nLen;

            }else if( inStr(sType, "|imgstyle|") > 0 || inStr(sType, "|*|") > 0 ){
                nLen= len(lalStr) - 1;
                if( inStr(lCase(lalStr), "url") > 0 && inStr(lCase(lalStr), ":") > 0 && inStr(lCase(lalStr), "(") > 0 ){
                    c= c + handleLink(httpurl, lalStr, "style", "", "", pubAHrefList, pubATitleList);
                    i= i + nLen;
                }else{
                    c= c + s;
                }
            }else{
                c= c + s;
            }
        }else{
            c= c + s;
        }
        doEvents();
    }
    return c;
}


//�滻������ȫ��JsĿ¼ 20150722  call rwend(handleConentUrl("/admin/js/", "<script src='aa/js.js' ><script src=""bb/js.js"" >","",""))
string replaceContentJsDir(string content, string dirPath, string pubAHrefList, string pubATitleList){
    string[] splStr; string s=""; string c="";
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( c != "" ){ c= c + vbCrlf() ;}
        if( inStr(s, "<script ") > 0 && inStr(s, "</"+"script>") > 0 ){
            s= handleLink(dirPath, s, "src", "", "replaceDir", pubAHrefList, pubATitleList);
        }
        c= c + s;
    }
    return c;
}


//�������ӵ�ַ HttpUrl=׷����ַ��Content=����  SType=����
//�滻Ŀ¼����  call rw(HandleLink("Js/", "111111<script src=""js/Jquery.Min.js""></"& "script>","src", "newurl", "replaceDir","",""))
string handleLink(string httpurl, string content, string sType, string setStr, string urlOrContent, string pubAHrefList, string pubATitleList){

    string[] splStr; int i=0; string s=""; string c=""; string tempContent=""; string findUrl=""; string handleUrl=""; string startStr=""; string endStr=""; string s1=""; string s2=""; string tempHttpUrl="";int nLen=0;
    tempHttpUrl= httpurl;
    urlOrContent= lCase(urlOrContent);
    content= replace(replace(content, "= ", "="), "= ", "=");
    content= replace(replace(content, " =", "="), " =", "=");
    tempContent= lCase(content);
    string handleLink= "";
    //û�������˳�
    if( inStr(tempContent, " href=")== 0 && inStr(tempContent, " src=")== 0 && sType != "style" ){
        handleLink= "";
        return handleLink;
    }else if( inStr(tempContent, " href=\\\"") > 0 ){
        content= replace(content, "\\\"", "\"") ; tempContent= lCase(content);
    }
    startStr= sType + "=\"";
    endStr= "\"";
    if( inStr(tempContent, startStr) > 0 && inStr(tempContent, endStr) > 0 ){
        //call echo("��ʾ","1")
        findUrl= strCut(content, startStr, endStr, 2);
        if( setStr != "" ){
            handleUrl= setStr;
        }else{
            handleUrl= fullHttpUrl(httpurl, findUrl);
            //�滻Ŀ¼
            if( urlOrContent== "replacedir" ){
                handleUrl= tempHttpUrl + handleFilePathArray(handleUrl)[2];
            }
            pubAHrefList= pubAHrefList + hanldeStyleBackgroundUrl(handleUrl, "", "") + vbCrlf();
            //���ӱ���
            nLen= inStr(content, ">");
            s2= right(content, len(content) - nLen);
            s2= mid(s2, 1, inStrRev(s2, "</") - 1);
            s2= replace(s2, vbCrlf(), "�����С�");
            pubATitleList= pubATitleList + s2 + vbCrlf();
        }
        if( findUrl != handleUrl ){
            //ǿǿǿ���滻
            nLen= inStr(content, startStr) - 1 + len(startStr); //��������TempContent������Content��Ϊ�д�Сд������20140726
            s2= right(content, len(content) - nLen);
            s2= mid(s2, inStr(s2, endStr),-1);
            s1= left(content, nLen);
            content= s1 + handleUrl + s2;
        }
        if( urlOrContent== "url" ){
            handleLink= handleUrl;
        }else{
            handleLink= content;
        }
        return handleLink;
    }
    startStr= sType + "='";
    endStr= "'";
    if( inStr(tempContent, startStr) > 0 && inStr(tempContent, endStr) > 0 ){
        //call echo("��ʾ","2")
        findUrl= strCut(tempContent, startStr, endStr, 2);
        if( setStr != "" ){
            handleUrl= setStr;
        }else{
            handleUrl= fullHttpUrl(httpurl, findUrl);
            //�滻Ŀ¼
            if( urlOrContent== "replacedir" ){
                handleUrl= tempHttpUrl + handleFilePathArray(handleUrl)[2];
            }
            pubAHrefList= pubAHrefList + hanldeStyleBackgroundUrl(handleUrl, "", "") + vbCrlf();
            //���ӱ���
            nLen= inStr(content, ">");
            s2= right(content, len(content) - nLen);
            s2= mid(s2, 1, inStrRev(s2, "</") - 1);
            s2= replace(s2, vbCrlf(), "�����С�");
            pubATitleList= pubATitleList + s2 + vbCrlf();
        }
        if( findUrl != handleUrl ){
            //ǿǿǿ���滻
            nLen= inStr(content, startStr) - 1 + len(startStr);
            s2= right(content, len(content) - nLen);
            s2= mid(s2, inStr(s2, endStr),-1);
            s1= left(content, nLen);
            content= s1 + handleUrl + s2;
        }
        if( urlOrContent== "url" ){
            handleLink= handleUrl;
        }else{
            handleLink= content;
        }
        return handleLink;
    }
    startStr= sType + "=";
    endStr= ">"; //�������֮�ҵ� �ո񻻳�>
    if( inStr(tempContent, startStr) > 0 && inStr(tempContent, endStr) > 0 ){
        findUrl= strCut(tempContent, startStr, endStr, 2);

        if( setStr != "" ){
            handleUrl= setStr;
        }else{
            handleUrl= fullHttpUrl(httpurl, findUrl);
            //�滻Ŀ¼
            if( urlOrContent== "replacedir" ){
                handleUrl= tempHttpUrl + handleFilePathArray(handleUrl)[2];
            }
            pubAHrefList= pubAHrefList + hanldeStyleBackgroundUrl(handleHttpUrl(handleUrl), "", "") + vbCrlf();
            //���ӱ���
            nLen= inStr(content, ">");
            s2= right(content, len(content) - nLen);
            s2= mid(s2, 1, inStrRev(s2, "</") - 1);
            s2= replace(s2, vbCrlf(), "�����С�");
            pubATitleList= pubATitleList + s2 + vbCrlf();
        }
        if( findUrl != handleUrl ){
            //ǿǿǿ���滻
            nLen= inStr(content, startStr) - 1 + len(startStr);
            s2= right(content, len(content) - nLen);
            s2= mid(s2, inStr(s2, endStr),-1);
            s1= left(content, nLen);
            content= s1 + handleUrl + s2;
        }
        if( urlOrContent== "url" ){
            handleLink= handleUrl;
        }else{
            handleLink= content;
        }
        return handleLink;
    }

    if( urlOrContent != "url" ){ handleLink= content ;}
    createAddFile("���������б�.txt", httpurl + vbCrlf() + content + vbCrlf() + sType + vbCrlf() + setStr + vbCrlf() + urlOrContent + vbCrlf() + "----------------------" + vbCrlf());
    return handleLink;
}


//�����վĿ¼�ļ������� \Templates\WeiZhanLue\  �õ�WeiZhanLue
string getEndUrlHandleName(string fileUrl){
    string url="";
    url= replace(aspTrim(fileUrl), "\\", "/");
    if( right(url, 1)== "/" ){ url= mid(url, 1, len(url) - 1) ;}
    url= mid(url, inStrRev(url, "/") + 1,-1);
    return url;
}


//����б��в�ͬ�����б�
string getUrlListInWebSiteList(string content){
    string url=""; string urlList=""; string[] splStr;
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        url= getWebSite(url);
        if( url != "" && inStr(vbCrlf() + urlList + vbCrlf(), vbCrlf() + url + vbCrlf())== 0 ){
            urlList= urlList + url + vbCrlf();
        }
        doEvents();
    }
    return urlList;
}

//��õ�ǰ��ַ���ļ�����
string getThisUrlFileName(){
    string url="";
    url= cStr(Request.ServerVariables["SCRIPT_NAME"]);
    if( left(url, 1)== "/" ){ url= right(url, len(url) - 1); }
    return url;
}

//������վHTML��Img    д�ò����ر�����ƺ�  Content = HandleWebHtmlImg("/aa/bb/",Content)
string handleWebHtmlImg(string rootPath, string content, string pubAHrefList, string pubATitleList){
    string imgList=""; string[] splStr; string imgUrl=""; string sNewImgUrl="";
    string startStr=""; string endStr="";
    imgList= getContentImgSrc("", content, pubAHrefList, pubATitleList);
    splStr= aspSplit(imgList, vbCrlf());
    foreach(var eachimgUrl in splStr){
        imgUrl=eachimgUrl;
        if( imgUrl != "" ){
            sNewImgUrl= handleHttpUrl(imgUrl);
            if( inStr(sNewImgUrl, "/") > 0 ){
                sNewImgUrl= mid(sNewImgUrl, inStrRev(sNewImgUrl, "/") + 1,-1);
            }
            sNewImgUrl= rootPath + sNewImgUrl;
            //Call Echo(NewImgUrl,ImgUrl)
            startStr= "src=\"" ; endStr= "\"";
            if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
                content= regExp_Replace(content, startStr + imgUrl + endStr, startStr + sNewImgUrl + endStr);
            }
            startStr= "src='" ; endStr= "'";
            if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
                content= regExp_Replace(content, startStr + imgUrl + endStr, startStr + sNewImgUrl + endStr);
            }
        }
    }
    return content;
}

//������վCss��Img    д�ò����ر�����ƺ�  Content = HandleWebHtmlImg("/aa/bb/",Content)
string handleWebCssImg(string rootPath, string content){
    string startStr=""; string endStr=""; string imgList=""; string[] splStr; string c=""; string imgUrl=""; string sNewImgUrl="";
    startStr= "url\\(";
    endStr= "\\)";
    imgList= getArray(content, startStr, endStr, false, false);
    //Call RwEnd(ImgList)
    splStr= aspSplit(imgList, "$Array$");
    foreach(var eachimgUrl in splStr){
        imgUrl=eachimgUrl;
        if( imgUrl != "" ){
            sNewImgUrl= handleHttpUrl(imgUrl);
            if( inStr(sNewImgUrl, "/") > 0 ){
                sNewImgUrl= mid(sNewImgUrl, inStrRev(sNewImgUrl, "/") + 1,-1);
            }
            sNewImgUrl= rootPath + sNewImgUrl;
            startStr= "url(";
            endStr= ")";
            if( inStr(content, startStr) > 0 && inStr(content, endStr) > 0 ){
                //call echo(StartStr,"StartStr")
                content= regExp_Replace(content, startStr + imgUrl + endStr, startStr + sNewImgUrl + endStr);
            }
        }
    }
    return content;
}

//����������ַ������
string batchHandleUrlIntegrity(string httpurl, string urlList){
    string[] splUrl; string url=""; string c="";
    splUrl= aspSplit(urlList, vbCrlf());
    foreach(var eachurl in splUrl){
        url=eachurl;
        if( url != "" ){
            url= fullHttpUrl(httpurl, url);
            if( inStr(vbCrlf() + c + vbCrlf(), vbCrlf() + url + vbCrlf())== 0 ){
                c= c + url + vbCrlf();
            }
        }
    }
    return c;
}

//��������������ͼƬ·�� Ŀ¼����ʾͼƬ
string replaceContentImagePath(string rootFolder, string content){
    string imageFile=""; string toImageFile=""; string imageList=""; string[] splxx;
    rootFolder= handleHttpUrl(rootFolder);
    if( right(rootFolder, 1) != "/" ){ rootFolder= rootFolder + "/" ;}
    imageList= getDirFileNameList(rootFolder, "");
    splxx= aspSplit(imageList, vbCrlf());
    foreach(var eachimageFile in splxx){
        imageFile=eachimageFile;
        if( imageFile != "" ){
            toImageFile= "file:///" + rootFolder + imageFile;
            //html��ͼƬ·���滻
            content= replace(content, "\"" + imageFile + "\"", "\"" + toImageFile + "\"");
            content= replace(content, "'" + imageFile + "'", "\"" + toImageFile + "\"");
            content= replace(content, "=" + imageFile + " ", "\"" + toImageFile + "\"");
            content= replace(content, "=" + imageFile + ">", "\"" + toImageFile + "\"");
            //Css��ͼƬ·���滻
            content= replace(content, "(" + imageFile + ")", "(" + toImageFile + ")");
            content= replace(content, "(" + imageFile + ";", "(" + toImageFile + ";");
        }
    }
    return content;
}

//���Css�����б� ������
string getCssLinkList( string content){
    string startStr=""; string endStr=""; string[] splStr; string s=""; string c=""; string fileName="";
    startStr= "<link" ; endStr= "/>";
    content= getArray(content, startStr, endStr, false, false);
    splStr= aspSplit(content, "$Array$");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(lCase(s), "stylesheet") > 0 ){
            fileName= strCut(s, "href=\"", "\"", 2);
            echo(fileName, s);
            c= c + fileName + vbCrlf();
        }
    }
    return c;
}

//���Html��ͼƬ��ַ�б�
string getHtmlBackGroundUrlList(string content){
    int i=0; string s=""; string yuanStr=""; string tempS=""; string lalType=""; int nLen=0; string lalStr=""; string c=""; string startStr=""; string endStr="";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( s== "<" ){
            yuanStr= mid(content, i,-1);
            tempS= lCase(yuanStr);
            lalStr= mid(yuanStr, 1, inStr(yuanStr, ">"));
            lalType= lCase(mid(tempS, 1, inStr(tempS, " ")));
            if( inStr(lalStr, "url(") > 0 ){
                startStr= "url(" ; endStr= ")";
                c= c + strCut(lalStr, startStr, endStr, 2) + vbCrlf();
                i= i + nLen;
            }
        }
        doEvents();
    }
    return c;
}

//���ͼƬ��ַ�б�
string getImgUrlList(string content){
    int i=0; string s=""; string yuanStr=""; string tempS=""; string lalType=""; int nLen=0; string lalStr=""; string c="";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( s== "<" ){
            yuanStr= mid(content, i,-1);
            tempS= lCase(yuanStr);
            lalStr= mid(yuanStr, 1, inStr(yuanStr, ">"));
            lalType= lCase(mid(tempS, 1, inStr(tempS, " ")));
            if( lalType== "<img " ){
                c= c + getLinkUrl(lalStr, "src") + vbCrlf();
                i= i + nLen;

            }
        }
        doEvents();
    }
    return c;
}

//���img��a��ַ ��GetLinkUrl(LalStr, "src")
string getLinkUrl( string linkStr, string linkType){
    string tempLinkStr=""; string startStr=""; string endStr=""; string linkUrl="";
    linkStr= replace(replace(linkStr, "= ", "="), "= ", "=");
    linkStr= replace(replace(linkStr, " =", "="), " =", "=");
    tempLinkStr= lCase(linkStr);
    string getLinkUrl="";
    startStr= linkType + "=\"";
    endStr= "\"";
    if( inStr(tempLinkStr, startStr) > 0 && inStr(tempLinkStr, endStr) > 0 ){
        linkUrl= strCut(tempLinkStr, startStr, endStr, 2);
        getLinkUrl= linkUrl;
    }
    return getLinkUrl;
}

//�Ǳ�����ַ��Զ����ַ        20141120
string localUrlOrRemoteUrl(string filePath, bool isUrlOK){
    string httpurl="";
    isUrlOK= false;
    filePath= aspTrim(filePath);
    httpurl= replace(filePath, "\\", "/");
    if( left(lCase(httpurl), 7)== "http://" || left(lCase(httpurl), 4)== "www." || left(lCase(httpurl), 4)== "[��ַ]" ){
        isUrlOK= true;
    }
    if( isUrlOK== false ){
        filePath= setFileName(filePath);
    }
    return filePath;
}

//����Ƿ�ΪԶ����ַ
bool checkRemoteUrl(string url){
    string httpurl="";
    url= aspTrim(url);
    httpurl= replace(unSetFileName(url), "\\", "/");
    bool checkRemoteUrl= false;
    if( left(lCase(httpurl), 7)== "http://" || left(lCase(httpurl), 4)== "www." || left(lCase(httpurl), 4)== "[��ַ]" ){
        if( left(lCase(httpurl), 4)== "[��ַ]" ){
            url= mid(url, inStr(url, "[��ַ]") + 1,-1);
        }
        checkRemoteUrl= true;
    }
    return checkRemoteUrl;
}



//�������ƣ�AsaiLinkAdd   20141217���ñ���
//�������ã������Զ��������
//����Response.Write AsaiLinkAdd("http://www.wzl99.com/С�С��http://www.wzl99.com/С�С��http://www.wzl99.com//С�С��http://www.wzl99.com/С�С��http://www.wzl99.com/С�С��www.wzl99.com/С�С��fengying789@126.comС�С��")
string asaiLinkAdd(string str){//���պ���
    return "";
}

//�������ƣ�AsaiLinkDel Asai(ǳ��)   20141217���ñ���
//�������ã������Զ�ȥ������
//����Response.Write AsaiLinkDel("<a href='http://www.wzl99.com/' target='_blank'>http://www.wzl99.com/</a>С�С��<a href='http://www.wzl99.com/' target='_blank'>http://www.wzl99.com/</a>С�")
string asaiLinkDel(string htmlStr){//���պ���
    return "";
}

//�ж������Ƿ�Ϸ�  �ݴ�
bool iswww(string strng){//���պ���
    return false;
}


//��ǿ��20150220
//׷����ַ����20150121  getUrlAddToParam("aa","&a=b","replace")   addto replace    STypeΪ׷�ӻ����滻
//Url = getUrlAddToParam("http://www.baidu.com/?a=1&b=2&c=3","?a=11&b=22&c=333","")        'http://www.baidu.com/?a=1&b=2&c=3
//Url = getUrlAddToParam("http://www.baidu.com/?a=1&b=2&c=3","?a=11&b=22&c=333","replace")        'http://www.baidu.com/?a=11&b=22&c=333
//Url = getUrlAddToParam(GetUrl(),"id=" & Rs("Id"),"replace")
//Call Echo(Url,getUrlAddToParam(Url,"id=1&aa=1&bb=2","delete"))          ����ɾ������
string getUrlAddToParam( string url, string addToUrl, string sType){
    string content=""; string[] splStr; string[] splxx; string s=""; string c=""; string httpurl=""; string urlFileName=""; string webSite="";
    string urlParam=""; //��ַ���� �ǻ����ַ��̨����ֵ
    string paramName=""; //��������
    string paramValue=""; //����ֵ
    string paramNameList=""; //���������б���ֹ�ظ�
    bool isHandle; //����Ϊ��

    addToUrl= handleHttpUrl(addToUrl);

    //������ַ
    url= aspTrim(url);
    //��ǰ��ַ���һ���ַ�Ϊ?��&��ɾ���� ����
    if( right(url, 1)== "?" || right(url, 1)== "&" ){
        url= aspTrim(mid(url, 1, len(url) - 1));
    }
    //����׷����ַ
    addToUrl= aspTrim(addToUrl);
    //׷����ַ���һ���ַ�Ϊ?��&��ɾ���� ����
    if( left(addToUrl, 1)== "?" || left(addToUrl, 1)== "&" ){
        addToUrl= aspTrim(mid(addToUrl, 2,-1));
    }
    //��ַΪ���򷵻�׷����ַ ���˳�
    if( url== "" ){ return "?" + addToUrl ; }

    httpurl= url;
    if( inStr(url, "?") > 0 ){
        httpurl= mid(url, 1, inStr(url, "?") - 1); //��õ�ǰ·����ַ
        webSite= getWebSite(url); //��������
        urlParam= mid(url, inStr(url, "?") + 1,-1);

        httpurl= handleHttpUrl(httpurl);
        if( right(httpurl, 1) != "/" ){
            urlFileName= mid(httpurl, inStrRev(httpurl, "/") + 1,-1);
            httpurl= left(httpurl, len(httpurl) - len(urlFileName));
            //Call Echo(HttpUrl,UrlFileName)
        }
    }

    //����ѡ��  ׷�� �����滻
    if( lCase(sType)== "replace" || sType== "�滻" ){
        content= addToUrl + "&" + urlParam;
        //Call echo("Content",Content)
        if( inStr(addToUrl, "?") > 0 ){
            urlFileName= mid(addToUrl, 1, inStr(addToUrl, "?") - 1);
            //Call Echo(AddToUrl,UrlFileName)
        }
    }else if((lCase(sType) != "delete" || lCase(sType) != "del") ){
        content= urlParam + "&" + addToUrl;
    }

    content= replace(content, "?", "&");

    //����ɾ������ 20150210
    if( lCase(sType)== "delete" || lCase(sType)== "del" ){
        splStr= aspSplit(lCase(addToUrl), "&") ; addToUrl= "&";
        foreach(var eachs in splStr){
            s=eachs;
            if( inStr(s, "=")>0 ){
                s= mid(s, 1, inStr(s, "=") - 1);
            }
            if( s != "" ){
                addToUrl= addToUrl + s + "&";
            }
        }
        //Call Eerr("AddToUrl",AddToUrl)
    }

    //Call Echo("Content",Content)
    splStr= aspSplit(content, "&");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "=") > 0 ){
            splxx= aspSplit(s, "=");
            paramName= splxx[0];
            paramValue= splxx[1];

            isHandle= true;

            if( lCase(sType)== "delete" || lCase(sType)== "del" ){
                if( inStr("&" + addToUrl + "&", "&" + lCase(paramName) + "&") > 0 ){
                    isHandle= false;
                }
            }

            if( inStr("|" + paramNameList + "|", "|" + lCase(paramName) + "|")== 0 && isHandle== true ){
                paramNameList= paramNameList + lCase(paramName) + "|";
                c= c + IIF(c== "", "?", "&");
                c= c + paramName + "=" + paramValue;
            }
        }
    }



    c= urlFileName + c;
    if( getWebSite(c)== "" ){
        if( left(addToUrl, 1)== "/" ){
            c= webSite + c;
        }else{
            c= httpurl + c;
        }

    }

    c= replace(c, "\\", "/"); //20160313
    return c;
}




//�����ַ 20150706 call echo("",groupUrl("www.baidu.com//","/1.asp"))
string groupUrl(string url1, string url2){
    string urlType=""; int i=0;
    urlType= "/";
    url1= replace(url1, IIF(urlType== "/", "\\", "/"), urlType);
    url2= replace(url2, IIF(urlType== "/", "\\", "/"), urlType);
    url1= phpTrim(url1);
    url2= phpTrim(url2);
    for( i= 0 ; i<= 99; i++){
        if( right(url1, 1)== urlType ){
            url1= mid(url1, 1, len(url1) - 1);
        }else{
            break;
        }
    }
    for( i= 0 ; i<= 99; i++){
        if( left(url2, 1)== urlType ){
            url2= mid(url2, 2,-1);
        }else{
            break;
        }
    }
    return url1 + urlType + url2;
}



//����POST��Cookes���ͷ�ʽ�Ĳ�������
string handlePostCookiesParame(string parame, string sType){
    string[] splStr; string s=""; string c=""; string leftC=""; string rightC="";
    splStr= aspSplit(parame, "&");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "=") > 0 ){
            leftC= mid(s, 1, inStr(s, "="));
            rightC= mid(s, inStr(s, "=") + 1,-1);
            if( lCase(sType)== "post" || sType== "" ){

                if( c != "" ){
                    c= c + "&";		 //������ת.netʱ����=�� ? 20160727  ������һ��
                }
                rightC= escape(rightC);
            }else if( lCase(sType)== "cookies" || lCase(sType)== "cookie" ){
                if( c != "" ){ c= c + ";" ;}
                rightC= urlEncoding(rightC);
            }
            c= c + leftC + rightC;
            //call echo(leftC,RightC)
        }
    }
    return c;
}


//�Ƴ���ַ�в���20150724
string remoteHttpUrlParameter(string httpurl){
    string[] splStr; string s=""; string c=""; string leftC=""; string rightC="";
    //û��?���˳�
    if( inStr(httpurl, "?")== 0 ){
        string remoteHttpUrlParameter= httpurl;
        return remoteHttpUrlParameter;
    }
    splStr= aspSplit(httpurl, "&");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "=") > 0 ){
            leftC= mid(s, 1, inStr(s, "="));
            rightC= mid(s, inStr(s, "=") + 1,-1);
            if( c != "" ){ c= c + "&" ;}
            c= c + leftC;
        }
    }
    return c;
}


//��⵱ǰ��ַ���ļ������Ƿ����(20150909)
//�÷� call echo("checkUrlName",checkUrlName("1|2|3|"))      <%=IIF(checkUrlName("1|2|3|"),"111","222")% >
bool checkUrlName(string searchUrlName){
    string[] splStr; string urlName=""; string url="";
    searchUrlName= lCase(searchUrlName); //������ַ����תСд
    url= lCase(cStr(Request.ServerVariables["script_name"]));
    splStr= aspSplit(searchUrlName, "|");
    foreach(var eachurlName in splStr){
        urlName=eachurlName;
        if( urlName != "" ){
            if( inStr(url, urlName) > 0 ){
                bool checkUrlName= true;
                return checkUrlName;
            }
        }
    }
    return false;
}



//ע��   �������������

//0 url��http://127.0.0.1/aa/4.asp?act=sdf&v=sdf
//1 urlDir��http://127.0.0.1/
//2 fileName��4.asp
//3 FileType��asp
//4 fileStr��4.asp?act=sdf&v=sdf
//5 HttpAgreement��http
//6 webSite��http://127.0.0.1/
//7 folderDir��aa/

//��ַ���������20150124  ����  0ԭ�ļ�·�� 1Ϊ�ļ�·��   2Ϊ�ļ�����  3Ϊȥ���ļ������ļ�����   4Ϊ�ļ����ͺ�׺��
string[] handleHttpUrlArray( string url){
    //on error resume next
    string urlDir=""; string fileName=""; string fileType=""; string fileStr=""; string httpAgreement=""; string webSite=""; string folderDir="";
    url= handleHttpUrl(url);

    urlDir= mid(url, 1, inStrRev(url, "/"));
    fileStr= mid(url, inStrRev(url, "/") + 1,-1) ; fileName= fileStr;
    if( inStr(fileStr, "?") > 0 ){
        fileName= mid(fileStr, 1, inStr(fileStr, "?") - 1);
    }
    fileType= mid(fileName, inStrRev(fileName, ".") + 1,-1);
    httpAgreement= mid(url, 1, inStr(url, ":") - 1);
    webSite= getWebSite(url);
    //Call echo("url", url)
    //����Ϊ���򷢴ֻ���ļ���Ŀ¼20160613
    if( webSite != "" ){
        folderDir= mid(urlDir, len(webSite),-1);
    }else{
        echoYellowB("ע�⣺������Ч��ַ", url);
    }
    //HandleHttpUrlArray = Array(url, urlDir, fileName, fileType, fileStr, HttpAgreement, webSite, folderDir)
    string[] arrayData;
    arrayData= aspSplit(url + vbCrlf() + urlDir + vbCrlf() + fileName + vbCrlf() + fileType + vbCrlf() + fileStr + vbCrlf() + httpAgreement + vbCrlf() + webSite + vbCrlf() + folderDir, vbCrlf());
    return arrayData;
}


//�Ƴ�jsCss��Ĳ���Param (20151019)
string remoteJsCssParam(string content, string pubAHrefList){
    return handleRemoteJsCssParam(content, pubAHrefList, "|�滻����|�滻��ַ");
}

//�����Ƴ�jsCss��Ĳ���Param (20151019)
string handleRemoteJsCssParam(string content, string urlList, string sType){
    string[] splStr; string c=""; string url=""; string[] arrData; string fileName=""; string fileType=""; string fileStr=""; string replaceStr="";
    sType= "|" + sType + "|";
    splStr= aspSplit(urlList, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        if( url != "" ){
            if( c != "" ){ c= c + vbCrlf() ;}
            arrData= handleHttpUrlArray(url);
            fileName= arrData[2];
            fileType= lCase(arrData[3]);
            fileStr= lCase(arrData[4]);
            if((fileType== "js" || fileType== "css") && inStr(fileStr, "?") > 0 && inStr(fileName, ".") > 0 ){
                replaceStr= mid(url, 1, inStr(url, "?") - 1);
                //call echo(replaceStr,fileStr)
                //�����滻�������ǲ���׼�����Ľ�
                if( inStr(sType, "|�滻����|") > 0 ){
                    content= replace(content, url, replaceStr);
                }
                if( inStr(sType, "|�滻��ַ|") > 0 ){
                    urlList= replace(urlList, url, replaceStr);
                }
            }
        }
    }
    return "";
}


//����������ַ����(20151022)
string batchHandleHttpUrlComplete(string httpurl, string content){
    string webSite=""; string[] splStr; string url=""; string lCaseUrl=""; string c="";
    webSite= getWebSite(httpurl);
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        url= phpTrim(url);
        lCaseUrl= lCase(url);
        if( lCaseUrl != "#" && left(lCaseUrl, 11) != "javascript:" ){
            if( inStr(vbCrlf() + lCase(c) + vbCrlf(), vbCrlf() + lCaseUrl + vbCrlf())== 0 ){
                if( c != "" ){ c= c + vbCrlf() ;}
                c= c + urlAddHttpUrl(webSite, url);
            }
        }
    }
    return c;
}


//���ͬ����(20151023)
bool isWebSite( string url1, string url2){
    return handleIsWebSite(url1, url2, "");
}
//���ͬ������(20151023)
bool isSonWebSite( string url1, string url2){
    return handleIsWebSite(url1, url2, "������");
}

//��������ַ�Ƿ�����ͬ��(20151023)
bool handleIsWebSite( string url1, string url2, string sType){
    url1= getWebSite(url1);
    url2= getWebSite(url2);
    if( inStr(url1, "://") > 0 ){
        url1= mid(url1, inStr(url1, "://") + 3,-1);
    }
    if( left(url1, 4)== "www." ){
        url1= mid(url1, 5,-1);
    }
    if( inStr(url2, "://") > 0 ){
        url2= mid(url2, inStr(url2, "://") + 3,-1);
    }
    if( left(url2, 4)== "www." ){
        url2= mid(url2, 5,-1);
    }

    if( sType== "������" ){
        string[] splStr; string s=""; string c="";
        c= ".com.hk|.sh.cn|.com.cn|.net.cn|.org.cn";
        c= c + "|.com|.net|.org|.tv|.cc|.info|.cn|.tw|:81|:99|.biz|.mobi|.hk|.us|.la|.gl|.in";
        splStr= aspSplit(c, "|");
        foreach(var eachs in splStr){
            s=eachs;
            if( s != "" ){
                url1= replace(url1, s, "");
                url2= replace(url2, s, "");
            }
        }

        if( inStr(url1, ".")>0 ){
            url1= mid(url1, inStr(url1, ".") + 1,-1);
        }
        if( inStr(url2, ".")>0 ){
            url2= mid(url2, inStr(url2, ".") + 1,-1);
        }


    }
    bool handleIsWebSite= false;
    if( url1== url2 ){
        handleIsWebSite= true;
    }
    return handleIsWebSite;
}


//�����������ַ�б�(20161025)
string getContentUrlList(string httpurl, string content){
    return handleGetContentUrlList(httpurl, content, "|*|����|");
}
//��������������ַ�б�(20161025)
string handleGetContentUrlList(string httpurl, string content, string sType){
    int i=0; string s=""; string sNext=""; string sEndLCase=""; string sEnd=""; string urlStr=""; int nLen=0; string urlList=""; string url=""; string urlLCase=""; string webSite=""; string labelType=""; bool isHandle; string valueLabel="";
    sType= "|" + lCase(aspTrim(sType)) + "|";
    webSite= getWebSite(httpurl);
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        sNext= mid(content + " ", i + 1, 1);
        sEnd= mid(content, i + 1,-1) ; sEndLCase= lCase(sEnd);
        if( s== "<" ){
            url= "";
            labelType= "";
            isHandle= false;
            if( left(sEndLCase, 2)== "a " ){
                labelType= "a";
                valueLabel= "href";
                isHandle= true;
            }else if( left(sEndLCase, 5)== "link " ){
                labelType= "link";
                valueLabel= "href";
                isHandle= true;
            }else if( left(sEndLCase, 4)== "img " ){
                labelType= "img";
                valueLabel= "src";
                isHandle= true;
            }else if( left(sEndLCase, 7)== "script " ){
                labelType= "script";
                valueLabel= "src";
                isHandle= true;
            }
            if( isHandle== true ){
                if( inStr(sType, "|" + labelType + "|") > 0 || inStr(sType, "|*|") > 0 ){
                    nLen= inStr(sEnd, ">");
                    urlStr= mid(sEnd, 1, nLen);
                    url= rParam(urlStr, valueLabel);
                    i= i + nLen;
                }
            }
            if( url != "" ){
                urlLCase= lCase(url);
                if( urlLCase != "#" && left(urlLCase, 11) != "javascript:" ){
                    if( inStr(vbCrlf() + urlList + vbCrlf(), vbCrlf() + url + vbCrlf())== 0 ){
                        url= fullHttpUrl(httpurl, url);
                        isHandle= isSonWebSite(url, httpurl);
                        if( inStr(sType, "|����|") > 0 ){
                            if( isHandle== true ){
                                urlList= urlList + url + vbCrlf();
                            }
                        }else if( inStr(sType, "|����|") > 0 ){
                            if( isHandle== false ){
                                urlList= urlList + url + vbCrlf();
                            }
                        }else{
                            urlList= urlList + url + vbCrlf();
                        }
                    }
                }
            }
        }
    }
    if( urlList != "" ){ urlList= left(urlList, len(urlList) - 2); }
    return urlList;
}
//�����ַ�������������б�
string getInChain(string httpurl, string urlList){
    string[] splStr; string url=""; string c=""; string urlLCase=""; bool isHandle;
    splStr= aspSplit(urlList, vbCrlf());
    urlList= "";
    foreach(var eachurl in splStr){
        url=eachurl;
        if( url != "" ){
            urlLCase= lCase(url);
            if( left(urlLCase, 1) != "#" && left(urlLCase, 11) != "javascript:" ){
                if( inStr(vbCrlf() + urlList + vbCrlf(), vbCrlf() + url + vbCrlf())== 0 ){
                    url= fullHttpUrl(httpurl, url);
                    isHandle= isSonWebSite(url, httpurl);
                    if( isHandle== true ){
                        urlList= urlList + url + vbCrlf();
                    }
                }
            }
        }
    }
    if( urlList != "" ){ urlList= left(urlList, len(urlList) - 2); }
    return urlList;
}


//����ɨ�����ַ�б� 20160428
string handleScanUrlList(string httpurl, string urlList){
    string[] splStr; string url=""; string c=""; string lCaseUrl="";
    splStr= aspSplit(urlList, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        url= phpTrim(url);
        lCaseUrl= lCase(url);
        if( url != "" && left(url, 10) != "tencent://" && left(url, 11) != "javascript:" && left(url, 1) != "#" ){
            url= fullHttpUrl(httpurl, url);
            if( inStr(vbCrlf() + c + vbCrlf(), vbCrlf() + url + vbCrlf())== 0 ){
                c= c + url + vbCrlf();
            }
        }
    }
    return c;
}

//������ͬ�����б� 20160501
string handleWithWebSiteList(string httpurl, string urllist){
    string webSite=""; string[] splStr; string url=""; string c=""; string urlWebsite=""; string s="";
    webSite= lCase(getWebSite(httpurl));
    splStr= aspSplit(urllist, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        if( url != "" ){
            if( right(url, 1) != "/" && inStr(url, "?")== 0 ){
                s= mid(url, inStrRev(url, "/"),-1);
                //call echo("s",s)
                if((inStr(s, ".")== 0 || inStr(s, ".com") > 0 || inStr(s, ".cn") > 0 || inStr(s, ".net") > 0) && inStr(s, "@")== 0 ){
                    url= url + "/";
                }
                //call echo("url",url)
            }
            urlWebsite= lCase(getWebSite(url));
            if( webSite== urlWebsite && inStr(vbCrlf() + c + vbCrlf(), vbCrlf() + url + vbCrlf())== 0 ){
                if( c != "" ){
                    c= c + vbCrlf();
                }
                c= c + url;
            }
        }
    }
    return c;
}
//����ͬ�����б� 20160501
string handleDifferenceWebSiteList(string httpurl, string urllist){
    string webSite=""; string[] splStr; string url=""; string c=""; string urlWebsite=""; string websiteList="";
    webSite= lCase(getWebSite(httpurl));
    splStr= aspSplit(urllist, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        urlWebsite= lCase(getWebSite(url));
        if( urlWebsite != "" && webSite != urlWebsite && inStr(vbCrlf() + websiteList + vbCrlf(), vbCrlf() + urlWebsite + vbCrlf())== 0 ){
            websiteList= websiteList + urlWebsite + vbCrlf();
        }
    }
    return websiteList;
}

//����������� 20160511   ����checkDomainName('http://www.baidu.com/a/b/sdf')
bool checkDomainName(string httpurl){
    string url=""; string url2=""; string s=""; string s2="";
    url= getWebSite(httpurl);
    url2= url + "/a/1/b/2/cdefg/";
    //call echo(url, url2)
    bool checkDomainName= true;
    s= getHttpUrl(url, "");
    s2= getHttpUrl(url2, "");
    if( s== s2 ){
        checkDomainName= false;
    }
    return checkDomainName;
}

//���url״̬˵��
string getHttpUrlStateAbout(int nState){
    string s=""; string c="";
    s= cStr(aspTrim(nState));
    switch ( s ){
        case "100" : c= "����";break;
        case "101" : c= "����Э��";break;
        case "200" : c= "�ɹ�";break;
        case "201" : c= "����";break;
        case "202" : c= "����";break;
        case "203" : c= "��Ȩ����Ϣ";break;
        case "204" : c= "��������";break;
        case "205" : c= "��������";break;
        case "206" : c= "��������";break;
        case "300" : c= "����ѡ��";break;
        case "301" : c= "�ƶ�����";break;
        case "302" : c= "��ʱ�ƶ�";break;
        case "303" : c= "��������";break;
        case "304" : c= "δ�޸�";break;
        case "305" : c= "ʹ�ô���";break;
        case "307" : c= "��ʱ�ض���";
        break;
        case "400" : c= "����Ҫ��";break;
        case "401" : c= "δ����Ȩ";break;
        case "402" : c= "����Ҫ��";break;
        case "403" : c= "��ֹ";break;
        case "404" : c= "δ�ҵ�";break;
        case "405" : c= "������ķ���";break;
        case "406" : c= "���ɽ���";break;
        case "407" : c= "������֤����";break;
        case "408" : c= "����ʱ";break;
        case "409" : c= "��ͻ";break;
        case "410" : c= "��ʧ";break;
        case "411" : c= "���賤��";break;
        case "412" : c= "�Ⱦ�����";break;
        case "413" : c= "����ʵ�����";break;
        case "414" : c= "������URI̫��";break;
        case "415" : c= "��֧�ֵ�ý������";break;
        case "416" : c= "������Χ������";break;
        case "417" : c= "����ʧ��";
        break;
        case "500" : c= "�ڲ�����������";break;
        case "501" : c= "δʵʩ";break;
        case "502" : c= "������";break;
        case "503" : c= "���񲻿���";break;
        case "504" : c= "���س�ʱ";break;
        case "505" : c= "��HTTP�汾��֧��";break;
        case "509" : c= "�������Ƴ���";
        break;
    }
    return c;
}

</script>
