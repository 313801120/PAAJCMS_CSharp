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


//�����滻�б� (ReplaceList�ĳ�handleReplaceList)Ϊ�˼���php����
string handleReplaceList( string content, string yStr, string tStr){
    string[] splYuan; string[] splTi; int i=0; string s="";
    splYuan= aspSplit(yStr, "|");
    splTi= aspSplit(tStr + "||||||||||||||||||||||||||||||||", "|");
    for( i= 0 ; i<= uBound(splYuan); i++){
        s= splYuan[i];
        if( s != "" ){
            content= replace(content, s, splTi[i]);
        }
    }
    return content;
}

//���ұ߿�ʼ��ȡ ���ã�ɾ����߿ո�
string leftSpace( string content, int nAdd){
    int i=0; int nCount=0;
    nCount= 0;
    for( i= len(content) ; i>= 1 ; i--){
        if( mid(content, i, 1) != " " ){ break; }
        nCount= nCount + 1;
    }
    if( nCount > nAdd && replace(content, " ", "") != "" ){
        content= left(content, len(content) - nCount + nAdd);
    }else if( nCount < nAdd && 1== 1 ){
        for( i= nCount ; i<= nAdd - 1; i++){
            content= content + " ";
        }
    }
    return content;
}
//���ұ߿�ʼ��ȡ ���ã�ɾ���ұ߿ո�
string rightSpace( string content, int nAdd){
    int i=0; int nCount=0;
    nCount= 0;
    for( i= 1 ; i<= len(content); i++){
        if( mid(content, i, 1) != " " ){ break; }
        nCount= nCount + 1;
    }
    if( nCount > nAdd ){
        content= right(content, len(content) - nCount + nAdd);
    }else if( nCount < nAdd ){
        for( i= nCount ; i<= nAdd - 1; i++){
            content= " " + content;
        }
    }
    return content;
}

//=============== ����Ϊ2013,11,1���
//���QQ
string getQQ( string content, int nOK){
    string[] splStr; int i=0; int j=0; string s=""; string c=""; string qQ=""; int nErr=0; bool isQQ;
    content= replace(replace(lCase(content), "&nbsp;", ""), " ", "");
    nOK= 0;
    if( inStr(content, "qq") > 0 ){
        splStr= aspSplit(content, "qq");
        for( i= 1 ; i<= uBound(splStr); i++){
            s= splStr[i];
            qQ= "" ; nErr= 14 ; isQQ= false;
            for( j= 1 ; j<= len(s); j++){
                if( inStr("0123456789", mid(s, j, 1))== 0 ){
                    if( isQQ== true ){ break; }//qq��ʼ�ۼ�ʱ���˳�
                    if( nErr== 0 ){ break; }//�������N�˳�
                    nErr= nErr - 1;
                    if( mid(s, j, 1)== "Ⱥ" ){ break; }//ΪQQȺ�˳�
                }else{
                    isQQ= true;
                    qQ= qQ + mid(s, j, 1);
                }
                if( j > 30 ){ break; }//j����20�˳�
            }
            if( len(qQ) >= 6 && len(qQ) <= 10 && inStr(vbCrlf() + c, vbCrlf() + qQ + vbCrlf())== 0 ){
                if( inStr(vbCrlf() + c, vbCrlf() + qQ + vbCrlf())== 0 ){
                    c= c + qQ + vbCrlf();
                    nOK= nOK + 1;
                }
            }
        }
    }
    return c;
}
//����ֻ�
string getTel( string content, int nOK){
    string[] splStr; int i=0; int j=0; string s=""; string c=""; string tel=""; int nErr=0; bool isTel;
    content= replace(replace(lCase(content), "&nbsp;", ""), "�ֻ�", "tel");
    nOK= 0;
    if( inStr(content, "tel") > 0 ){
        splStr= aspSplit(content, "tel");
        for( i= 1 ; i<= uBound(splStr); i++){
            s= splStr[i];
            tel= "" ; nErr= 14 ; isTel= false;
            for( j= 1 ; j<= len(s); j++){
                if( inStr("0123456789", mid(s, j, 1))== 0 ){
                    if( isTel== true ){ break; }//Tel��ʼ�ۼ�ʱ���˳�
                    if( nErr== 0 ){ break; }//�������N�˳�
                    nErr= nErr - 1;
                }else{
                    isTel= true;
                    tel= tel + mid(s, j, 1);
                }
                if( j > 30 ){ break; }//j����20�˳�
            }
            if( len(tel)== 11 ){
                if( inStr(vbCrlf() + c, vbCrlf() + tel + vbCrlf())== 0 ){
                    c= c + tel + vbCrlf();
                    nOK= nOK + 1;
                }
            }
        }
    }
    return c;
}
//�������
string getMail( string content, int nOK){
    string[] splStr; int i=0; int j=0; string s=""; string c=""; string mail=""; int nErr=0; bool isEMail;
    content= replace(replace(lCase(content), "&nbsp;", ""), "����", "mail");
    nOK= 0;
    if( inStr(content, "mail") > 0 ){
        splStr= aspSplit(content, "mail");
        for( i= 1 ; i<= uBound(splStr); i++){
            s= splStr[i];
            mail= "" ; nErr= 14 ; isEMail= false;
            for( j= 1 ; j<= len(s); j++){
                if( inStr("0123456789abcdefghijklmnopqrstuvwxyz.@", mid(s, j, 1))== 0 ){
                    if( isEMail== true ){ break; }//Mail��ʼ�ۼ�ʱ���˳�
                    if( nErr== 0 ){ break; }//�������N�˳�
                    nErr= nErr - 1;
                }else{
                    isEMail= true;
                    mail= mail + mid(s, j, 1);
                }
                if( j > 30 ){ break; }//j����20�˳�
            }
            if( inStr(mail, ".") > 0 && inStr(mail, "@") > 0 ){
                if( inStr(vbCrlf() + c, vbCrlf() + mail + vbCrlf())== 0 ){
                    c= c + mail + vbCrlf();
                    nOK= nOK + 1;
                }
            }
        }
    }
    return c;
}
//���ͼƬ�б�
string getImgStr(string httpurl, string content, int nOK){
    string[] splStr; int i=0; string c=""; string url=""; string urlList="";
    content= getIMG(content);
    splStr= aspSplit(content, vbCrlf());
    nOK= 0;
    foreach(var eachurl in splStr){
        url=eachurl;
        if( inStr("|" + urlList + "|", "|" + url + "|")== 0 ){
            if( left(url, 1)== "/" ||(inStr(url, "http://")== 0 && inStr(url, "www.")== 0) ){
                url= fullHttpUrl(httpurl, url);
            }
            nOK= nOK + 1;
            c= c + url + "<br>";
        }
    }
    return c;
}
//���Css�б�
string getCssStr(string httpurl, string content, int nOK){
    string[] splStr; int i=0; string c=""; string url=""; string urlList="";
    content= getCssHref(content);
    splStr= aspSplit(content, vbCrlf());
    nOK= 0;
    foreach(var eachurl in splStr){
        url=eachurl;
        if( inStr("|" + urlList + "|", "|" + url + "|")== 0 ){
            if( left(url, 1)== "/" ||(inStr(url, "http://")== 0 && inStr(url, "www.")== 0) ){
                url= fullHttpUrl(httpurl, url);
            }
            nOK= nOK + 1;
            c= c + url + "<br>";
        }
    }
    return c;
}
//���Js�б�
string getJsStr(string httpurl, string content, int nOK){
    string[] splStr; int i=0; string c=""; string url=""; string urlList="";
    content= getJsSrc(content);
    splStr= aspSplit(content, vbCrlf());
    nOK= 0;
    foreach(var eachurl in splStr){
        url=eachurl;
        if( inStr("|" + urlList + "|", "|" + url + "|")== 0 ){
            if( left(url, 1)== "/" ||(inStr(url, "http://")== 0 && inStr(url, "www.")== 0) ){
                url= fullHttpUrl(httpurl, url);
            }
            nOK= nOK + 1;
            c= c + url + "<br>";
        }
    }
    return c;
}


//�����ַ�б� Rw(GetUrlStr("", "����", Content, "", 0))
string getUrlStr(string httpurl, string sType, string content, string searchValue, int nOK){
    string hrefList=""; string[] splStr; string title=""; string url=""; string c=""; string s=""; string lcaseUrl=""; string urlList=""; string titleList=""; string sVal=""; string sVal2="";
    if( sType== "����" ){
        hrefList= getAURL(content);
    }else if( sType== "���ӱ���" ){
        hrefList= getATitle(content);
    }else{
        hrefList= getAURLTitle(content);
    }
    splStr= aspSplit(hrefList, vbCrlf());
    nOK= 0;
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            if( sType== "����" ){
                url= aspTrim(s);
                if( inStr(lCase(url), "javascript:")== 0 && url != "#" && url != "" ){
                    if( inStr("|" + urlList + "|", "|" + url + "|")== 0 ){
                        if( left(url, 1)== "/" ||(inStr(url, "http://")== 0 && inStr(url, "www.")== 0) ){
                            url= fullHttpUrl(httpurl, url);
                        }
                        nOK= nOK + 1;
                        c= c + url + vbCrlf();
                    }
                }
            }else if( sType== "���ӱ���" ){
                if( inStr("|" + titleList + "|", "|" + s + "|")== 0 ){
                    titleList= titleList + s + "|";
                    nOK= nOK + 1;
                    c= c + s + vbCrlf();
                }
            }else{
                if( inStr(s, "��_|-��") > 0 ){
                    url= mid(s, 1, inStr(s, "��_|-��") - 1) ; lcaseUrl= url;
                    title= mid(s, len(url) + 6,-1);
                    if( inStr(lCase(lcaseUrl), "javascript:")== 0 && url != "#" && url != "" ){
                        if( inStr("|" + urlList + "|", "|" + url + "|")== 0 ){
                            if( left(url, 1)== "/" ||(inStr(lcaseUrl, "http://")== 0 && inStr(lcaseUrl, "www.")== 0) ){
                                url= fullHttpUrl(httpurl, url);
                            }
                            nOK= nOK + 1;
                            c= c + url + "   &nbsp;  " + title + vbCrlf();
                        }
                    }
                }
            }
        }
    }
    if( searchValue != "" ){
        searchValue= replace(searchValue, " and ", " And ");
        splStr= aspSplit(searchValue, " And ");
        sVal= splStr[0];
        if( uBound(splStr) > 0 ){ sVal2= splStr[1] ;}
        splStr= aspSplit(c, vbCrlf()) ; nOK= 0 ; c= "";
        foreach(var eachs in splStr){
            s=eachs;
            if( sVal2 != "" ){
                if( inStr(s, sVal) > 0 && inStr(s, sVal2) > 0 ){
                    if( inStr(vbCrlf() + c + vbCrlf(), vbCrlf() + s + vbCrlf())== 0 ){
                        nOK= nOK + 1;
                        c= c + s + "<br>";
                    }
                }
            }else{
                if( inStr(s, sVal) > 0 ){
                    if( inStr(vbCrlf() + c + vbCrlf(), vbCrlf() + s + vbCrlf())== 0 ){
                        nOK= nOK + 1;
                        c= c + s + "<br>";
                    }
                }
            }
        }
    }
    return c;
}
//���Html�з�ҳ���� (�ɼ��õ�)
string getPageConfig(string httpurl, string content, string sType){
    string[] splStr; int i=0; string s=""; string s2=""; string c=""; string url=""; string tempUrl=""; string[] arrUrl=aspArray(99); string urlList=""; string pageUrl=""; int nLen=0; string[] splxx;int nIndex=0;int n3=0;
    content= getAURL(content);
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        if( url != "" ){
            tempUrl= url;
            if( inStr(lCase(url), "javascript:")== 0 && url != "#" ){
                url= replace(replace(replace(replace(replace(url, "0", ""), "1", ""), "2", ""), "3", ""), "4", "");
                url= replace(replace(replace(replace(replace(url, "5", ""), "6", ""), "7", ""), "8", ""), "9", "");
                if( sType== "ע��" && inStr(url, "?") > 0 ){ url= handlSqlInUrl(url) ;}
                c= c + url + vbCrlf();
                nLen= inStr(vbCrlf() + urlList, vbCrlf() + url + "����");
                if( nLen > 0 ){
                    s= mid(urlList, nLen,-1);
                    s= mid(s, 1, inStr(s, vbCrlf()) - 1);
                    splxx= aspSplit(s, "����");
                    s2= splxx[0];
                    n3= cInt(splxx[1]) + 1;
                    urlList= replace(urlList, s, s2 + "����" + n3 + "����" + splxx[2]);
                    pageUrl= url;
                }else{
                    urlList= urlList + url + "����0����" + fullHttpUrl(httpurl, tempUrl) + vbCrlf();
                }
            }
        }
    }
    splStr= aspSplit(urlList, vbCrlf()) ; c= "";
    if( sType== "ע��" ){
        foreach(var eachs in splStr){
            s=eachs;
            if( inStr(s, "����") > 0 ){
                splxx= aspSplit(s, "����");
                url= aspTrim(splxx[2]);
                if( inStr(url, httpurl) > 0 ){
                    if( inStr(vbCrlf() + c, vbCrlf() + url + vbCrlf())== 0 ){
                        c= c + url + vbCrlf();
                    }
                }
            }
        }
        return c ;
    }
    //�����ҳ����
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "����") > 0 ){
            splxx= aspSplit(s, "����");
            nIndex=cInt(splxx[1]);
            if( nIndex > 0 ){ arrUrl[nIndex]= splxx[0] + "  &nbsp; | &nbsp; " + splxx[2] ;}
        }
    }
    for( i= 99 ; i>= 0 ; i--){
        if( arrUrl[i] != "" ){
            if( inStr(arrUrl[i], httpurl) > 0 ){
                c= c + arrUrl[i] + "��   ����[" + i + "]��<br>";
            }
        }
    }
    return c;
}
//�����ҳ��ע����ַ
string getSqlInUrl(string httpurl, string content, string sType){
    string[] splStr; int i=0; string s=""; string s2=""; string c=""; string url=""; string tempUrl=""; string[] arrUrl=aspArray(99); string urlList=""; string pageUrl=""; int nLen=0; string[] splxx;int n3=0;int nIndex=0;
    content= getAURL(content);
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachurl in splStr){
        url=eachurl;
        if( url != "" ){
            tempUrl= url;
            if( inStr(url, "?") > 0 ){
                c= c + url + vbCrlf();
                url= handlSqlInUrl(url);
                nLen= inStr(vbCrlf() + urlList, vbCrlf() + url + "����");
                if( nLen > 0 ){
                    s= mid(urlList, nLen,-1);
                    s= mid(s, 1, inStr(s, vbCrlf()) - 1);
                    splxx= aspSplit(s, "����");
                    s2= splxx[0];
                    n3= cInt(splxx[1]) + 1;
                    urlList= replace(urlList, s, s2 + "����" + n3 + "����" + splxx[2]);
                    pageUrl= url;
                }else{
                    urlList= urlList + url + "����0����" + fullHttpUrl(httpurl, tempUrl) + vbCrlf();
                }
            }
        }
    }
    splStr= aspSplit(urlList, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "����") > 0 ){
            splxx= aspSplit(s, "����");
            if( n3 > 0 ){
                nIndex=cInt(splxx[1]);
                if( sType== "ע��" ){
                    arrUrl[nIndex]= splxx[2];
                }else{
                    arrUrl[nIndex]= splxx[0] + "  &nbsp; | &nbsp; " + splxx[2];
                }
            }
        }
    }
    c= "";
    for( i= 99 ; i>= 0 ; i--){
        if( arrUrl[i] != "" ){
            if( sType== "ע��" ){
                c= c + arrUrl[i] + vbCrlf();
            }else{
                c= c + arrUrl[i] + "��   ����[" + i + "]��<br>";
            }
        }
    }
    return c;
}
//����ע����ַ�����û����վע����ַ
string handlSqlInUrl(string httpurl){
    string url=""; string[] splStr; int i=0; string s="";
    splStr= aspSplit(httpurl, "=");
    for( i= 0 ; i<= uBound(splStr); i++){
        if( i== uBound(splStr) ){ url= url + "=" ; break; }
        s= splStr[i];
        if( i % 2== 0 ){
            url= url + splStr[i];
        }else{
            if( inStr(s, "&") > 0 ){
                url= url + "=" + mid(s, inStr(s, "&"),-1);
            }
        }
    }
    return url;
}

//���ɨ����������б�
string getScanFunctionNameList(string content){

    string[] splStr; int i=0; bool isASP; bool isWord; string sx=""; string s=""; string wc=""; string zc=""; string s1=""; string aspCode=""; int nYHCount=0; string upWord=""; string funList=""; string s2="";
    string upWordn=""; string tempS=""; string dimList="";
    bool yesFunction; //�����Ƿ�Ϊ��
    isASP= false; //��ASP Ĭ��Ϊ��
    yesFunction= false; //�Ǻ��� Ĭ��Ϊ��
    isWord= false; //�ǵ��� Ĭ��Ϊ��
    nYHCount= 0; //˫����Ĭ��Ϊ0
    splStr= aspSplit(content, vbCrlf()); //�ָ���
    //ѭ������
    foreach(var eachs in splStr){
        s=eachs;
        //ѭ��ÿ���ַ�
        for( i= 1 ; i<= len(s); i++){
            sx= mid(s, i, 1);
            //Asp��ʼ
            if( sx== "<" && wc== "" ){ //����ı�����Ϊ�� WcΪ������� ��"<%" �ų� �޸���20140412
                if( mid(s, i + 1, 1)== "%" ){
                    isASP= true; //ASPΪ��
                    i= i + 1; //��1�����ܼ�2��Ҫ��Ȼ<%function Test() �ͽ�ȡ����
                }
                //ASP����
            }else if( sx== "%" && mid(s, i + 1, 1)== ">" && wc== "" ){ //WcΪ�������
                isASP= false; //ASPΪ��
                i= i + 1; //���ܼ�2��ֻ�ܼ�1����Ϊ���ﶨ��ASPΪ�٣���������һ����ʾ����� 'ASP����Ϊ��
            }
            if( isASP== true ){
                //�����ı�
                if((sx== "\"" || wc != "") ){
                    //˫�����ۼ�
                    if( sx== "\"" ){ nYHCount= nYHCount + 1 ;}
                    //�ж��Ƿ�"�����
                    if( nYHCount % 2== 0 ){
                        if( mid(s, i + 1, 1) != "\"" ){
                            wc= wc + sx;
                            s1= right(replace(mid(s, 1, i - len(wc)), " ", ""), 1); //����������Ҫ�������
                            if( yesFunction== true ){ aspCode= aspCode + wc ;}//���������ۼ�
                            nYHCount= 0 ; wc= ""; //���
                        }else{
                            wc= wc + sx;
                        }
                    }else{
                        wc= wc + sx;
                    }
                }else if( sx== "'" ){ //ע�����˳�
                    if( yesFunction== true ){ aspCode= aspCode + mid(s, i,-1) ;}
                    break;
                    //��ĸ
                }else if( checkABC(sx)== true ||(sx== "_" && zc != "") || zc != "" ){
                    zc= zc + sx;
                    s1= lCase(mid(s + " ", i + 1, 1));
                    if( inStr("abcdefghijklmnopqrstuvwxyz0123456789", s1)== 0 && (s1== "_" && zc != "") ){//����ж�
                        tempS= mid(s, i + 1,-1);
                        if( inStr("|function|sub|", "|" + lCase(zc) + "|")>0 ){
                            //������ʼ
                            if( yesFunction== false && lCase(upWord) != "end" ){
                                yesFunction= true;
                                s2= mid(s, i + 2,-1);
                                s2= mid(s2, 1, inStr(s2, "(") - 1);
                                funList= funList + s2 + vbCrlf();
                            }else if( yesFunction== true && lCase(upWord)== "end" ){ //�����һ������
                                aspCode= aspCode + zc + vbCrlf();

                                yesFunction= false;
                            }

                        }
                        upWord= zc; //��ס��ǰ����
                        if( yesFunction== true ){ aspCode= aspCode + zc ;}
                        zc= "";
                    }

                }
            }
            doEvents();
        }

    }
    return funList;
}
//��ú�������20150402
string getFunName( string c){
    if( inStr(c, "(") > 0 ){
        c= mid(c, 1, inStr(c, "(") - 1);
        c= phpTrim(c);
    }
    return c;
}
//��ú��������б���
string getFunDimName( string c){
    string startStr=""; string endStr=""; string s="";
    c= lCase(c);
    startStr= "(";
    endStr= ")";
    if( inStr(c, startStr) > 0 && inStr(c, endStr) > 0 ){
        c= strCut(c, startStr, endStr, 2);
    }
    if( c != "" ){
        c= replace(replace(c, "byref ", ""), "byref,", "");
        c= replace(replace(c, "byval ", ""), "byval,", "");
        c= replace(c, " ", "");
    }
    return c;
}
//��ñ������������
string getDimName( string c){
    string startStr=""; string endStr=""; string s="";
    c= lCase(c);
    startStr= ":";
    if( inStr(c, startStr) > 0 ){
        c= mid(c, 1, inStr(c, ":") - 1);
    }
    if( c != "" ){
        c= replace(replace(c, "byref ", ""), "byref,", "");
        c= replace(replace(c, "byval ", ""), "byval,", "");
        c= replace(c, " ", "");
    }
    if( inStr(c, "'") > 0 ){
        c= phpTrim(mid(c, 1, inStr(c, "'") - 1)); //�h�������'ע�Ͳ���    20150330
    }
    return c;
}

//���JS�������������
string getVarName( string content){
    string[] splStr; int i=0; string s=""; string c="";
    content= lCase(content);
    if( inStr(content, ";") > 0 ){
        content= mid(content, 1, inStr(content, ";") - 1);
    }
    splStr= aspSplit(content, ",");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "=") > 0 ){
            s= mid(s, 1, inStr(s, "=") - 1);
        }
        s= aspTrim(s);
        if( c != "" ){
            c= c + ",";
        }
        if( s != "" ){
            c= c + s;
        }
    }
    //call rwend(c)
    return c;
}


//���Set������������� ��ʱ�ò���
string getSetName( string c){
    c= phpTrim(lCase(c));
    c= mid(c, 1, inStr(c, "=") - 1);
    return aspTrim(c);
}
//�滻����
string replaceDim( string content){
    string[] splStr; string s=""; string tempS=""; string c=""; string lCaseS=""; int nDimInTHNumb=0;
    splStr= aspSplit(content, ",");
    foreach(var eachs in splStr){
        s=eachs;
        s= aspTrim(s);
        lCaseS= lCase(s);
        if( s != "" ){
            //�Ա�����()����
            if( inStr(s, "(") > 0 ){
                s= mid(s, 1, inStr(s, "(") - 1);
            }
            nDimInTHNumb= inStr(lCase(replaceDimList), "," + lCaseS + "="); //�滻����
            if( nDimInTHNumb > 0 ){ //�滻����
                tempS= mid(replaceDimList, nDimInTHNumb + 1,-1);
                tempS= mid(tempS, 1, inStr(tempS, ",") - 1);
                tempS= mid(tempS, inStr(tempS, "=") + 1,-1);
                if( inStr(funDim + rootFunDim, "," + tempS + ",")== 0 ){
                    s= tempS;
                }
            }
            c= c + s + ",";
        }
    }
    if( c != "" ){ c= left(c, len(c) - 1); }
    return c;
}
//�滻��������(���Ż�����) ��ǿ��20141017
string replaceDim2( string dimList, string dimName){
    string ZD=""; string[] splStr; int i=0; string s=""; int nMod=0; int nInt=0; string c="";
    string replaceDim2= dimName;
    ZD= "abcdefghijklmnopqrstuvwxyz";
    dimName= lCase(dimName);
    splStr= aspSplit(dimList, ",");
    for( i= 0 ; i<= uBound(splStr); i++){
        s= splStr[i];
        if( s== dimName ){
            nMod=(i) % len(ZD) + 1;
            nInt= fix((i) / len(ZD));

            if( nMod != 0 ){
                c= c + mid(ZD, nMod, 1);
            }
            if( nInt != 0 ){
                c= c + copyStr(mid(ZD, nInt, 1), nInt);
            }
            //Call Echo(I,Len(ZD))
            //Call Echo("nMod",nMod)
            //Call Echo("nInt",nInt)
            //Call Echo("C",C)
            replaceDim2= c;
            return replaceDim2;
        }
    }
    return replaceDim2;
}
//�ҵ�ǰ�ļ����ظ�����������
string findFolderRepeatFunction(string folderPath){
    string filePath=""; string s=""; string c=""; string content=""; string funs=""; string funList=""; string allFunList=""; int nOK=0; int nErr=0; string[] splStr; string[] splxx; int nAllOK=0; int nAllErr=0; int nI=0;
    string errFunList=""; string allErrFunList="";
    handlePath(folderPath); //�������·��
    c= "�����ļ���" + folderPath + vbCrlf();
    content= getDirFileList(folderPath, "");
    splStr= aspSplit(content, vbCrlf()) ; nI= 0;
    foreach(var eachfilePath in splStr){
        filePath=eachfilePath;
        nI= nI + 1;
        s= nI + "��" + filePath;
        content= getFText(filePath);
        content= getScanFunctionNameList(content); //���ASP���������б�
        nOK= 0 ; nErr= 0 ; nAllOK= 0 ; nAllErr= 0 ; funList= "" ; errFunList= "" ; allErrFunList= "";
        splxx= aspSplit(content, vbCrlf());
        foreach(var eachfuns in splxx){
            funs=eachfuns;
            if( funs != "" ){
                if( inStr("|" + funList + "|", "|" + funs + "|")== 0 ){
                    funList= funList + funs + "|";
                    nOK= nOK + 1;
                }else{
                    errFunList= errFunList + funs + "|";
                    nErr= nErr + 1;
                }
                if( inStr("|" + allFunList + "|", "|" + funs + "|")== 0 ){
                    allFunList= allFunList + funs + "|"; //ȫ������
                    nAllOK= nAllOK + 1;
                }else{
                    allErrFunList= allErrFunList + funs + "|";
                    nAllErr= nAllErr + 1;
                }
            }
            doEvents();
        }
        //Call CreateFile("allfun.txt", AllFunList)
        c= c + s + "��������" + uBound(splStr) + 1 + "�����ظ�(" + nErr + "[" + errFunList + "])ȫ�������ظ�(" + nAllErr + "[" + allErrFunList + "])" + vbCrlf();
        doEvents();
    }
    return c;
}
//�ҵ�ǰ�ļ��ظ����� (����)
string findFileRepeatFunction(string filePath){
    return handleContentRepeatFunction(getFText("FilePath"), "2");
}
//�ҵ�ǰ�����ظ�����
string findContentRepeatFunction(string content){
    return handleContentRepeatFunction(content, "2");
}
//���������ظ������б� sTypeΪ0Ϊ����ʾ��1Ϊ��ʾ�����б�2Ϊ��ʾ�ظ������б�3Ϊ��ʾ�������ظ������б�
string handleContentRepeatFunction(string content, string sType){
    string c=""; string funs=""; string funList=""; int nOK=0; int nErr=0; string[] splxx; string errFunList="";

    content= getScanFunctionNameList(content); //���ASP���������б�
    nOK= 0 ; nErr= 0;
    splxx= aspSplit(content, vbCrlf());
    foreach(var eachfuns in splxx){
        funs=eachfuns;
        if( funs != "" ){
            if( inStr("|" + funList + "|", "|" + funs + "|")== 0 ){
                funList= funList + funs + "|";
                nOK= nOK + 1;
            }else{
                errFunList= errFunList + funs + vbCrlf();
                nErr= nErr + 1;
            }
        }
        doEvents();
    }
    c= "�ҵ���������" + uBound(splxx) + 1 + "�����ظ�(" + nErr + ")" + vbCrlf();
    //�����б�
    if( sType== "1" || sType== "3" ){
        c= c + vbCrlf() + "�����б�" + vbCrlf() + funList;
    }
    //�ظ������б�
    if( errFunList != "" &&(sType== "1" || sType== "3") ){
        c= c + vbCrlf() + "�ظ������б�" + vbCrlf() + errFunList;
    }
    return c;
}
//�滻�ַ�����2 �Լ�д��һ�� ���ִ�Сд
string replace2(string content, string searchStr, string replaceStr){
    string leftStr=""; string rightStr="";
    if( inStr(content, searchStr) > 0 ){
        leftStr= mid(content, 1, inStr(content, searchStr) - 1);
        rightStr= mid(content, len(leftStr) + len(searchStr) + 1,-1);
        content= leftStr + replaceStr + rightStr;
    }
    return content;
}
//�滻ȫ���ַ�����2 �Լ�д��һ�� ���ִ�Сд
string allReplace(string content, string searchStr, string replaceStr){
    string leftStr=""; string rightStr=""; int i=0;
    for( i= 1 ; i<= 99; i++){
        if( inStr(content, searchStr) > 0 ){
            leftStr= mid(content, 1, inStr(content, searchStr) - 1);
            rightStr= mid(content, len(leftStr) + len(searchStr) + 1,-1);
            content= leftStr + replaceStr + rightStr;
        }else{
            break;
        }
    }
    return content;
}
//�滻һ�Σ������ִ�Сд
string replaceOneNOLU(string content, string searchStr, string replaceStr){
    string leftStr=""; string rightStr=""; string lCaseContent="";
    searchStr= lCase(searchStr);
    lCaseContent= lCase(content);
    if( inStr(lCaseContent, searchStr) > 0 ){
        leftStr= mid(content, 1, inStr(lCaseContent, searchStr) - 1);
        rightStr= mid(content, len(leftStr) + len(searchStr) + 1,-1);
        content= leftStr + replaceStr + rightStr;
    }
    return content;
}
//�Ż�ASP���� ɾ�����ҿո�
string optimizeAspCode(string content){
    string[] splStr; string s=""; string c=""; int i=0;
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        s= trimVbTab(s);
        if( s != "" ){ c= c + trimVbTab(s) + vbCrlf() ;}
    }
    return c;
}

//��ȡ�����������
string cutDaKuoHao(string content){
    string cutDaKuoHao= cutFunctionvValue(content, "{", "}");

    return cutDaKuoHao;
}
//��ȡ������С����
string cutXianKuoHao(string content){
    return cutFunctionvValue(content, "(", ")");
}

//��ȡ����ֵ20150716
string cutFunctionvValue(string content, string startStr, string endStr){
    int n1=0; int n2=0; int i=0; string s=""; string c="";
    n1= 1 ; n2= 0;
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        c= c + s;
        if( s== startStr ){
            n1= n1 + 1;
        }else if( s== endStr ){
            n2= n2 + 1;
            if( n1== n2 ){
                break;
            }
        }
    }
    return c;
}
//����ַ�����������ִ���20150721
int getStrIntContentNumb(string content, string findStr){
    string[] splStr;
    int getStrIntContentNumb= 0;
    if( inStr(content, findStr) > 0 ){
        splStr= aspSplit(content, findStr);
        getStrIntContentNumb= uBound(splStr);
    }
    return getStrIntContentNumb;
}
</script>
