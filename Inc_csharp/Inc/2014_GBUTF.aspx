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
//编码互换 GB2312与UTF-8转换

//UTF转GB23
string UTF2GB( string UTFStr){
    int nDig=0; string GBStr="";
    for( nDig= 1 ; nDig<= len(UTFStr); nDig++){
        //如果UTF8编码文字以%开头则进行转换
        if( mid(UTFStr, nDig, 1)== "%" ){
            //UTF8编码文字大于8则转换为汉字
            if( len(UTFStr) >= nDig + 8 ){
                GBStr= GBStr + convChinese(mid(UTFStr, nDig, 9));
                nDig= nDig + 8;
            }else{
                GBStr= GBStr + mid(UTFStr, nDig, 1);
            }
        }else{
            GBStr= GBStr + mid(UTFStr, nDig, 1);
        }
    }
    return GBStr;
}

//检测是否可以把UTF转GB2312
bool checkUTFToGB2312( string UTFStr){
    int nDig=0; string GBStr="";
    bool checkUTFToGB2312= true;
    for( nDig= 1 ; nDig<= len(UTFStr); nDig++){
        //如果UTF8编码文字以%开头则进行转换
        if( mid(UTFStr, nDig, 1)== "%" ){
            //UTF8编码文字大于8则转换为汉字
            if( len(UTFStr) >= nDig + 8 ){
                if( convChinese(mid(UTFStr, nDig, 9))== "[出错Error]" ){
                    return false ;
                }
                nDig= nDig + 8;
            }else{
                GBStr= GBStr + mid(UTFStr, nDig, 1);
            }
        }else{
            GBStr= GBStr + mid(UTFStr, nDig, 1);
        }
    }
    return checkUTFToGB2312;
}

//UTF8编码文字将转换为汉字 配合上面




//转UTF8编码，非常精准，好用，(2014 12 11)


//GB转UTF8--将GB编码文字转换为UTF8编码文字  这个不精准，等以后给废除掉


//GB转unicode---将GB编码文字转换为unicode编码文字
string chinese2unicode(string str){
    int i=0;
    string str_one="";
    string str_unicode="";
    if((isNull(str)) ){
        return "";
    }
    for( i= 1 ; i<= len(str); i++){
        str_one= mid(str, i, 1);
        str_unicode= str_unicode + chr(38).ToString();
        str_unicode= str_unicode + chr(35).ToString();
        str_unicode= str_unicode + chr(120).ToString();
        str_unicode= str_unicode + hex(ascW(str_one));
        str_unicode= str_unicode + chr(59).ToString();
    }
    return str_unicode;
}


//判断是否为有效的十六进制代码


//---------------------- 自己文件引用 ----------------------------
//作用同JS的escape一样
string escape(string str){
    int i=0; string s=""; string c=""; int n=0;
    s= "";
    for( i= 1 ; i<= len(str); i++){
        c= mid(str, i, 1);
        n= ascW(c);
        if((n >= 48 && n <= 57) ||(n >= 65 && n <= 90) ||(n >= 97 && n <= 122) ){
            s= s + c;
        }else if( inStr("@*_-./", c) > 0 ){
            s= s + c;
        }else if( n > 0 && n < 16 ){
            s= s + "%0" + hex(n);
        }else if( n >= 16 && n < 256 ){
            s= s + "%" + hex(n);
        }else{
            s= s + "%u" + hex(n);
        }
    }
    return s;
}


//作用同JS的unescape一样
string unescape(string str){
    int i=0; string sNew=""; string c=""; bool isChar; string lastChar=""; string sNext_1_c=""; int next_1_num=0;
    isChar= false;
    lastChar= "";
    sNew= "";
    for( i= 1 ; i<= len(str); i++){
        c= mid(str, i, 1);
        if( c== "+" ){
            sNew= sNew + " ";
        }else if( mid(str, i, 2)== "%u" && i <= len(str) - 5 ){
            if( is_numeric("&H" + mid(str, i + 2, 4)) ){
                sNew= sNew + chr(cInt("&H" + mid(str, i + 2, 4)));
                i= i + 5;
            }else{
                sNew= sNew + c;
            }
            //ElseIf c="%" and i<=Len(str)-2 Then
            //If IsNumeric("&H" & Mid(str,i+1,2)) Then
            //newstr = newstr & ChrW(CInt("&H" & Mid(str,i+1,2)))
            //i = i+2
            //Else
            //newstr = newstr & c
            //End If
        }else if( c== "%" && i <= len(str) - 2 ){
            sNext_1_c= mid(str, i + 1, 2);
            if( is_numeric("&H" + sNext_1_c) ){
                next_1_num= cInt("&H" + sNext_1_c);
                if( isChar== true ){
                    isChar= false;
                    sNew= sNew + chr(cInt("&H" + lastChar + sNext_1_c).ToString());
                }else{
                    if( abs(next_1_num) <= 127 ){
                        sNew= sNew + chr(next_1_num).ToString();
                    }else{
                        isChar= true;
                        lastChar= sNext_1_c;
                    }
                }
                i= i + 2;
            }else{
                sNew= sNew + c;
            }
        }else{
            sNew= sNew + c;
        }
    }
    return sNew;
}

//汉字乱码如：（中国，&#x4E2D;&#x56FD;）
string chineseToUnicode(string str){
    int i=0; string s=""; string c="";
    for( i= 1 ; i<= len(str); i++){
        s= mid(str, i, 1);
        s= "&#x" + hex(ascW(s)) + ";";
        c= c + s;
    }
    return c;
}

//解汉字乱码
string unicodeToChinese(string content){
    string[] splStr; string s=""; string c="";
    splStr= aspSplit(content, ";");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "&#x") > 0 ){
            s= right(s, len(s) - 3);
            c= c + chr(cInt("&H" + s));
        }
    }
    return c;
}



//Url加码
string URLEncode( string strUrl){
    int i=0;
    string tempStr="";
    string URLEncode="";
    for( i= 1 ; i<= len(strUrl); i++){
        if( asc(mid(strUrl, i, 1)) < 0 ){
            tempStr= "%" + right(cStr(hex(asc(mid(strUrl, i, 1)))), 2);
            tempStr= "%" + left(cStr(hex(asc(mid(strUrl, i, 1)))), len(cStr(hex(asc(mid(strUrl, i, 1))))) - 2) + tempStr;
            URLEncode= URLEncode + tempStr;
        }else if((asc(mid(strUrl, i, 1)) >= 65 && asc(mid(strUrl, i, 1)) <= 90) ||(asc(mid(strUrl, i, 1)) >= 97 && asc(mid(strUrl, i, 1)) <= 122) ){
            URLEncode= URLEncode + mid(strUrl, i, 1);
        }else{
            URLEncode= URLEncode + "%" + hex(asc(mid(strUrl, i, 1)));
        }
    }
    return URLEncode;
}

//URL解码   这个好用
string URLDecode( string strUrl){
    int i=0;
    string URLDecode="";
    if( inStr(strUrl, "%")== 0 ){ return strUrl ; }
    for( i= 1 ; i<= len(strUrl); i++){
        if( mid(strUrl, i, 1)== "%" ){
            if( val("&H" + mid(strUrl, i + 1, 2)) > 127 ){
                URLDecode= URLDecode + chr(val("&H" + mid(strUrl, i + 1, 2) + mid(strUrl, i + 4, 2)).ToString());
                i= i + 5;
            }else{
                URLDecode= URLDecode + chr(val("&H" + mid(strUrl, i + 1, 2)).ToString());
                i= i + 2;
            }
        }else{
            URLDecode= URLDecode + mid(strUrl, i, 1);
        }
    }
    return URLDecode;
}

//GB2312到UTF-8


//加密解密URL
string encDecURL( string content, string sType){
    string yStr=""; string tStr=""; string temp=""; int i=0; string s=""; string c="";
    yStr= "abcdefghijklmnopqrstuvwxyz0123456789:/.=& ()%";
    tStr= "9ab1cdefghij234klmnopqrst678uvwxyz:/.05*-$[]@";
    if( sType== "解密" || sType== "0" ){
        temp= yStr;
        yStr= tStr;
        tStr= temp;
    }
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr(yStr, s) > 0 ){
            s= mid(tStr, inStr(yStr, s), 1);
        }
        c= c + s;
    }
    return c;
}

//加密URL (辅助)
string encURL(string content){
    return encDecURL(content, "加密");
}

//解密URL (辅助)
string decURL(string content){
    return encDecURL(content, "解密");
}

</script>
