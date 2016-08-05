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
//���뻥�� GB2312��UTF-8ת��

//UTFתGB23
string UTF2GB( string UTFStr){
    int nDig=0; string GBStr="";
    for( nDig= 1 ; nDig<= len(UTFStr); nDig++){
        //���UTF8����������%��ͷ�����ת��
        if( mid(UTFStr, nDig, 1)== "%" ){
            //UTF8�������ִ���8��ת��Ϊ����
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

//����Ƿ���԰�UTFתGB2312
bool checkUTFToGB2312( string UTFStr){
    int nDig=0; string GBStr="";
    bool checkUTFToGB2312= true;
    for( nDig= 1 ; nDig<= len(UTFStr); nDig++){
        //���UTF8����������%��ͷ�����ת��
        if( mid(UTFStr, nDig, 1)== "%" ){
            //UTF8�������ִ���8��ת��Ϊ����
            if( len(UTFStr) >= nDig + 8 ){
                if( convChinese(mid(UTFStr, nDig, 9))== "[����Error]" ){
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

//UTF8�������ֽ�ת��Ϊ���� �������




//תUTF8���룬�ǳ���׼�����ã�(2014 12 11)


//GBתUTF8--��GB��������ת��ΪUTF8��������  �������׼�����Ժ���ϳ���


//GBתunicode---��GB��������ת��Ϊunicode��������
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


//�ж��Ƿ�Ϊ��Ч��ʮ�����ƴ���


//---------------------- �Լ��ļ����� ----------------------------
//����ͬJS��escapeһ��
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


//����ͬJS��unescapeһ��
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

//���������磺���й���&#x4E2D;&#x56FD;��
string chineseToUnicode(string str){
    int i=0; string s=""; string c="";
    for( i= 1 ; i<= len(str); i++){
        s= mid(str, i, 1);
        s= "&#x" + hex(ascW(s)) + ";";
        c= c + s;
    }
    return c;
}

//�⺺������
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



//Url����
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

//URL����   �������
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

//GB2312��UTF-8


//���ܽ���URL
string encDecURL( string content, string sType){
    string yStr=""; string tStr=""; string temp=""; int i=0; string s=""; string c="";
    yStr= "abcdefghijklmnopqrstuvwxyz0123456789:/.=& ()%";
    tStr= "9ab1cdefghij234klmnopqrst678uvwxyz:/.05*-$[]@";
    if( sType== "����" || sType== "0" ){
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

//����URL (����)
string encURL(string content){
    return encDecURL(content, "����");
}

//����URL (����)
string decURL(string content){
    return encDecURL(content, "����");
}

</script>
