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
//���ܽ���(2014)


//����Html�ϴ����ܽ��� 20150121 specialHtmlUploadEncryptionDecrypt(Content,"Decrypt")
string specialHtmlUploadEncryptionDecrypt(string content, string sType){
    string[] splStr; string[] splxx; string c=""; string s="";
    c= "��|[*-24156*]" + vbCrlf();
    splStr= aspSplit(c, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "|") > 0 ){
            splxx= aspSplit(s, "|");
            if( sType== "1" || sType== "����" || sType== "Decrypt" ){
                content= replace(content, splxx[1], splxx[0]);
            }else{
                content= replace(content, splxx[0], splxx[1]);
            }
        }
    }
    return content;
}

//����ASP��������
string encAspContent( string content){
    string[] splStr; string c=""; string s=""; string THStr="";
    c= "Str=Str&\"|Str=Str & |If | Then|End If|&vbCrlf|Temp |Rs(|Rs.|.AddNew|(\"Title\")|(\"Content\")|=False|ElseIf|";
    c= c + "Conn.Execute(\"| Exit For|[Product]|.Open|.Close|Exit For|Exit Function|MoveNext:Next:|Str ";
    splStr= aspSplit(c, "|");
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            THStr= upperCaseORLowerCase(s);
            content= replace(content, chr(9).ToString(), ""); //Chr(9) = Tab
            content= replace(content, s, THStr);
        }
    }
    return content;
}
//�ô�Сд�ҵ�
string upperCaseORLowerCase( string content){
    int i=0; string s=""; string c=""; int nRnd=0;
    c= "";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);

        nRnd= cInt(rnd() * 1);
        if( nRnd== 0 ){
            c= c + lCase(s);
        }else{
            c= c + uCase(s);
        }
    }
    return c;
}
//����  Encryption
string encCode( string content){
    int i=0; string c="";
    c= "";
    for( i= 1 ; i<= len(content); i++){
        c= c + "%" + hex(asc(mid(content, i, 1)));
    }
    return c;
}
//����  Decrypt
string decCode( string content){
    int i=0; string c=""; string[] splStr;string s="";
    c= "";
    splStr= aspSplit(content, "%");
    for( i= 1 ; i<= uBound(splStr); i++){
        if( splStr[i] != "" ){
            s="&H" + splStr[i];
            c= c + chr(cInt(s).ToString());
        }
    }
    return c;
}
//�����ֵ�ת��Ϊ&#��ͷ��unicode�ַ�����ʽ
string toUnicode(string str){
    int i=0; int j=0; string c=""; string p="";
    string toUnicode= "";
    c= "";
    p= "";
    for( i= 1 ; i<= len(str); i++){
        c= mid(str, i, 1);
        j= ascW(c);
        if( j < 0 ){
            j= j + 65536;
        }
        if( j >= 0 && j <= 128 ){
            if( p== "c" ){
                toUnicode= " " + toUnicode;
                p= "e";
            }
            toUnicode= toUnicode + c;
        }else{
            if( p== "e" ){
                toUnicode= toUnicode + " ";
                p= "c";
            }
            toUnicode= toUnicode + "&#" + j + ";";
        }
    }
    return toUnicode;
}
//����26��ĸ����
string japan( string iStr, string sType){
    string japan= "";
    if( isNull(iStr) || isEmpty(iStr) ){
        japan= "";
        return japan;
    }
    string[] arrF; int i=0; string[] arrE;
    if( sType== "" ){ sType= "0";}
    //F=array("��","��","��","��","��","��","��","��","��","��",_
    //"��","��","��","��","��","��","��","��","��","��","��","��",_
    //"��","��","��","��")
    //E = Array("Jn0;", "Jn1;", "Jn2;", "Jn3;", "Jn4;", "Jn5;", "Jn6;", "Jn7;", "Jn8;", "Jn9;", "Jn10;", "Jn11;", "Jn12;", "Jn13;", "Jn14;", "Jn15;", "Jn16;", "Jn17;", "Jn18;", "Jn19;", "Jn20;", "Jn21;", "Jn22;", "Jn23;", "Jn24;", "Jn25;")
    arrE= aspSplit("Jn0;,Jn1;,Jn2;,Jn3;,Jn4;,Jn5;,Jn6;,Jn7;,Jn8;,Jn9;,Jn10;,Jn11;,Jn12;,Jn13;,Jn14;,Jn15;,Jn16;,Jn17;,Jn18;,Jn19;,Jn20;,Jn21;,Jn22;,Jn23;,Jn24;,Jn25;", ",");

    //F = Array(Chr( -23116), Chr( -23124), Chr( -23122), Chr( -23120),    Chr(-23118), Chr( -23114), Chr( -23112), Chr( -23110),     Chr(-23099), Chr( -23097), Chr( -23095), Chr( -23075),   Chr(-23079), Chr( -23081), Chr( -23085), Chr( -23087),  Chr(-23052), Chr( -23076), Chr( -23078), Chr( -23082),  Chr(-23084), Chr( -23088), Chr( -23102), Chr( -23104), Chr(-23106), Chr( -23108))
    arrF= aspSplit(chr( -23116).ToString() + "," + chr( -23124).ToString() + "," + chr( -23122).ToString() + "," + chr( -23120).ToString() + "," + chr( -23118).ToString() + "," + chr( -23114).ToString() + "," + chr( -23112).ToString() + "," + chr( -23110).ToString() + "," + chr( -23099).ToString() + "," + chr( -23097).ToString() + "," + chr( -23095).ToString() + "," + chr( -23075).ToString() + "," + chr( -23079).ToString() + "," + chr( -23081).ToString() + "," + chr( -23085).ToString() + "," + chr( -23087).ToString() + "," + chr( -23052).ToString() + "," + chr( -23076).ToString() + "," + chr( -23078).ToString() + "," + chr( -23082).ToString() + "," + chr( -23084).ToString() + "," + chr( -23088).ToString() + "," + chr( -23102).ToString() + "," + chr( -23104).ToString() + "," + chr( -23106).ToString() + "," + chr( -23108).ToString(), ",");
    japan= iStr;
    for( i= 0 ; i<= 25; i++){
        if( sType== "0" ){
            japan= replace(japan, arrF[i], arrE[i]);
        }else{
            japan= replace(japan, arrE[i], arrF[i]);
        }
    }
    return japan;
}
//����26��ĸ ����
string japan26(string iStr){
    return japan(iStr, "0");
}
//����26��ĸ ����
string unJapan26(string iStr){
    return japan(iStr, "1");
}
//��������������Ϊ��HTML����
string handleHTML( string content){
    //Content = Replace(Content, "&", "&amp;")
    content= replace(content, "<", "&lt;");
    return content;
}
//�⿪ ��������������Ϊ��HTML����
string unHandleHTML( string content){
    //Content = Replace(Content, "&amp;", "&")
    content= replace(content, "&lt;", "<");
    return content;
}
//Сд����   [����չΪ��д������]
string lcaseEnc(string str){
    int i=0; int n=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(str); i++){
        s= mid(str, i, 1);
        n= ascW(s);
        if( n >= 97 && n <= 122 ){
            c= c + chr(n + 1);
        }else{
            c= c + s;
        }
    }
    c= replace(c, vbCrlf(), "��");
    return c;
}
//Сд����
string lcaseDec(string str){
    int i=0; int n=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(str); i++){
        s= mid(str, i, 1);
        n= ascW(s);
        if( n >= 97 && n <= 123 ){
            c= c + chr(n - 1);
        }else{
            c= c + s;
        }
    }
    c= replace(c, "��", vbCrlf());
    return c;
}

//htmlת����js
string htmlToJs( string c){
    c= replace("" + c, "\\", "\\\\");
    c= replace(c, "/", "\\/");
    c= replace(c, "'", "\\'");
    c= replace(c, "\"", "\\\"");
    c= join(aspSplit(c, vbCrlf()), "\");" + vbCrlf() + "document.write(\"");
    c= "document.write(\"" + c + "\");";
    return c;
}
//jsת����html
string jsToHtml( string c){
    c= replace(c, "document.write(\"", "");
    c= replace(c, "\");", "");
    c= replace(c, "\\\"", "\"");
    c= replace(c, "\\'", "'");
    c= replace(c, "\\/", "/");
    c= replace(c, "\\\\", "\\");
    return c;
}
//htmlת����Asp
string htmlToAsp( string c){
    c= replace(c, "\"", "\"\"");
    c= join(aspSplit(c, vbCrlf()), "\")" + vbCrlf() + "Response.Write(\"");
    c= "Response.Write(\"" + c + "\")";
    return c;
}
//HtmlתAsp�����洢
string htmlToAspDim( string c){
    c= replace(c, "\"", "\"\"");
    c= join(aspSplit(c, vbCrlf()), "\"" + vbCrlf() + "C=C & \"");
    c= "C=C & \"" + c + "\"";
    return c;
}
//Aspת����html
string aspToHtml( string c){
    c= replace(c, "Response.Write(\"", "");
    c= replace(c, "\"\"", "\"");
    return c;
}
//�ļ���������
string setFileName( string fileName){
    int i=0; string s=""; string tStr=""; 
    //sArrayA = array("\", "/", ":", "*", "?", """", "<", ">", "|", ".", ",")       '�����ַ�����Ϊ����PHP��ͨ�� 20160511
    //sArrayB = array("Ʋ", "�X", "��", "��", "��", "��", "��", "��", "��", "��", "��")
    string[] sArrayA= {"\\", "/", ":", "*", "?", "\"", "<", ">", "|", ".", ","};
    string[] sArrayB= {"Ʋ", "�X", "��", "��", "��", "��", "��", "��", "��", "��", "��"};
    for( i= 0 ; i<= uBound(sArrayA); i++){
        s= sArrayA[i];
        tStr= sArrayB[i];
        fileName= replace(fileName, s, tStr);
    }
    fileName= replace(fileName, "&nbsp;", " ");
    fileName= replace(fileName, "&quot;", "˫");
    fileName= replace(fileName, vbCrlf(), "");
    return fileName;
}
//�ļ���������⿪
string unSetFileName( string fileName){
    int i=0; string s=""; string tStr=""; 
    string[] sArrayA= {"\\", "/", ":", "*", "?", "\"", "<", ">", "|", ".", ","};
    string[] sArrayB= {"Ʋ", "�X", "��", "��", "��", "��", "��", "��", "��", "��", "��"};
    for( i= 0 ; i<= uBound(sArrayA); i++){
        s= sArrayA[i];
        tStr= sArrayB[i];
        fileName= replace(fileName, tStr, s);
    }

    return fileName;
}

//��Htmlת��ASP�������ַ�ת��Chr(*)��ʽ
string HTMLToAscChr(string title){
    int i=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(title); i++){
        s= mid(title, i, 1);
        c= c + "Chr(" + asc(s) + ")&";
    }
    if( c != "" ){ c= left(c, len(c) - 1); }
    string HTMLToAscChr= c;
    //HTMLToAscChr = "<" & "%=" & C & "%" & ">"
    return HTMLToAscChr;
}
//����AscChr
string unHTMLToAscChr(string content){
    int i=0; string s=""; string c=""; string[] splStr; string temp="";
    c= content ; temp= "";
    c= replace(c, "Chr(", "");
    c= replace(c, ")&", " ");
    c= replace(c, ")", " ");
    splStr= aspSplit(c, " ");
    for( i= 0 ; i<= uBound(splStr) - 1; i++){
        s= splStr[i];
        temp= temp + chr(s).ToString();
    }
    return temp;
}

//������λ
string variableDisplacement(string content, int nPass){
    string c=""; int i=0; string s=""; string letterGroup=""; string digitalGroup=""; int nLetterGroup=0; int nDigitalGroup=0; int nLetterLen=0; int nDigitalLen=0; int nX=0;
    //��ĸ��
    //LetterGroup="abcdefghijklmnopqrstuvwxyz"
    letterGroup= "yzoehijklmfgqrstuvpabnwxcd";
    //��ĸ��
    nLetterGroup= len(letterGroup);
    //������
    //DigitalGroup="0123456789"
    digitalGroup= "4539671820";
    //���ֳ�
    nDigitalGroup= len(digitalGroup);
    c= "";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        nLetterLen= inStr(letterGroup, s);
        nDigitalLen= inStr(digitalGroup, s);
        //��ĸ����
        if( nLetterLen > 0 ){
            nX= nLetterLen + nPass;
            if( nX > nLetterGroup ){
                nX= nX - nLetterGroup;
            }else if( nX <= 0 ){
                //Call Echo("nX",nX & "," & (nLetterGroup - nX) & "/" & nLetterGroup)
                nX= nX + nLetterGroup;
            }
            s= mid(letterGroup, nX, 1);
            //���ִ���
        }else if( nDigitalLen > 0 ){
            nX= nDigitalLen + nPass;
            if( nX > nDigitalGroup ){
                nX= nX - nDigitalGroup;
            }else if( nX <= 0 ){
                //Call Echo("nX",nX & "," & (nLetterGroup - nX) & "/" & nLetterGroup)
                nX= nX + nDigitalGroup;
            }
            s= mid(digitalGroup, nX, 1);


        }
        c= c + s;
    }
    return c;
}


</script>
