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
//Check��֤ (2013,10,26)

//���URL�ļ������Ƿ����������:.js?  .css?  �÷� checkUrlFileNameParam("http://sdfsd.com/aaa.js","js|css|")
bool checkUrlFileNameParam(string httpurl, string sList){
    string url=""; string[] splStr; string searchStr="";
    url= lCase(httpurl);
    sList= lCase(sList);
    splStr= aspSplit(sList, "|");
    foreach(var eachsearchStr in splStr){
        searchStr=eachsearchStr;
        if( searchStr != "" ){
            searchStr= "." + searchStr + "?";
            //call echo("searchStr",searchStr)
            if( inStr(url, searchStr) > 0 ){
                bool checkUrlFileNameParam= true;
                return checkUrlFileNameParam;
            }
        }
    }
    return false;
}

//�Ǵ�д 20160105
bool isUCase(string content){
    int i=0; string s="";
    bool isUCase= true;
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", s)== 0 ){
            isUCase= false;
            return isUCase;
        }
    }
    return isUCase;
}
//��Сд 20160105
bool isLCase(string content){
    int i=0; string s="";
    bool isLCase= true;
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("abcdefghijklmnopqrstuvwxyz", s)== 0 ){
            isLCase= false;
            return isLCase;
        }
    }
    return isLCase;
}


//������


//���ַ�trueת������ falseת������
bool strToTrueFalse( string content){
    content= aspTrim(lCase(content));
    return IIF(content== "true", true, false);
}
//���ַ�trueת1���� falseת0����
int strTrueFalseToInt( string content){
    content= aspTrim(lCase(content));
    return IIF(content== "true", 1, 0);
}
//��黻��
bool checkVbCrlf(string content){
    bool checkVbCrlf= false;
    if( inStr(content, vbCrlf()) > 0 ){ checkVbCrlf= true ;}

    return checkVbCrlf;
}
//��黻��    ����
bool checkBr(string content){
    return checkVbCrlf(content);
}

//�ж�������ż
bool isParity(int numb){
    bool isParity= true;




    if( numb % 2== 0 ){
        isParity= false;
    }
    return isParity;
}
//���eval��ȷ��
bool checkEval(string content){//���պ���
    return false;
}
//********************************************************
//����SQL�Ƿ��ַ�����ʽ��html����
//********************************************************
string replace_SQLText(string fString){
    string replace_SQLText= "";
    if( isNull(fString) ){
        return replace_SQLText;
    }else{
        fString= aspTrim(fString);
        fString= replace(fString, "'", "''");
        fString= replace(fString, ";", "��");
        fString= replace(fString, "--", "��");
        fString= displayHtml(fString);
        replace_SQLText= fString;
    }
    return replace_SQLText;
}
//********************************************************
//����Ƿ��ⲿ�ύ����
//********************************************************
bool chkPost(){
    string server_v1=""; string server_v2="";
    bool chkPost= false;
    server_v1= cStr(cStr(Request.ServerVariables["HTTP_REFERER"]));
    server_v2= cStr(cStr(Request.ServerVariables["SERVER_NAME"]));
    echo(server_v1, server_v2);
    if( mid(server_v1, 8, len(server_v2)) != server_v2 ){
        chkPost= false;
    }else{
        chkPost= true;
    }
    return chkPost;
}
//Response.Write(IsMail("asdf@sdf.com"))
//��������֤����
bool isMail(string email){//���պ���
    return false;
}
//������֤�ڶ���
bool isValidEmail(string email){
    string[] splNames; string sName=""; int i=0; string c="";
    bool isValidEmail= true;
    splNames= aspSplit(email, "@");
    if( uBound(splNames) != 1 ){ return false ; }
    foreach(var eachsName in splNames){
        sName=eachsName;
        if( len(sName) <= 0 ){ return false ; }
        for( i= 1 ; i<= len(sName); i++){
            c= lCase(mid(sName, i, 1));
            if( inStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 && is_numeric(c) ){ return false ; }
        }
        if( left(sName, 1)== "." || right(sName, 1)== "." ){ return false ; }
    }
    if( inStr(splNames[1], ".") <= 0 ){ return false ; }
    i= len(splNames[1]) - inStrRev(splNames[1], ".");
    if( i != 2 && i != 3 ){ return false ; }
    isValidEmail=IIF(inStr(email, "..") > 0,false,true);


    return isValidEmail;
}
//���Ϊ��Ч�ַ�
bool isCode( string content){
    string c="";
    c= replace(content, " ", "");
    c= replace(c, chr(13).ToString(), "");
    c= replace(c, chr(10).ToString(), "");
    c= replace(c, vbTab(), "");
    return IIF(c != "", true, false);
}
//�����Ƿ�Ϊ����
bool checkNumber( string content){
    int i=0; string s="";
    bool checkNumber= true;
    if( len(content)== 0 ){ return false ; }
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", lCase(s))== 0 ){
            checkNumber= false;
            return checkNumber;
        }
    }
    return checkNumber;
}
//�����Ƿ�Ϊ���� (����)
bool isNumber( string content){
    return checkNumber(content);
}
//�����Ƿ�Ϊ��ĸ
bool checkABC( string content){
    int i=0; string s="";
    bool checkABC= true;
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("abcdefghijklmnopqrstuvwxyz", lCase(s))== 0 ){
            checkABC= false;
            return checkABC;
        }
    }
    return checkABC;
}
//����ַ����� �����������ַ�
int getLen(string content){
    int i=0; int nS=0; int n=0;
    n= 0;
    for( i= 1 ; i<= len(content); i++){
        nS= asc(mid(cStr(content), i, 1));
        if( nS < 0 ){
            n= n + 2;
        }else{
            n= n + 1;
        }
    }
    return n;
}
//�����Ƿ�Ϊʱ������
bool checkTime(string sTime){
    return IIF(isDate(sTime)== false, false, true);
}
//�ж��Ƿ�Ϊ��



//****************************************************
//��������FoundInArr
//��  �ã����һ������������Ԫ���Ƿ����ָ���ַ���
//ʱ  �䣺2011��10��13��
//��  ���� strArr
//strToFind
//strSplit
//����ֵ���ַ���
//��  �ԣ�SHtml=R_("{����}",Function FoundInArr(strArr, strToFind, strSplit))
//****************************************************
bool foundInArr(string strArr, string strToFind, string strSplit){
    string[] arrTemp; int i=0;
    bool foundInArr= false;
    if( inStr(strArr, strSplit) > 0 ){
        arrTemp= aspSplit(strArr, strSplit);
        for( i= 0 ; i<= uBound(arrTemp); i++){
            if( lCase(aspTrim(arrTemp[i]))== lCase(aspTrim(strToFind)) ){
                foundInArr= true ; break;
            }
        }
    }else{
        if( lCase(aspTrim(strArr))== lCase(aspTrim(strToFind)) ){ foundInArr= true ;}
    }
    return foundInArr;
}
</script>
