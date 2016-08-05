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
//Check验证 (2013,10,26)

//检测URL文件名称是否带参数，如:.js?  .css?  用法 checkUrlFileNameParam("http://sdfsd.com/aaa.js","js|css|")
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

//是大写 20160105
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
//是小写 20160105
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


//检测错误


//把字符true转真类型 false转假类型
bool strToTrueFalse( string content){
    content= aspTrim(lCase(content));
    return IIF(content== "true", true, false);
}
//把字符true转1类型 false转0类型
int strTrueFalseToInt( string content){
    content= aspTrim(lCase(content));
    return IIF(content== "true", 1, 0);
}
//检查换行
bool checkVbCrlf(string content){
    bool checkVbCrlf= false;
    if( inStr(content, vbCrlf()) > 0 ){ checkVbCrlf= true ;}

    return checkVbCrlf;
}
//检查换行    辅助
bool checkBr(string content){
    return checkVbCrlf(content);
}

//判断数字奇偶
bool isParity(int numb){
    bool isParity= true;




    if( numb % 2== 0 ){
        isParity= false;
    }
    return isParity;
}
//检测eval正确性
bool checkEval(string content){//留空函数
    return false;
}
//********************************************************
//过滤SQL非法字符并格式化html代码
//********************************************************
string replace_SQLText(string fString){
    string replace_SQLText= "";
    if( isNull(fString) ){
        return replace_SQLText;
    }else{
        fString= aspTrim(fString);
        fString= replace(fString, "'", "''");
        fString= replace(fString, ";", "；");
        fString= replace(fString, "--", "—");
        fString= displayHtml(fString);
        replace_SQLText= fString;
    }
    return replace_SQLText;
}
//********************************************************
//检查是否外部提交数据
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
//正则表达验证邮箱
bool isMail(string email){//留空函数
    return false;
}
//邮箱验证第二种
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
//检测为有效字符
bool isCode( string content){
    string c="";
    c= replace(content, " ", "");
    c= replace(c, chr(13).ToString(), "");
    c= replace(c, chr(10).ToString(), "");
    c= replace(c, vbTab(), "");
    return IIF(c != "", true, false);
}
//测试是否为数字
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
//测试是否为数字 (辅助)
bool isNumber( string content){
    return checkNumber(content);
}
//测试是否为字母
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
//获得字符长度 汉字算两个字符
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
//测试是否为时间类型
bool checkTime(string sTime){
    return IIF(isDate(sTime)== false, false, true);
}
//判断是否为空



//****************************************************
//函数名：FoundInArr
//作  用：检查一个数组中所有元素是否包含指定字符串
//时  间：2011年10月13日
//参  数： strArr
//strToFind
//strSplit
//返回值：字符串
//调  试：SHtml=R_("{测试}",Function FoundInArr(strArr, strToFind, strSplit))
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
