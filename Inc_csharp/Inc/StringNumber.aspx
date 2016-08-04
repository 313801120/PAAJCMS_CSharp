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

//判断是否为字符转义
bool isStrTransferred(string content){
    string[] splStr; int i=0; string s=""; int nCount=0;
    nCount= 0;
    bool isStrTransferred= false;
    for( i= 0 ; i<= len(content) - 1 ; i++){
        s= mid(content, len(content) - i, 1);
        if( s== "\\" ){
            nCount= nCount + 1;
        }else{
            isStrTransferred= IIF(nCount % 2== 1, true, false);
            return isStrTransferred;
        }
    }
    return isStrTransferred;
}

//计算比率，游戏开发中用到 20150601
int[] getBL(int nSetWidth, int nSetHeight, int nDanFuXianZhi){
    int[] arrNSplStr=aspNArray(3);
    int nWidthZheFu= 1; //宽正负
    int nHeightZheFu= 1; //高正负
    int nBFB=0; //百分比
    int nXXFBX=0; //每个百分比，因为要判断他不能超过10
    if( nSetWidth < 0 ){
        nSetWidth= nSetWidth * - 1;
        nWidthZheFu= -1;
    }
    if( nSetHeight < 0 ){
        nSetHeight= nSetHeight * - 1;
        nHeightZheFu= -1;
    }
    if( nSetWidth > nSetHeight ){
        nBFB= formatNumber(nSetWidth / nSetHeight, 2); ////长宽 百分比
        arrNSplStr[0]= nBFB;
        arrNSplStr[1]= 1;
    }else{
        nBFB= formatNumber(nSetHeight / nSetWidth, 2); ////高宽 百分比
        arrNSplStr[0]= 1;
        arrNSplStr[1]= nBFB;
    }
    //每步超出指定值，处理
    //if nBFB>=nDanFuXianZhi then
    nXXFBX= formatNumber(nDanFuXianZhi / nBFB, 2);
    arrNSplStr[0]= arrNSplStr[0] * nXXFBX;
    arrNSplStr[1]= arrNSplStr[1] * nXXFBX;
    //end if

    arrNSplStr[0]= arrNSplStr[0] * nWidthZheFu;
    arrNSplStr[1]= arrNSplStr[1] * nHeightZheFu;
    arrNSplStr[2]= nBFB;
    arrNSplStr[3]= getCountPage(nSetWidth, arrNSplStr[0]);
    //Call echo("page count 页数", arrNSplStr(3))
    arrNSplStr[3]= getCountStep(nSetWidth, nSetHeight, arrNSplStr[0], arrNSplStr[1], arrNSplStr[3]);

    return arrNSplStr;
}

//获得总步数
int getCountStep(int nWidthStep, int nHeightStep, int nWidthBL, int nHeightBL, int nCountPage){
    int i=0;
    int getCountStep= 0;
    if( nWidthStep < 0 ){
        nWidthStep= nWidthStep * - 1;
    }
    if( nHeightStep < 0 ){
        nHeightStep= nHeightStep * - 1;
    }
    if( nWidthBL < 0 ){
        nWidthBL= nWidthBL * - 1;
    }
    if( nHeightBL < 0 ){
        nHeightBL= nHeightBL * - 1;
    }
    for( i= nCountPage - 10 ; i<= nCountPage; i++){
        //call echo(i & "、nWidthBL*i>=nWidthStep",nWidthBL*i &">="&nWidthStep    & "   |  " & nHeightBL*i &">="& nHeightStep)
        if( nWidthBL * i >= nWidthStep || nHeightBL * i >= nHeightStep ){
            getCountStep= i;
            //call echo("getCountStep",getCountStep)
        }
    }
    return getCountStep;
}


//获得中文汉字内容
string getChina(string content){
    int i=0; string c=""; int j=0; string s="";
    for( i= 1 ; i<= len(content); i++){
        j= asc(mid(content, i, 1));
        s= mid(content, i, 1);
        //是汉字累加
        if( j < 0 ){
            if((j <= -22033 && j >= -24158)== false ){
                c= c + s;
            }
        }
    }
    return c;
}
//判断是否有中文
bool isChina(string content){
    int i=0; int j=0; string s="";
    bool isChina= false;
    for( i= 1 ; i<= len(content); i++){
        j= asc(mid(content, i, 1));
        s= mid(content, i, 1);
        //是汉字累加
        if( j < 0 ){
            if((j <= -22033 && j >= -24158)== false ){
                isChina= true;
                return isChina;
            }
        }
    }
    return isChina;
}
//判断是否有中文 (辅助)
bool checkChina(string content){
    return isChina(content);
}

//PHP里Rand使用20150212


//引用上面，为什么？因为我老写错这个单词
int phpRnd(int nMinimum, int nMaximum){
    return phpRand(nMinimum, nMaximum);
}


//删除重复内容  20141220
string deleteRepeatStr(string content, string sSplType){
    string[] splStr; string s=""; string c="";
    c= "";
    splStr= aspSplit(content, sSplType);
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            if( inStr(sSplType + c + sSplType, sSplType + s + sSplType)== 0 ){
                c= c + s + sSplType;
            }
        }
    }
    if( c != "" ){ c= left(c, len(c) - len(sSplType)); }
    return c;
}

//替换内容N次 20141220
string replaceN(string content, string yunStr, string replaceStr, int nNumb){
    int i=0;string sNumb="";
    sNumb= handleNumber(cStr(nNumb));			//为了给.netc用
    if( sNumb== "" ){
        nNumb= 1;
    }else{
        nNumb= cInt(sNumb);
    }
    for( i= 1 ; i<= nNumb; i++){
        content= replace(content, yunStr, replaceStr);
    }

    return content;
}

//asp日期补0函数   引用别人20141216
string fillZero(string content){
    string fillZero="";
    if( len(content)== 1 ){
        fillZero= "0" + content;
    }else{
        fillZero= content;
    }
    return fillZero;
}

//不分大小写替换，作者：小云，写于20140925 用法Response.Write(CaseInsensitiveReplace("112233aabbbccddee","b","小云你牛"))
string caseInsensitiveReplace(string content, string sCheck, string sReplace){
    int nStartLen=0; int nEndLen=0; string lowerCase=""; string startStr=""; string endStr=""; string c=""; int i=0;
    c= "";
    if( lCase(sCheck)== lCase(sReplace) ){
        string caseInsensitiveReplace= content;
    }
    lowerCase= lCase(content);
    for( i= 1 ; i<= 99; i++){
        if( inStr(lowerCase, sCheck) > 0 ){
            nStartLen= inStr(lowerCase, sCheck) - 1;
            startStr= left(content, nStartLen);
            nEndLen= nStartLen + len(sCheck) + 1;
            endStr= mid(content, nEndLen,-1);
            content= startStr + sReplace + endStr;
            //Call Echo(StartLen,EndLen)
            //Call Echo(StartStr,EndStr)
            //Call Echo("Content",Content)
            lowerCase= lCase(content);
        }else{
            break;
        }
    }
    return content;
}

//数组数字排序 (2013,10,1)
int[] array_Sort(int[] arrnArray){
    int i=0; int j=0; int nMinmaxSlot=0; int nMinmax=0; int nTemp=0;
    for( i= uBound(arrnArray) ; i>= 0 ; i--){
        nMinmax= arrnArray[i];
        nMinmaxSlot= 0;
        for( j= 1 ; j<= i; j++){
            if( arrnArray[j] > nMinmax ){
                nMinmax= arrnArray[j];
                nMinmaxSlot= j;
            }
        }
        if( nMinmaxSlot != i ){
            nTemp= arrnArray[nMinmaxSlot];
            arrnArray[nMinmaxSlot]= arrnArray[i];
            arrnArray[i]= nTemp;
        }
    }
    return arrnArray;
}

//处理Zip大小


////生成随机数
string getRnd( int nCount){

    char charS; int i=0; string c="";
    for( i= 1 ; i<= nCount; i++){
        if( i % 2== 0 ){
            charS= chr((57 - 48) * rnd() + 48); //0~9
        }else if( i % 3== 0 ){
            charS= chr((90 - 65) * rnd() + 65); //A~Z
        }else{
            charS= chr((122 - 97) * rnd() + 97); //a~z
        }
        c= c + charS;
    }
    return c;
}

//获得随机数，仿js(20150826)
string mathRandom(){
    int i=0; string c="";
    c= "";

    for( i= 1 ; i<= 16; i++){
        c= c + cInt(cInt(rnd() * 9));
    }
    return "0." + c;
}



//获得指定位数随机A到Z字符
string getRndAZ(int nCount){
    string zd=""; int i=0; string s=""; string c="";
    c= "" ; zd= "";

    zd= "abcdefghijklmnopqrstuvwxyz" + uCase(zd);
    for( i= 1 ; i<= nCount; i++){
        s= mid(zd, phpRnd(1, len(zd)), 1);
        c= c + s;
    }
    return c;
}

//获得随机数 （辅助上面）
string getRand( int nCount){
    return getRnd(nCount);
}

//拷贝内容N次 InputStr输入值  nMultiplier乘数和php里面一样2014 12 02
//如果 nMultiplier 被设置为小于等于0，函数返回空字符串。
string copyStrNumb( string inputStr, int nMultiplier){
    int i=0; string s="";
    if( nMultiplier > 0 ){
        s= inputStr;
        for( i= 1 ; i<= nMultiplier - 1; i++){
            inputStr= inputStr + s;
        }
    }else{
        inputStr= "";
    }
    return inputStr;
}

//拷贝内容N次  PHP里函数
string str_Repeat( string inputStr, int nMultiplier){
    return copyStrNumb(inputStr, nMultiplier);
}

//引用上面的
string copyStr(string inputStr, int nMultiplier){
    return copyStrNumb(inputStr, nMultiplier);
}

//内容加Tab
string contentAddTab( string content, int nNumb){
    return copyStr("    ", nNumb) + join(aspSplit(content, vbCrlf()), vbCrlf() + copyStr("    ", nNumb));
}

//删除最后指定字符20150228 Content=DeleteEndStr(Content,2)
string deleteEndStr(string content, int nLen){
    if( content != "" ){ content= left(content, len(content) - nLen); }
    return content;
}


//StringNumber (2013,9,27)
int toNumber( int n, int nD){
    return formatNumber(n, nD, - 1);
}

//处理成数字
string handleNumber( string content){
    int i=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", s) > 0 ){
            c= c + s;
        }
    }
    return c;
}

//字符串中提取数字 20150507
string strDrawInt( string content){
    return handleNumber(content);
}

//处理成数字 首字符可以是-符号
string getFirstNegativeNumber( string content){
    int i=0; string s=""; string c="";
    c= "";
    content= aspTrim(content);
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( s== "-" && c== "" ){
            c= c + s;
        }else if( inStr("0123456789", s) > 0 ){
            c= c + s;
        }
    }
    if( c== "" ){ c= "0";}
    return c;
}

//检测是否为数字类型
string checkNumberType( string content){
    return handleNumber(content);
}

//检测字符内容为数字类型
bool checkStrIsNumberType( string content){
    int i=0; string s="";
    bool checkStrIsNumberType= true;
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", s)== 0 ){
            checkStrIsNumberType= false;
            return checkStrIsNumberType;
        }
    }
    return checkStrIsNumberType;
}

//处理成数字类型
string handleNumberType( string content){
    int i=0; string s=""; string c="";
    c= "";
    content= aspTrim(content);
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( i== 1 && inStr("+-*/", left(content, 1)) > 0 ){
            c= c + s;
        }else if( i > 1 && s== "." ){
            c= c + s;
        }else if( inStr("0123456789", s) > 0 ){
            c= c + s;
        }
    }
    return c;
}

//获得数字 只单独获得数字 并且第一个字数不能为零0     20150322
string getNumber( string content){
    int i=0; string s=""; string c="";
    c= "";
    content= aspTrim(content);
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", s) > 0 ){
            if( c== "" && s== "0" ){ //待改进，因为现在脑子不够用了，就这么定敢20150322
            }else{
                c= c + s;
            }
        }
    }
    return c;
}

//检测是否为数字
bool checkNumb(string s){
    bool checkNumb= false;
    if( inStr("0123456789.", s) > 0 ){
        checkNumb= true;
    }
    return checkNumb;
}

//获得有小数点数字
string getDianNumb( string content){
    int i=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789.", s) > 0 ){
            c= c + s;
        }
    }
    return c;
}

//获得总页数
int getCountPage(int nCount, int nPageSize){
    //把负数转成正确进行计算20150502
    int nCountPage=0;
    if( nCount < 0 ){
        nCount= nCount * - 1;
    }
    if( nPageSize < 0 ){
        nPageSize= nPageSize * - 1;
    }
    nCountPage= fix(nCount / nPageSize);
    if( nCount % nPageSize > 0 ){ nCountPage= nCountPage + 1 ;}
    return nCountPage;
}

//获得处理后页数
int getPageNumb(int nRecordCount, int nPageSize){
    int n=0;
    n= cInt(nRecordCount / nPageSize);
    if( nRecordCount % nPageSize > 0 ){
        n= n + 1;
    }
    return n;
}

//处理获得采集总页数
string getCaiHandleCountPage(string content){
    content= delHtml(content);
    content= handleNumber(content);
    string getCaiHandleCountPage= "";
    if( len(content) < 10 ){
        getCaiHandleCountPage= right(content, 1);
    }else if( len(content) < 200 ){
        getCaiHandleCountPage= right(content, 2);
    }
    return getCaiHandleCountPage;
}

//获得采集排序总页数 20150312
string getCaiSortCountPage( string content){
    int i=0; string s="";
    string getCaiSortCountPage= "";
    content= delHtml(content);
    content= handleNumber(content);
    for( i= 1 ; i<= 30; i++){
        s= mid(content, 1, len(i));
        if( s== cStr(i) ){
            getCaiSortCountPage= cStr(i);
            //Call Echo(i,s)
            content= right(content, len(content) - len(i));
        }
    }
    return getCaiSortCountPage;
}

//最大与最小之间 Between the minimum and maximum
int minMaxBetween(int nMin, int nMax, int nValue){
    //nMin = CInt(nMin)                                                         '最小数
    //nMax = CInt(nMax)                                                         '最大数
    //nValue = CInt(nValue)                                                     '当前数
    int minMaxBetween= nMin;
    if( nMin > nMax ){
        minMaxBetween= nMax;
    }else if( nValue > nMin ){
        minMaxBetween= nValue;
        if( nValue > nMax ){
            minMaxBetween= nMax;
        }
    }
    return minMaxBetween;
}

//获得内容文件名称中类型  (在FSO文件里已经有这个功能了20141220)
string getStrFileType(string fileName){
    string c="";
    c= "";
    if( inStr(fileName, ".") > 0 ){
        c= lCase(mid(fileName, inStrRev(fileName, ".") + 1,-1));
        if( inStr(c, "?") > 0 ){
            c= mid(c, 1, inStr(c, "?") - 1);
        }
    }
    return c;
}

//将字符类型转成数字类型
int val( string s){
    int val= cInt(s);
    if( s + ""== "" || is_numeric(s) ){
        val= 0;
    }
    return val;
}

//返回字符串左边N个byte


//循环加缩进 AddIndent(Content,"    ")
string addIndent(string content, string indentStr){
    string[] splStr; string s=""; string c="";
    c= "";
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        c= c + indentStr + s + vbCrlf();
    }
    return trimVbCrlf(c);
}

//获得数字前字符 2014 12 12(作用是为夏文强切换分块服务的)
string getNumberBeforeStr(string content){
    int i=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", s) > 0 ){ break; }
        c= c + s;
    }
    return c;
}

//获得随机数 20141212
//用法response.write makePassword(6)
string makePassword( int nMaxLen){
    string strNewPass="" ;int n=0;
    int nWhatsNext=0; int nUpper=0; int nLower=0; int nCounter=0;

    strNewPass= "";
    for( nCounter= 1 ; nCounter<= nMaxLen; nCounter++){
        nWhatsNext= cInt((1 - 0 + 1) * rnd() + 0);
        if( nWhatsNext== 0 ){
            nUpper= 90;
            nLower= 65;
        }else{
            nUpper= 57;
            nLower= 48;
        }
        n=cInt( (nUpper - nLower + 1) * rnd() + nLower);
        strNewPass= strNewPass + chr(n).ToString();
    }
    return strNewPass;
}

//功能说明：生成随机字符串，包括大小写字母，数字，和其它符合，常用于干扰码。 20141212
//参数说明：nMin--干扰码最小长度，ends--干扰码最大长度
//用法'Response.Write rndCode(20, 330)
string rndCode( int nMin, int nMax){
    int nRndLen=0; int i=0;

    string rndCode= "";
    nRndLen= cInt(nMin * rnd() + nMax - nMin);
    for( i= 1 ; i<= nRndLen; i++){

        rndCode= rndCode + chr(cInt(127 * rnd() + 1).ToString());
    }
    return rndCode;
}

//获得随机手机号码 没什么意义，引用别人的  20141217
//例:CAll Rw(GetRandomPhoneNumber(41))
string getRandomPhoneNumber(int nCount){
    string s=""; string rndnum=""; int j=0; string c="" ;int n=0;
    c= "" ; rndnum= "";
    j= 1;
    while( j < nCount){

        while( len(rndnum) < 9 ){//产生随机数的个数
            n= cInt((57 - 48) * rnd() + 48);
            rndnum= rndnum + chr(n).ToString();
        }
        c= c + 13 + rndnum + vbCrlf();
        rndnum= "";
        j= j + 1;
    }
    if( c != "" ){ c= left(c, len(c) - 2); }
    return c;
}

//获得字符长度


//数组转字符串

//移除数字(20151022)
string remoteNumber(string content){
    int i=0; string s=""; string c="";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789.", s)== 0 ){
            c= c + s;
        }
    }
    return c;
}


//================================================= 判断有特殊字符 start
//处理有无指定字符
bool handleHaveStr(string content, string zd){
    string s=""; int i=0;
    bool handleHaveStr= false;
    for( i= 1 ; i<= len(zd); i++){
        s= mid(zd, i, 1);
        if( inStr(content, s) > 0 ){
            handleHaveStr= true;
            return handleHaveStr;
        }
    }
    return handleHaveStr;
}
//有小写(20151224)
bool haveLowerCase(string content){
    return handleHaveStr(content, "abcdefghijklmnopqrstuvwxyz");
}
//有大写(20151224)
bool haveUpperCase(string content){
    return handleHaveStr(content, "ABCDEFGHIJKLMNOPQRSTUVWXYZ");
}
//有数字(20151224)
bool haveNumber(string content){
    return handleHaveStr(content, "0123456789");
}
//有汉字(20151224)
bool haveChina(string content){
    int i=0; int j=0;
    bool haveChina= false;
    for( i= 1 ; i<= len(content); i++){
        j= asc(mid(content, i, 1));
        //是汉字累加
        if( j < 0 ){
            if((j <= -22033 && j >= -24158)== false ){
                haveChina= true;
                return haveChina;
            }
        }
    }
    return haveChina;
}
//================================================= 判断有特殊字符 end




//*************************************************
//函数名：EncodeJP
//作  用：过滤日本片假名导致Access搜索失效的bug
//*************************************************
string encodeJP(string str){
    if( str== "" ){
        str= replace(str, "ガ", "&#12460;");
        str= replace(str, "ギ", "&#12462;");
        str= replace(str, "グ", "&#12464;");
        str= replace(str, "ア", "&#12450;");
        str= replace(str, "ゲ", "&#12466;");
        str= replace(str, "ゴ", "&#12468;");
        str= replace(str, "ザ", "&#12470;");
        str= replace(str, "ジ", "&#12472;");
        str= replace(str, "ズ", "&#12474;");
        str= replace(str, "ゼ", "&#12476;");
        str= replace(str, "ゾ", "&#12478;");
        str= replace(str, "ダ", "&#12480;");
        str= replace(str, "ヂ", "&#12482;");
        str= replace(str, "ヅ", "&#12485;");
        str= replace(str, "デ", "&#12487;");
        str= replace(str, "ド", "&#12489;");
        str= replace(str, "バ", "&#12496;");
        str= replace(str, "パ", "&#12497;");
        str= replace(str, "ビ", "&#12499;");
        str= replace(str, "ピ", "&#12500;");
        str= replace(str, "ブ", "&#12502;");
        str= replace(str, "ブ", "&#12502;");
        str= replace(str, "プ", "&#12503;");
        str= replace(str, "ベ", "&#12505;");
        str= replace(str, "ペ", "&#12506;");
        str= replace(str, "ボ", "&#12508;");
        str= replace(str, "ポ", "&#12509;");
        str= replace(str, "ヴ", "&#12532;");
    }
    return str;
}
</script>
