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
//Cai采集内容(2014)

//获得采集内容

//处理获得采集内容
string handleGetHttpPage(string httpurl,char char_Set){return "";return "";//留空函数
    return "";
}
//获得请求url状态

//获得请求url的服务器名称


//获得采集内容 (辅助)

//获得采集内容 (辅助)

string bytesToBstr(string body,string cset){return "";//留空函数
    return "";
}
//截取字符串 更新20160114
//c=[A]abbccdd[/A]
//0=abbccdd
//1=[A]abbccdd[/A]
//3=[A]abbccdd
//4=abbccdd[/A]
//截取字符串
string strCut( string content, string startStr, string endStr, int nType){
    int nS1=0; string s1Str=""; int nS2=0; int nS3=0; string c=""; string tempContent=""; string tempStartStr=""; string tempEndStr="" ;string cutType="";
    tempStartStr= startStr;
    tempEndStr= endStr;
    tempContent= content;
    cutType= "|" + nType + "|";
    //不区分大小写
    if( inStr(cutType, "|lu|") > 0 ){
        content= lCase(content);
        startStr= lCase(startStr);
        endStr= lCase(endStr);
    }
    if( inStr(content, startStr)== 0 || inStr(content, endStr)== 0 ){
        c= "";
        return "";
    }
    if( inStr(cutType, "|1|") > 0 ){
        nS1= inStr(content, startStr);
        s1Str= mid(content, nS1 + len(startStr),-1);
        nS2= nS1 + inStr(s1Str, endStr) + len(startStr) + len(endStr) - 1; //为什么要减1
    }else{
        nS1= inStr(content, startStr) + len(startStr);
        s1Str= mid(content, nS1,-1);
        //nS2 = InStr(nS1, content, EndStr)
        nS2= nS1 + inStr(s1Str, endStr) - 1;
    }
    nS3= nS2 - nS1;
    if( nS3 >= 0 ){
        c= mid(tempContent, nS1, nS3);
    }else{
        c= "";
    }
    if( inStr(cutType, "|3|") > 0 ){
        c= tempStartStr + c;
    }
    if( inStr(cutType, "|4|") > 0 ){
        c= c + tempEndStr;
    }
    return c;
}
//获得截取内容,20150305
string getStrCut( string content, string startStr, string endStr, int nType){
    return strCut(content, startStr, endStr, nType);
}
//接取字符 CutStr(Content,22,"null")
string cutStr( string content, int nCutValue, string MoreStr){
    int i=0; int nS=0; int n=0;
    n= 0; //转换成数字类型    追加于20141107
    if( MoreStr== "" ){ MoreStr= "..." ;}
    if( lCase(MoreStr)== "none" || lCase(MoreStr)== "null" ){ MoreStr= "" ;}
    string cutStr= content;
    for( i= 1 ; i<= len(content); i++){
        nS= asc(mid(cStr(content), i, 1));
        if( nS < 0 ){ nS= nS + 65536 ;}
        if( nS < 255 ){ n= n + 1 ;}
        if( nS > 255 ){ n= n + 2 ;}
        if( n >= nCutValue ){ return left(content, i) + MoreStr ; }
    }
    return cutStr;
}
//截取内容，不区分大小写 20150327  C=CutStrNOLU(c,"<heAd",">")
string cutStrNOLU(string content, string startStr, string endStr){
    string s=""; string LCaseContent=""; int nStartLen=0; int nEndLen=0; string sNewStrStart="";
    startStr= lCase(startStr);
    endStr= lCase(endStr);
    LCaseContent= lCase(content);
    string cutStrNOLU="";
    if( inStr(LCaseContent, startStr) > 0 ){
        nStartLen= inStr(LCaseContent, startStr);
        s= mid(content, nStartLen,-1);
        LCaseContent= mid(s, len(startStr) + 1,-1);
        sNewStrStart= mid(s, 1, len(startStr) + 1); //获得开始字符

        LCaseContent= replace(LCaseContent, "<", "&lt;");
        //Call eerr("111",LCaseContent)

        nEndLen= inStr(LCaseContent, endStr);
        echo("nEndLen", nEndLen);

        s= mid(content, nStartLen, nEndLen + len(startStr));
        //Call Echo(nStartLen,nEndLen)
        //Call Echo("S",S)
        cutStrNOLU= s;
    }
    return cutStrNOLU;
}

//接取TD字符
string setCutTDStr( string content, string TDWidth, string moreColor){
    int i=0; string s=""; string c=""; int n=0; int nEnd=0; bool isMore;
    content= cStr(content + "");
    if( content== "" ){ return content ; }
    if( TDWidth== "" ){ return content ; }//TDWidth为空，则为自动
    n= 0 ; isMore= false;
    nEnd= cInt(cInt(TDWidth) / 6.3);
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( n >= nEnd ){
            isMore= true;
            break;
        }else{
            c= c + s;
        }
        if( asc(s) < 0 ){
            n= n + 2;
        }else{
            n= n + 1;
        }
    }
    if( isMore== true ){
        //需要处理Title标题的HTML
        c= "<span Title=\"" + displayHtml(content) + "\" style=\"background-color:" + moreColor + ";\">" + c + "</span>";
    }
    return c;
}
//接取TD字符 (辅助)
string cutTDStr(string content, string TDWidth){
    return setCutTDStr(content, TDWidth, "#FBE3EF");
}
//分割字符




//创建于20160801
string getArrayNew( string content, string startStr, string endStr, bool isStart, bool isEnd){
    int i=0;string s="";string listStr="";int nStartLen=0;int nEndLen=0;
    for( i= 0 ; i<= 999 				; i++){//30为截取条件
        //echo(content . "=" & instr(content, startStr))
        //call echo(instr(content, startStr) , instr(content, endStr))
        nStartLen=inStr(content, startStr);
        nEndLen=inStr(content, endStr);
        if( nStartLen > 0 && nEndLen > 0 ){

            s=mid(content,1,nEndLen-1);
            content=mid(content,nEndLen+len(endStr),-1);
            s=mid(s,inStr(s,startStr)+len(startStr),-1);

            if( listStr!="" ){
                listStr=listStr + "$Array$";
            }
            if( isStart==true ){
                s=startStr + s;
            }
            if( isEnd==true ){
                s=s + endStr;
            }
            listStr=listStr + s;


        }else{
            break;
        }
    }
    return listStr;
}

//分割字符 不处理字符 (辅助)
string getArray(string content, string startStr, string endStr, bool isStart, bool isEnd){
    return getArrayList(content, startStr, endStr, isStart, isEnd, "");
}
//分割字符 去掉字符 (辅助)
string getArray1(string content, string startStr, string endStr, bool isStart, bool isEnd){
    return getArrayList(content, startStr, endStr, isStart, isEnd, "去掉字符");
}
//截取指定分割值
string getSplit( string content, string sSplit, int n){
    string[] splxx;
    splxx= aspSplit(content, sSplit);
    return splxx[n];
}
//获得分数总数
int getSplitCount( string content, string sSplit){
    string[] splxx;
    splxx= aspSplit(content, sSplit);
    int getSplitCount=0;
    getSplitCount= uBound(splxx);
    if( getSplitCount > 0 ){ getSplitCount= getSplitCount + 1 ;}//不为空加一

    return getSplitCount;
}

//代理 因为它不能与VB软件共存
string agent(string httpurl){//留空函数
    return "";
}

</script>
