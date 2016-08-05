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
//Time 时间操作类 (2013,9,27)

//时间处理
string format_Time(string timeStr, int nType){
    string y=""; string m=""; string d=""; string h=""; string mi=""; string s="" ;string c="";int nWeek=0;
    if( isDate(timeStr)== false ){ return ""; }
    y= cStr(year(timeStr));
    m= cStr(month(timeStr));
    if( len(m)== 1 ){ m= "0" + m ;}
    d= cStr(day(timeStr)); //在vb.net里要这样用  D = CStr(CDate(timeStr).Day)
    //周
    nWeek=fix(day(timeStr)/7);
    if( day(timeStr) % 7>0 ){
        nWeek=nWeek+1;
    }

    if( len(d)== 1 ){ d= "0" + d ;}
    h= cStr(hour(timeStr));
    if( len(h)== 1 ){ h= "0" + h ;}
    mi= cStr(minute(timeStr));
    if( len(mi)== 1 ){ mi= "0" + mi ;}
    s= cStr(second(timeStr));
    if( len(s)== 1 ){ s= "0" + s;}
    switch ( nType ){
        case 1:
        //yyyy-mm-dd hh:mm:ss
        c= y + "-" + m + "-" + d + " " + h + ":" + mi + ":" + s ;break;
        case 2:
        //yyyy-mm-dd
        c= y + "-" + m + "-" + d ;break;
        case 3:
        //hh:mm:ss
        c= h + ":" + mi + ":" + s ;break;
        case 4:
        //yyyy年mm月dd日
        c= y + "年" + m + "月" + d + "日" ;break;
        case 5:
        //yyyymmdd
        c= y + m + d ;break;
        case 6:
        //yyyymmddhhmmss
        c= y + m + d + h + mi + s ;break;
        case 7:
        //mm-dd
        c= m + "-" + d ;break;
        case 8:
        //yyyy年mm月dd日
        c= y + "年" + m + "月" + d + "日" + " " + h + ":" + mi + ":" + s ;break;
        case 9:
        //yyyy年mm月dd日H时mi分S秒 早上
        c= y + "年" + m + "月" + d + "日" + " " + h + "时" + mi + "分" + s + "秒，" + getDayStatus(h, 1) ;break;
        case 10:
        //yyyy年mm月dd日H时
        c= y + "年" + m + "月" + d + "日" + h + "时" ;break;
        case 11:
        //yyyy年mm月dd日H时mi分S秒
        c= y + "年" + m + "月" + d + "日" + " " + h + "时" + mi + "分" + s + "秒" ;break;
        case 12:
        //yyyy年mm月dd日H时mi分
        c= y + "年" + m + "月" + d + "日" + " " + h + "时" + mi + "分" ;break;
        case 13:
        //yyyy年mm月dd日H时mi分 早上
        c= m + "月" + d + "日" + " " + h + ":" + mi + " " + getDayStatus(h, 0) ;break;
        case 14:
        //yyyy年mm月dd日
        c= y + "/" + m + "/" + d ;break;
        case 15:
        //yyyy年mm月 第1周
        c= y + "年" + m + "月 第" + nWeek + "周";
        break;
    }
    return c;
}
//获得当前时期还可以计算

//获得当前天状态
string getDayStatus(string h, int nType){
    string c="";int nH=0;
    if( left(h, 1)== "0" ){
        h= right(h, 1);
    }
    nH= cInt(h);
    if( nH >= 0 && nH <= 5 ){
        c= "凌晨";
    }else if( nH >= 6 && nH <= 8 ){
        c= "早上";
    }else if( nH >= 9 && nH <= 12 ){
        c= "上午";
    }else if( nH >= 13 && nH <= 18 ){
        c= "下午";
    }else if( nH >= 19 && nH <= 24 ){
        c= "晚上";
    }else{
        c= "深夜";
    }
    if( nType== 1 ){ c= "<b>" + c + "</b>" ;}
    return c;
}
//时间计算
string printTimeValue( int nV){

    int nVTemp=0; string c=""; int n=0;
    nVTemp=nV;
    if( nV >= 86400 ){
        n= nV % 86400;
        nV= fix(nV / 60 / 60 / 24); //这里面是除24，不是60，因为一天有24个小时，小云你这个傻X  20150119
        c= c + nV + "天";
        nV= n;
    }
    if( nV >= 3600 ){
        n= nV % 3600;
        nV= fix(nV / 60 / 60);
        c= c + nV + "小时";
        nV= n;
    }

    if( nV >= 60 ){
        n= nV % 60;
        nV= fix(nV / 60);
        c= c + nV + "分";
        nV= n;
    }
    if( nV > 0 ){
        c= c + nV + "秒";
    }

    return c;
}
//计算整时  (问答以小时或分钟显示)
string printAskTime(int nV){

    string c=""; int n=0;
    if( nV >= 3600 ){
        n= nV % 3600;
        nV= fix(nV / 60 / 60);
        c= c + nV + "小时";
        nV= n;
        return c ;
    }
    if( nV >= 60 ){
        n= nV % 60;
        nV= fix(nV / 60);
        c= c + nV + "分钟";
        nV= n;
        return c ;
    }
    if( nV >= 0 ){
        c= c + nV + "秒钟";
        return c ;
    }
    return "";
}
//获得时间
int getTimerSet(){
    return calculationTimer(pubTimer,now());
}
//计算时间
int calculationTimer(System.DateTime startTime,System.DateTime endTime){
    int n=0;
    //n = formatNumber((timer() - pubTimer) * 1000, 2, - 1) / 1000
    //calculationTimer = toNumber(n, 3)
    return dateDiff("s", startTime,endTime);
}

//获得时间
string getTimer(){
    return "<br>------------------------------------<br>页面执行时间：" + getTimerSet() + " 秒";
}
//获得时间
string vbGetTimer(){
    string vbGetTimer= vbCrlf() + "------------------------------------" + vbCrlf() + "运行时间：" + calculationTimer(pubTimer,now());
    return "";
}
//获得第三种
string vbEchoTimer(){
    string vbEchoTimer= "运行时间：" + calculationTimer(pubTimer,now()) + " 秒";
    return "";
}
//获得时间第四种
string vbRunTimer(System.DateTime startTime){
    string vbRunTimer= "运行时间：" + calculationTimer(pubTimer,now()) + " 秒";
    return "";
}



//添加时间
string sAddTime(string timeObj,string sType,int nValue){
    string s="";
    //秒
    if( sType=="s" ){
        s=timeObj+0.00001*nValue;
        //分钟
    }else if( sType=="n" ){
        s=timeObj+0.00060*nValue;
        //小时
    }else if( sType=="h" ){
        s=timeObj+0.00001*nValue*60*60;
        //日
    }else if( sType=="d" ){
        s=timeObj+0.00001*nValue*60*60*24;
        //周
    }else if( sType=="w" ){
        s=timeObj+0.00001*nValue*60*60*24*7;
        //月
    }else if( sType=="m" ){
        s=timeObj+0.00001*nValue*60*60*24*30;
        //季节
    }else if( sType=="q" ){
        s=timeObj+0.00001*nValue*60*60*24*120;
        //年
    }else if( sType=="y" ){
        s=timeObj+0.00001*nValue*60*60*24*366;
    }
    return s;
}
</script>
