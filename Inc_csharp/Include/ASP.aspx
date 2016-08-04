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
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.IO" %>
<%@ import namespace="System.Diagnostics" %>
<%@ import namespace="System.Configuration" %>
<%@ import namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %>

<%@ import namespace="System.Linq" %>
<%@ Import Namespace="Microsoft.Win32" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="System.Net" %>  

<%@ Import Namespace="System.Threading" %>
<%@ Import Namespace="System.Net.Sockets" %>
<%@ Import Namespace="System.Diagnostics" %>
<%@ Import Namespace="System.Text.RegularExpressions" %> 
<%@ Import Namespace="System.Collections.Generic" %> 

<%@ Import Namespace="System.Web" %> 
<%@ Import Namespace="System.Web.Security" %> 
<%@ Import Namespace="System.Web.UI" %> 
<%@ Import Namespace="System.Web.UI.HtmlControls" %> 
<%@ Import Namespace="System.Web.UI.WebControls" %> 
<%@ Import Namespace="System.Web.UI.WebControls.WebParts" %> 
<%@ Import Namespace="System.Xml.Linq" %> 
<%@ Import Namespace="System.Text"		 %> 
<%@ Import Namespace="System.Data.SqlClient" %> 
<%@ Import Namespace="System.Configuration" %> 
<%@ Import Namespace="System.Collections.Generic" %> 
<%@ Import Namespace="System.Security.Cryptography" %> 
<%@ Import Namespace="System.Net" %> 
<%@ Import Namespace="System.Net.Sockets" %>   
<%@ Import Namespace="ADOX" %>    
<%--@ Import Namespace="1111111" --%>  

<script runat="server" language="c#"> 
/*
此文件的目的就是让与ASP程序里保持一致

string path = @"D:\aaa.txt";  这种用法待学习

*/
string gbl_ASP="asp.aspx";




//换行
string vbCrlf(){
	return "\n";	
}
//退格  不过它用不上
string vbTab(){
	return "\t";	
}
//从左向右搜索字符
int inStr(string content,string searchStr){
	return content.IndexOf(searchStr)+1;
}
int inStr(int n,string searchStr){
	string s=n.ToString();
	return inStr(s,searchStr);
}
//从右向左搜索字符
int inStrRev(string content,string searchStr){
	return content.LastIndexOf(searchStr)+1;
}
//字符长度
int len(string content){
	return content.Length;
}
int len(int n){
	string s=n.ToString();
	return len(s);
}
//截取字符
string mid(string content,int nStart,int nLength=-1){
	nStart--;
	if(content==""){
		return "";
	}else if(nStart>content.Length){
		return "";
	}
	if(nLength+nStart>content.Length){
		nLength=-1;
	}
	if(nLength==-1){
		return content.Substring(nStart);
	} 
	try{
		return content.Substring(nStart,nLength);
	}catch(Exception e){  
		
		eerr(content, content.Length + ">" + nStart+"+"+nLength);
		return "";
	}
	
}
//替换
string replace(string content,string findStr,string replaceStr){
	if(content==""){
		return "";
	}else if(findStr==""){
		return content;
	}else{
		return content.Replace(findStr,replaceStr);
	}
	//return content.Split(findStr).Join(replaceStr);			//js能用
}
string replace(string content,char findStr,string replaceStr){
	return content.Replace(findStr.ToString(),replaceStr);
	//return content.Split(findStr).Join(replaceStr);			//js能用
}
string replace(string content,string findStr,char replaceStr){
	return content.Replace(findStr,replaceStr.ToString());
	//return content.Split(findStr).Join(replaceStr);			//js能用
}
string replace(string content,string findStr,int replaceStr){
	return content.Replace(findStr,replaceStr.ToString());
	//return content.Split(findStr).Join(replaceStr);			//js能用
}
//随机数

double rnd(){
	Random rndObj = new Random(); 
	return  Convert.ToDouble("0."+rndObj.Next(111111,999999)); 
}
/*
int rnd(){
	Random rndObj = new Random(); 
	return  Convert.ToInt32("0."+rndObj.Next(111111,999999));  
}
*/
//转小写
string lCase(string content){
	return content.ToLower();
}
string lCase(int content){
	return content.ToString().ToLower();
}
string lCase(double content){
	return content.ToString().ToLower();
}
//转大写
string uCase(string content){
	return content.ToUpper();
}
string uCase(int content){
	return content.ToString().ToUpper();
}
string uCase(double content){
	return content.ToString().ToUpper();
}
//字符转ASC 
int asc(string s) { 
	char charS=Convert.ToChar(mid(s,1,1));	
	return (int)charS;
}
int ascW(string s) { 
	return asc(s);
}

double abs(double n){
	return System.Math.Abs(n);
}

//数字转CHR
char chr(int n){
	return (char)n;
} 
char chr(double n){
	return chr((int)n);
} 
char chr(string s){
	return chr(Convert.ToInt32(s));
} 
//left从左边开始截取
string left(string param, int nLength){
	if(param=="" || param==null){
		return "";
	}else{
		if(nLength>param.Length){
			nLength=param.Length;
		}
		return param.Substring(0, nLength);
	}
}
//right从右边开始截取
string right(string param, int nLength){
	if(param=="" || param==null){
		return "";
	}else{
		int n=param.Length - nLength;
		if(n<0){
			n=0;
		}
		if(nLength>param.Length){
			nLength=param.Length;
		}
		return param.Substring(n, nLength);
	}
}

//asp版分割字符 
string[] aspSplit(string str,string sSplit){
	//必需起用  System.Text.RegularExpressions  包
	//一个只符让它用这种方法，要不然经不法处理\\ 这种的，晕
	if(sSplit.Length==1){
		string[] sArray=str.Split(Convert.ToChar(sSplit));  
		return sArray; 
	}else{	
		 string pattern = Regex.Escape(sSplit);
		string[] sArray=Regex.Split(str,pattern,RegexOptions.IgnoreCase); 
		return sArray; 
	}
} 
string[] aspSplit(object o,string sSplit){
	string s=o.ToString();
	return aspSplit(s,sSplit);
}
//定义数组  与asp版一致
string[] aspArray(int n){
	return new string[n+1];
} 
int[] aspNArray(int n){
	return new int[n+1];
} 
//获得数组长度
int uBound(string[] splstr){
	return splstr.Length-1;
}
int uBound(int[] splstr){
	return splstr.Length-1;
}
//获得数组开始长度
int lBound(string[] splstr){
	return 0;
} 
int lBound(int[] splstr){
	return 0;
} 
//返回指定数值的整数部分
/*
int fix(double n){
	string str = n.ToString();
	if (str.IndexOf(".")!=-1)     {          
		return Convert.ToInt32( str.Substring(0,str.IndexOf(".")) );
	}
	return Convert.ToInt32(str);
}
*/ 
int fix(double nD) {   
	if (nD >= 0) {   
		return (int)Math.Floor(nD);   
	}   
	return (int)Math.Ceiling(nD);   
}
int hex(int n){
	return n;
}

//返回指定变量子类型的信息
//string typeName(str){  做不了，暂时留着

//.net里没有 int函数
//将表达式转化为Integer数值子类型
int cInt(string str){ 
	return Convert.ToInt32(str);
}
int cInt(double n){ 
	return Convert.ToInt32(n);
}
int cInt(object n){ 
	return Convert.ToInt32(n);
}
int cInt(int n){ 
	return n;
}

//将表达式转化为Integer数值子类型
int cLng(string str){ 
	return Convert.ToInt32(str);
}
//转字符
string cStr(int n){ 
	return n+"";
}
//转字符
string cStr(double d){
	//echo("","d");
	return d+"";
}
//转字符
string cStr(string s){
	//echo("","s");
	if(s==null){
		s="";
	}
	return s;
}
//转字符
string cStr(object o){
	//echo("","o");
	if(o==null){
		return "";
	}else{
		return o.ToString();	
	}
}
string toString(string s){
	return cStr(s);
}
//转bool
bool cBool(object o){
	string s=lCase(o.ToString());
	if(s=="false" || s=="0"){
		return false;
	}else{
		return true;
	}
}
//转时间
System.DateTime cTime(string sTime){
	return Convert.ToDateTime(sTime);
}

//清除左边空格 ASP版
string aspLTrim(string content){ 
	for(int i=1 ; i<= len(content); i++){
		string s=mid(content,i,1);
		if(s!=" "){
			content=mid(content,i,-1);
			 break;
		 }
	}
	return content;
	//return content.TrimStart(); 
}
//清除右边空格 ASP版
string aspRTrim(string content){
	for(int i=len(content); i>=1; i--){
		string s=mid(content,i,1);
		if(s!=" "){
			content=mid(content,1,i);
			 break;
		 }
	}
	return content;
	
	//return content.TrimEnd(); 
}
//两除两边空格 ASP版
string aspTrim(string content){ 
	return aspRTrim(aspLTrim(content)); 
}
string aspTrim(int n){ 
	string s=n.ToString();
	return aspTrim(s); 
}
//清除左边空格+换行 PHP版
string phpLTrim(string str){ 
	for(int i=1 ; i<= len(str); i++){
		string s=mid(str,i,1);
		if(s!=" " && s!="\t" && s!="\n"){
			str=mid(str,i,-1);
			 break;
		 }
	}
	return str;
}
//清除右边空格+换行 PHP版
string phpRTrim(string str){ 
	for(int i=len(str) ; i>=1; i--){
		string s=mid(str,i,1);
		if(s!=" " && s!="\t" && s!="\n"){
			str=mid(str,1,i);
			 break;
		 }
	}
	return str;
}
//清除两这空格+换行 PHP版
string phpTrim(string str){ 
	return phpRTrim(phpLTrim(str)); 
}
string trimVbCrlf(string str){
	return phpTrim(str);
}
string trimVbTab(string str){
	return  phpRTrim(phpLTrim(str));
}



//获得日期
string aspDate() {
	return System.DateTime.Now.ToString("d");
}
//获得时间
string aspTime() {
	return System.DateTime.Now.ToString("T");
}
//判断时间
bool isDate(string strDate){ 
	DateTime dtDate;
	if (DateTime.TryParse(strDate, out dtDate)){
		return true;
	}else{
		return false;
	}
}
//获得当前时间
System.DateTime now() {
	return System.DateTime.Now;
	//return System.DateTime.Now.ToString("G");
}
//获得当前时间
System.DateTime timer() {
	return System.DateTime.Now;
	//return System.DateTime.Now.ToString("G");
}

//获得年
int year(System.DateTime  timeStr){ 
	return timeStr.Year; 
}
//获得年
int year(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return year(data1);
}
//获得月
int month(System.DateTime  timeStr){
	return timeStr.Month; 
}
//获得月
int month(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return month(data1);
}
//获得日
int day(System.DateTime  timeStr){
	return timeStr.Day; 
}
//获得日 
int day(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return day(data1);
}
//获得小时
int hour(System.DateTime  timeStr){
	return timeStr.Hour; 
}
//获得小时
int hour(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return hour(data1);
}
//获得分种
int minute(System.DateTime  timeStr){
	return timeStr.Minute; 
}
//获得分种
int minute(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return minute(data1);
}
//获得秒
int second(System.DateTime  timeStr){
	return timeStr.Second; 
}
//获得秒
int second(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return second(data1);
} 
//添加时间
string sAddTime(System.DateTime timeObj, string sType, int nValue){
	return "";
}

//订义结构体
enum DateInterval{
    Second, Minute, Hour, Day, Week, Month, Quarter, Year
}
//计算两个日期之间的时间间隔
int dateDiff(string Interval, DateTime arg_StartDate, DateTime arg_EndDate){
    DateInterval objDateInterval;
    switch (Interval){	
		//秒
        case "s":
            objDateInterval = DateInterval.Second;
            break;
		//分钟
        case "n":
            objDateInterval = DateInterval.Minute;
            break;
		//小时
        case "h":
            objDateInterval = DateInterval.Hour;
            break;
		//日
        case "d":
            objDateInterval = DateInterval.Day;
            break;
		//周
        case "w":
            objDateInterval = DateInterval.Week;
            break;
		//月
        case "m":
            objDateInterval = DateInterval.Month;
            break;
		//季节
        case "q":
           objDateInterval = DateInterval.Quarter;
            break;
		//年
        case "y":
           objDateInterval = DateInterval.Year;
            break;
        default :
            objDateInterval = DateInterval.Second;
            break;
    }
    return this.dateDiff(objDateInterval,arg_StartDate,arg_EndDate);
}
//多型2，传入结构体
int dateDiff(DateInterval arg_Interval, DateTime arg_StartDate, DateTime arg_EndDate){
    int lngDateDiffValue = 0;
    System.TimeSpan objTimeSpan = new System.TimeSpan(arg_EndDate.Ticks - arg_StartDate.Ticks);
    switch (arg_Interval){
        case DateInterval.Second:
            lngDateDiffValue = (int)objTimeSpan.TotalSeconds;
            break;
        case DateInterval.Minute:
            lngDateDiffValue = (int)objTimeSpan.TotalMinutes;
            break;
        case DateInterval.Hour:
            lngDateDiffValue = (int)objTimeSpan.TotalHours;
            break;
        case DateInterval.Day:
            lngDateDiffValue = (int)objTimeSpan.TotalDays;
            break;
        case DateInterval.Week:
            lngDateDiffValue = (int)(objTimeSpan.TotalSeconds / ( 7 * 24 * 60 * 60 )); //一个星期7天
            break;
        case DateInterval.Month:
            lngDateDiffValue = (int)(objTimeSpan.TotalSeconds / (30 * 24 * 60 * 60)); //一个月计30天
            break;
        case DateInterval.Quarter:
            lngDateDiffValue = (int)(objTimeSpan.TotalSeconds / (90 * 24 * 60 * 60)); //一季计3月
            break;
        case DateInterval.Year:
            lngDateDiffValue = (int)(objTimeSpan.TotalSeconds / (365 * 24 * 60 * 60));   //一年计365天
            break;
    }
    return (lngDateDiffValue);
}
/// 返回已添加指定时间间隔的日期 
DateTime dateAdd(String interval, Double number, DateTime date){
	string tempInterval = interval.ToLower();
	switch (tempInterval){
	//年 
	case "yyyy": 
	case "y":
		return date.AddYears(Convert.ToInt32(number));
		break;
	//月 
	case "m":
		return date.AddMonths(Convert.ToInt32(number));
		break;
	//日 
	case "d":
		return date.AddDays(number);
		break;
	//小时 
	case "h":
		return date.AddHours(number);
		break;
	//分钟 
	case "n":
		return date.AddMinutes(number);
		break;
	//秒 
	case "s":
		return date.AddSeconds(number);
		break; 
	default:
		return date;
		break;
	}
}

/// 返回给定日期的指定部分
double datePart(String interval, DateTime date){
	string tempInterval = interval.ToLower();
	switch (tempInterval)	{
	//年 
	case "yyyy": 
	case "y":
		return Convert.ToDouble(date.Year);
		break;
	//月 
	case "m":
		return Convert.ToDouble(date.Month);
		break;
	//日 
	case "d":
		return Convert.ToDouble(date.Day);
		break;
	//小时 
	case "h":
		return Convert.ToDouble(date.Hour);
		break;
	//分钟 
	case "n":
		return Convert.ToDouble(date.Minute);
		break;
	//秒 
	case "s":
		return Convert.ToDouble(date.Second);
		break;  
	default:
		return 0.0000;
		break;
	}
} 

//将各个分量，加上指定分隔符，合成一个字符串。
string join(string[]parts, string s){
	return string.Join(s, parts); 
}



//功能:ASP里的IIF 如：IIf(1 = 2, "a", "b") 
string IIF(bool bExp, string sVal1, string sVal2){
	//if(bExp==true){return sVal1;}else{return sVal2;}
	return bExp?sVal1:sVal2;		//二元运算符
}
int IIF(bool bExp, int sVal1, int sVal2){ 
	return bExp?sVal1:sVal2;		//二元运算符
}
bool IIF(bool bExp, bool sVal1, bool sVal2){ 
	return bExp?sVal1:sVal2;		//二元运算符
}
//位置 与ASP里的统一
string mapPath(string path){
	path=path+"/";
	string ffPath=Server.MapPath(path  );
	ffPath=mid(ffPath,1,len(ffPath)-1);
	return ffPath;
}
void die(){
	Response.End();
}

//随机数
int phpRand(int nMinimum, int nMaximum){
	Random rndObj = new Random(); 
	return  rndObj.Next(nMinimum,nMaximum); 
}


//格式化成价格如 108.00 (20150806使用ASP转PHP制作)
int formatNumber(double nNumb,int n,int p=1){
	string dianLeft=""; string dianRight=""; int i; string c=""; string s;
	string content=nNumb.ToString();
	if( inStr(content,".")> 0 ){
		dianLeft = mid(content,1,inStr(content,".")-1);
		dianRight = mid(content,inStr(content,".")+1,-1);
	}else{
		dianLeft=content;
	 }
	dianRight=dianRight + "0000000000";
	for( i=1 ; i<= n; i++){
		s=mid(dianRight,i,1);
		c=c + s;  
	}  
	if( n>0 ){ 
		dianLeft = dianLeft + ".";
	 } 
	c = dianLeft +c;  
	return Convert.ToInt32(c);
} 
int formatNumber(int nNumb,int n){
	return formatNumber((double)nNumb,n);
}

bool isEmpty(string s){
	if(s==""){
		return true;
	}else{
		return false;
	}
}
 
//md5加密
string aspMD5(string pwd, int n){
	MD5 md5 = new MD5CryptoServiceProvider();
	byte[] data = System.Text.Encoding.Default.GetBytes(pwd);//将字符编码为一个字节序列 
	byte[] md5data = md5.ComputeHash(data);//计算data字节数组的哈希值 
	md5.Clear();
	string str = "";
	for (int i = 0; i < md5data.Length - 1; i++){
		str += md5data[i].ToString("x2").PadLeft(2, '0');
	}  
	return str;
}
//我自己的md5
string myMD5(string s){
	return aspMD5(s,32);
	//s = System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(s, "md5").ToString();
    //return s.ToLower();
}
 

bool isNull(string s){
	if(s==""){
		return true;
	}else{
		return false;
	}
}


/*一个测试类   
调用方法  Response.Write(gbl.myTest());
gbl cOBJ=new gbl();
Response.Write(cOBJ.myFun());
*/
public class gbl{   
	public static   int   intXXX=1;   
	public static string myTest(){
		return "myTest";
	}   
	public string myFun(){
		return "myFun";
	}
}    

//************************************************* 等完善
 

 
</script>
