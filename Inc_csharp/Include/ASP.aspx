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
���ļ���Ŀ�ľ�������ASP�����ﱣ��һ��

string path = @"D:\aaa.txt";  �����÷���ѧϰ

*/
string gbl_ASP="asp.aspx";




//����
string vbCrlf(){
	return "\n";	
}
//�˸�  �������ò���
string vbTab(){
	return "\t";	
}
//�������������ַ�
int inStr(string content,string searchStr){
	return content.IndexOf(searchStr)+1;
}
int inStr(int n,string searchStr){
	string s=n.ToString();
	return inStr(s,searchStr);
}
//�������������ַ�
int inStrRev(string content,string searchStr){
	return content.LastIndexOf(searchStr)+1;
}
//�ַ�����
int len(string content){
	return content.Length;
}
int len(int n){
	string s=n.ToString();
	return len(s);
}
//��ȡ�ַ�
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
//�滻
string replace(string content,string findStr,string replaceStr){
	if(content==""){
		return "";
	}else if(findStr==""){
		return content;
	}else{
		return content.Replace(findStr,replaceStr);
	}
	//return content.Split(findStr).Join(replaceStr);			//js����
}
string replace(string content,char findStr,string replaceStr){
	return content.Replace(findStr.ToString(),replaceStr);
	//return content.Split(findStr).Join(replaceStr);			//js����
}
string replace(string content,string findStr,char replaceStr){
	return content.Replace(findStr,replaceStr.ToString());
	//return content.Split(findStr).Join(replaceStr);			//js����
}
string replace(string content,string findStr,int replaceStr){
	return content.Replace(findStr,replaceStr.ToString());
	//return content.Split(findStr).Join(replaceStr);			//js����
}
//�����

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
//תСд
string lCase(string content){
	return content.ToLower();
}
string lCase(int content){
	return content.ToString().ToLower();
}
string lCase(double content){
	return content.ToString().ToLower();
}
//ת��д
string uCase(string content){
	return content.ToUpper();
}
string uCase(int content){
	return content.ToString().ToUpper();
}
string uCase(double content){
	return content.ToString().ToUpper();
}
//�ַ�תASC 
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

//����תCHR
char chr(int n){
	return (char)n;
} 
char chr(double n){
	return chr((int)n);
} 
char chr(string s){
	return chr(Convert.ToInt32(s));
} 
//left����߿�ʼ��ȡ
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
//right���ұ߿�ʼ��ȡ
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

//asp��ָ��ַ� 
string[] aspSplit(string str,string sSplit){
	//��������  System.Text.RegularExpressions  ��
	//һ��ֻ�����������ַ�����Ҫ��Ȼ����������\\ ���ֵģ���
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
//��������  ��asp��һ��
string[] aspArray(int n){
	return new string[n+1];
} 
int[] aspNArray(int n){
	return new int[n+1];
} 
//������鳤��
int uBound(string[] splstr){
	return splstr.Length-1;
}
int uBound(int[] splstr){
	return splstr.Length-1;
}
//������鿪ʼ����
int lBound(string[] splstr){
	return 0;
} 
int lBound(int[] splstr){
	return 0;
} 
//����ָ����ֵ����������
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

//����ָ�����������͵���Ϣ
//string typeName(str){  �����ˣ���ʱ����

//.net��û�� int����
//�����ʽת��ΪInteger��ֵ������
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

//�����ʽת��ΪInteger��ֵ������
int cLng(string str){ 
	return Convert.ToInt32(str);
}
//ת�ַ�
string cStr(int n){ 
	return n+"";
}
//ת�ַ�
string cStr(double d){
	//echo("","d");
	return d+"";
}
//ת�ַ�
string cStr(string s){
	//echo("","s");
	if(s==null){
		s="";
	}
	return s;
}
//ת�ַ�
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
//תbool
bool cBool(object o){
	string s=lCase(o.ToString());
	if(s=="false" || s=="0"){
		return false;
	}else{
		return true;
	}
}
//תʱ��
System.DateTime cTime(string sTime){
	return Convert.ToDateTime(sTime);
}

//�����߿ո� ASP��
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
//����ұ߿ո� ASP��
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
//�������߿ո� ASP��
string aspTrim(string content){ 
	return aspRTrim(aspLTrim(content)); 
}
string aspTrim(int n){ 
	string s=n.ToString();
	return aspTrim(s); 
}
//�����߿ո�+���� PHP��
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
//����ұ߿ո�+���� PHP��
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
//�������ո�+���� PHP��
string phpTrim(string str){ 
	return phpRTrim(phpLTrim(str)); 
}
string trimVbCrlf(string str){
	return phpTrim(str);
}
string trimVbTab(string str){
	return  phpRTrim(phpLTrim(str));
}



//�������
string aspDate() {
	return System.DateTime.Now.ToString("d");
}
//���ʱ��
string aspTime() {
	return System.DateTime.Now.ToString("T");
}
//�ж�ʱ��
bool isDate(string strDate){ 
	DateTime dtDate;
	if (DateTime.TryParse(strDate, out dtDate)){
		return true;
	}else{
		return false;
	}
}
//��õ�ǰʱ��
System.DateTime now() {
	return System.DateTime.Now;
	//return System.DateTime.Now.ToString("G");
}
//��õ�ǰʱ��
System.DateTime timer() {
	return System.DateTime.Now;
	//return System.DateTime.Now.ToString("G");
}

//�����
int year(System.DateTime  timeStr){ 
	return timeStr.Year; 
}
//�����
int year(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return year(data1);
}
//�����
int month(System.DateTime  timeStr){
	return timeStr.Month; 
}
//�����
int month(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return month(data1);
}
//�����
int day(System.DateTime  timeStr){
	return timeStr.Day; 
}
//����� 
int day(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return day(data1);
}
//���Сʱ
int hour(System.DateTime  timeStr){
	return timeStr.Hour; 
}
//���Сʱ
int hour(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return hour(data1);
}
//��÷���
int minute(System.DateTime  timeStr){
	return timeStr.Minute; 
}
//��÷���
int minute(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return minute(data1);
}
//�����
int second(System.DateTime  timeStr){
	return timeStr.Second; 
}
//�����
int second(string timeStr){
	System.DateTime data1=Convert.ToDateTime(timeStr);
	return second(data1);
} 
//���ʱ��
string sAddTime(System.DateTime timeObj, string sType, int nValue){
	return "";
}

//����ṹ��
enum DateInterval{
    Second, Minute, Hour, Day, Week, Month, Quarter, Year
}
//������������֮���ʱ����
int dateDiff(string Interval, DateTime arg_StartDate, DateTime arg_EndDate){
    DateInterval objDateInterval;
    switch (Interval){	
		//��
        case "s":
            objDateInterval = DateInterval.Second;
            break;
		//����
        case "n":
            objDateInterval = DateInterval.Minute;
            break;
		//Сʱ
        case "h":
            objDateInterval = DateInterval.Hour;
            break;
		//��
        case "d":
            objDateInterval = DateInterval.Day;
            break;
		//��
        case "w":
            objDateInterval = DateInterval.Week;
            break;
		//��
        case "m":
            objDateInterval = DateInterval.Month;
            break;
		//����
        case "q":
           objDateInterval = DateInterval.Quarter;
            break;
		//��
        case "y":
           objDateInterval = DateInterval.Year;
            break;
        default :
            objDateInterval = DateInterval.Second;
            break;
    }
    return this.dateDiff(objDateInterval,arg_StartDate,arg_EndDate);
}
//����2������ṹ��
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
            lngDateDiffValue = (int)(objTimeSpan.TotalSeconds / ( 7 * 24 * 60 * 60 )); //һ������7��
            break;
        case DateInterval.Month:
            lngDateDiffValue = (int)(objTimeSpan.TotalSeconds / (30 * 24 * 60 * 60)); //һ���¼�30��
            break;
        case DateInterval.Quarter:
            lngDateDiffValue = (int)(objTimeSpan.TotalSeconds / (90 * 24 * 60 * 60)); //һ����3��
            break;
        case DateInterval.Year:
            lngDateDiffValue = (int)(objTimeSpan.TotalSeconds / (365 * 24 * 60 * 60));   //һ���365��
            break;
    }
    return (lngDateDiffValue);
}
/// ���������ָ��ʱ���������� 
DateTime dateAdd(String interval, Double number, DateTime date){
	string tempInterval = interval.ToLower();
	switch (tempInterval){
	//�� 
	case "yyyy": 
	case "y":
		return date.AddYears(Convert.ToInt32(number));
		break;
	//�� 
	case "m":
		return date.AddMonths(Convert.ToInt32(number));
		break;
	//�� 
	case "d":
		return date.AddDays(number);
		break;
	//Сʱ 
	case "h":
		return date.AddHours(number);
		break;
	//���� 
	case "n":
		return date.AddMinutes(number);
		break;
	//�� 
	case "s":
		return date.AddSeconds(number);
		break; 
	default:
		return date;
		break;
	}
}

/// ���ظ������ڵ�ָ������
double datePart(String interval, DateTime date){
	string tempInterval = interval.ToLower();
	switch (tempInterval)	{
	//�� 
	case "yyyy": 
	case "y":
		return Convert.ToDouble(date.Year);
		break;
	//�� 
	case "m":
		return Convert.ToDouble(date.Month);
		break;
	//�� 
	case "d":
		return Convert.ToDouble(date.Day);
		break;
	//Сʱ 
	case "h":
		return Convert.ToDouble(date.Hour);
		break;
	//���� 
	case "n":
		return Convert.ToDouble(date.Minute);
		break;
	//�� 
	case "s":
		return Convert.ToDouble(date.Second);
		break;  
	default:
		return 0.0000;
		break;
	}
} 

//����������������ָ���ָ������ϳ�һ���ַ�����
string join(string[]parts, string s){
	return string.Join(s, parts); 
}



//����:ASP���IIF �磺IIf(1 = 2, "a", "b") 
string IIF(bool bExp, string sVal1, string sVal2){
	//if(bExp==true){return sVal1;}else{return sVal2;}
	return bExp?sVal1:sVal2;		//��Ԫ�����
}
int IIF(bool bExp, int sVal1, int sVal2){ 
	return bExp?sVal1:sVal2;		//��Ԫ�����
}
bool IIF(bool bExp, bool sVal1, bool sVal2){ 
	return bExp?sVal1:sVal2;		//��Ԫ�����
}
//λ�� ��ASP���ͳһ
string mapPath(string path){
	path=path+"/";
	string ffPath=Server.MapPath(path  );
	ffPath=mid(ffPath,1,len(ffPath)-1);
	return ffPath;
}
void die(){
	Response.End();
}

//�����
int phpRand(int nMinimum, int nMaximum){
	Random rndObj = new Random(); 
	return  rndObj.Next(nMinimum,nMaximum); 
}


//��ʽ���ɼ۸��� 108.00 (20150806ʹ��ASPתPHP����)
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
 
//md5����
string aspMD5(string pwd, int n){
	MD5 md5 = new MD5CryptoServiceProvider();
	byte[] data = System.Text.Encoding.Default.GetBytes(pwd);//���ַ�����Ϊһ���ֽ����� 
	byte[] md5data = md5.ComputeHash(data);//����data�ֽ�����Ĺ�ϣֵ 
	md5.Clear();
	string str = "";
	for (int i = 0; i < md5data.Length - 1; i++){
		str += md5data[i].ToString("x2").PadLeft(2, '0');
	}  
	return str;
}
//���Լ���md5
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


/*һ��������   
���÷���  Response.Write(gbl.myTest());
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

//************************************************* ������
 

 
</script>
