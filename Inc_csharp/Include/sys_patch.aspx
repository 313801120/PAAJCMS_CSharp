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
/*
���򲹶����������Լ�д�ó������в���ĺ���
*/ 

 

//��ʽ��ʱ��
string format_Time(System.DateTime timeStr, int nType){
	string s=timeStr.ToString();
	return format_Time(s,nType);
}
string getRandArticleId(string s,string s2){
	int n=-99;
	if(s2!=""){
		n=cInt(s2);
	}
	return getRandArticleId(s,n);
}
//ɾ��Css��ע��
string deleteCssNote(string content){
	return content;
}
void rw(int n){
	string s=n.ToString();
	rw(s);
}
void eerr(int n,int n2){
	string s=n.ToString();
	string s2=n2.ToString();
	eerr(s,s2);
}
void eerr(int n,string s2){
	string s=n.ToString(); 
	eerr(s,s2);
} 
void eerr(string s,int n2){
	string s2=n2.ToString();
	eerr(s,s2);
}


void echoRed(string s,int n2){
	string s2=n2.ToString();
	echoRed(s,s2);
}

string urlEncoding(string s){
	return s;
}
//��Ϊ��������
bool is_numeric(string s){
	return false;
}
//Ϊ��������
bool is_numeric(int s){
	return true;
}




//��õ�ǰʱ�ڻ����Լ���
string getHandleDate(int nAdd){
    return format_Time(now(), 2); 
}

//�ж����ͣ�������
string typeName(string s){
	return "";
}
//ɾ��html��ǩ��������
string delHtml(string s){
	return s;
} 
bool isNul(string s){
	return isNull(s);
}

void echo(string t,int n){
	string s=n.ToString();
	echo(t,s);
}
void echo(int n,string t){
	string s=n.ToString();
	echo(s,t);
}
void echo(int n,int n2){
	string s=n.ToString();
	string s2=n2.ToString();
	echo(s,s2);
}
void echo(string t,object o){
	string s=o.ToString();
	echo(t,s);
} 
void echo(object o1,object o2){
	string s1=o1.ToString();
	string s2=o2.ToString();
	echo(s1,s2);
} 
string getWebSite(object o1){
	string s1=o1.ToString();
	return getWebSite(s1);
}

string replaceValueParam(string s,string s2,int n){
	string s3=n.ToString();
	return replaceValueParam(s,s2,s3);
} 


string getImgJsUrl(string s,string s2){
	return "";
}

 
string regExp_Replace(string s1,string s2,string s3){
	return replace(s1,s2,s3);
} 
string cssCompression(string s,int n){
	return "";
}
string getWebTemplate( ){
	return "";
}
string getWebSkins( ){
	return "";
}
string getWebImages( ){
	return "";
}
string handleTemplateAction(string s,bool b){
	return s;
}
string addModuleReplaceArray(string s,string s2){
	return "";
}

string toUTFChar(string s){
	return s;
}
string getArrayList(string content, string startStr, string endStr,   bool isStart,bool isEnd, string sType){
	return "";
}
string getIMG(string s){
	return s;
}
string getCssHref(string s){
	return s;
}
string getJsSrc(string s){
	return s;
}
string getAURL(string s){
	return s;
}
string getATitle(string s){
	return s;
}
string getAURLTitle(string s){
	return s;
}
string URLEncoding(string s){
	return s;
}
string c16to2(string s){
	return s;
}
string c2to16(string s){
	return s;
}
string convChinese(string s){
	return s;
}
string toUTF8(string s){
	return s;
}
bool isvalidhex(string s){
	return false;
}





</script>
