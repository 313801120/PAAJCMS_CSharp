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
//��ô�������ֵ
string request(String VarName){
	string varValue = "";
	if (HttpContext.Current.Request[VarName]!=null){
		varValue = HttpContext.Current.Request[VarName].ToString();
	}
	return varValue;
}

/*
//��ñ��� ��192.168.1.133
string getIP(){
	///��ȡ���ص�IP��ַ
	string AddressIP = string.Empty;
	foreach (IPAddress _IPAddress in Dns.GetHostEntry(Dns.GetHostName()).AddressList){
		if (_IPAddress.AddressFamily.ToString() == "InterNetwork")
		{
			AddressIP = _IPAddress.ToString();
		}
	}
	return AddressIP;
}
string getIP2(){
	return getIP();
}



//ϵͳ����  ������Բ����ˣ���Ϊ.net����asp�ﹲ��һ������
string serverVariables(string action){
	action=uCase(action);
	if(action=="LOCAL_ADDR"){
		return getIP();
	}else if(action=="SERVER_NAME"){
		return Request.Url.Host;
	}else if(action=="SERVER_PORT"){ 
		return  cStr(HttpContext.Current.Request.Url.Port);
	}else if(action=="SERVER_SOFTWARE"){ 
		return cStr(Request.ServerVariables["Server_SoftWare"]);

		
	}
	return "";
} 
*/
</script>
