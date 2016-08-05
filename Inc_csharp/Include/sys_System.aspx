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
//获得传过来的值
string request(String VarName){
	string varValue = "";
	if (HttpContext.Current.Request[VarName]!=null){
		varValue = HttpContext.Current.Request[VarName].ToString();
	}
	return varValue;
}

/*
//获得本地 如192.168.1.133
string getIP(){
	///获取本地的IP地址
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



//系统变量  这个可以不用了，因为.net里与asp里共用一个类型
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
