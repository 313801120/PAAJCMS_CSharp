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
//����cookie
void setCookie(string cookieName, string cookieValue, int nAddMinutes){
	Response.Cookies[cookieName].Value=cookieValue;
	Response.Cookies[cookieName].Expires=dateAdd("s",nAddMinutes, now());
	//Response.Cookies["username"].Expires=DateTime.Now.AddDays(1);
	//cookies.Expires = DateTime.Now.AddMinutes(nAddMinutes);
}
//���cookie
string getCookie(string cookieName){
	if(Request.Cookies[cookieName]!=null){
		return Server.HtmlEncode(Request.Cookies[cookieName].Value);
	}else{
		return "";
	}
}
//ɾ��cookie
void deleteCookie(string cookieName){	
	HttpCookie objCookie = new HttpCookie(cookieName.Trim()); 
	objCookie.Expires = DateTime.Now.AddYears(-1);
	HttpContext.Current.Response.Cookies.Add(objCookie);
			
}

//����cookie
void setCookie(string cookieName, int cookieValue, int nAddMinutes){ 
	string s=cookieValue.ToString();
	setCookie(cookieName,s,nAddMinutes);
}
</script>
