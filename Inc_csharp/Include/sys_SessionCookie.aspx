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
//设置cookie
void setCookie(string cookieName, string cookieValue, int nAddMinutes){
	Response.Cookies[cookieName].Value=cookieValue;
	Response.Cookies[cookieName].Expires=dateAdd("s",nAddMinutes, now());
	//Response.Cookies["username"].Expires=DateTime.Now.AddDays(1);
	//cookies.Expires = DateTime.Now.AddMinutes(nAddMinutes);
}
//获得cookie
string getCookie(string cookieName){
	if(Request.Cookies[cookieName]!=null){
		return Server.HtmlEncode(Request.Cookies[cookieName].Value);
	}else{
		return "";
	}
}
//删除cookie
void deleteCookie(string cookieName){	
	HttpCookie objCookie = new HttpCookie(cookieName.Trim()); 
	objCookie.Expires = DateTime.Now.AddYears(-1);
	HttpContext.Current.Response.Cookies.Add(objCookie);
			
}

//设置cookie
void setCookie(string cookieName, int cookieValue, int nAddMinutes){ 
	string s=cookieValue.ToString();
	setCookie(cookieName,s,nAddMinutes);
}
</script>
