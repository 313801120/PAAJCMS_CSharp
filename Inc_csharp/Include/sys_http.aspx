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
public byte[] getData(string url){
	System.Net.WebClient wc = new System.Net.WebClient();
	byte[] data = wc.DownloadData(url);
	return data;
}

//获得网址内容
string getHttpUrl(string strUrl,string set_Code){
	return getHttpPage(strUrl, set_Code);
}

/*当然getData也可以用下面的函数代替*/
/*getData2需要命名空间：System.IO和System.Net*/
public string getHttpPage(string strUrl,string set_Code){
    string strMsg = string.Empty;
	if(set_Code==""){
		set_Code="gb2312";
	}
	string sException = null;
    try{
        WebRequest request = WebRequest.Create(strUrl);
        WebResponse response = request.GetResponse();
        StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding(set_Code));
        strMsg = reader.ReadToEnd();
        reader.Close();
        reader.Dispose();
        response.Close();
    }
    catch(WebException e){
        sException=e.Message.ToString();
        Response.Write(sException);
    }
    return strMsg;
}
//获得网页状态
int getHttpUrlState(string httpurl){ 
	string statusCode;
	try{
		HttpWebRequest req = (HttpWebRequest)WebRequest.CreateDefault(new Uri(httpurl));
		req.Method = "HEAD";//设置请求方式为请求头，这样就不需要把整个网页下载下来 
		//req.Timeout = 2000; //这里设置超时时间，如果不设置，默认为10000 
		req.Timeout = 10000; //这里设置超时时间，如果不设置，默认为10000 
		HttpWebResponse res = (HttpWebResponse)req.GetResponse();
		//statusCode = res.StatusCode.ToString();
		return Convert.ToInt32(res.StatusCode);
	//使用try catch方式，如果正常，则返回OK，不正常就返回对应的错误。 
	}catch (WebException e){
		//MessageBox.Show(e.Message);
		return -1;
	} 
}
//待完善
string xmlPost(string httpurl,string content){
	return "";
}
//获得主机
string getHost(){
	return "http://"+ Request.Url.Host +"/";
}
string[] handleXmlGet(string httpUrl, string codeset){
	string[] splstr={"a","1","c","d"};
	return splstr;
}



/*

			
echo("",Request.Url.Host);
Request.ApplicationPath:                /testweb
Request.CurrentExecutionFilePath:       /testweb/default.aspx
Request.FilePath:                       /testweb/default.aspx
Request.Path:                           /testweb/default.aspx
Request.PhysicalApplicationPath:        E:\WWW\testwebRequest.PhysicalPath:                   E:\WWW\testweb\default.aspx
Request.RawUrl:                         /testweb/default.aspx
Request.Url.AbsolutePath:               /testweb/default.aspx
Request.Url.AbsoluteUrl:                http://www.test.com/testweb/default.aspx
Request.Url.Host:                       www.test.com
Request.Url.LocalPath:                  /testweb/default.aspx

*/




</script>
