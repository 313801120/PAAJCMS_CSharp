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
public byte[] getData(string url){
	System.Net.WebClient wc = new System.Net.WebClient();
	byte[] data = wc.DownloadData(url);
	return data;
}

//�����ַ����
string getHttpUrl(string strUrl,string set_Code){
	return getHttpPage(strUrl, set_Code);
}

/*��ȻgetDataҲ����������ĺ�������*/
/*getData2��Ҫ�����ռ䣺System.IO��System.Net*/
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
//�����ҳ״̬
int getHttpUrlState(string httpurl){ 
	string statusCode;
	try{
		HttpWebRequest req = (HttpWebRequest)WebRequest.CreateDefault(new Uri(httpurl));
		req.Method = "HEAD";//��������ʽΪ����ͷ�������Ͳ���Ҫ��������ҳ�������� 
		//req.Timeout = 2000; //�������ó�ʱʱ�䣬��������ã�Ĭ��Ϊ10000 
		req.Timeout = 10000; //�������ó�ʱʱ�䣬��������ã�Ĭ��Ϊ10000 
		HttpWebResponse res = (HttpWebResponse)req.GetResponse();
		//statusCode = res.StatusCode.ToString();
		return Convert.ToInt32(res.StatusCode);
	//ʹ��try catch��ʽ������������򷵻�OK���������ͷ��ض�Ӧ�Ĵ��� 
	}catch (WebException e){
		//MessageBox.Show(e.Message);
		return -1;
	} 
}
//������
string xmlPost(string httpurl,string content){
	return "";
}
//�������
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
