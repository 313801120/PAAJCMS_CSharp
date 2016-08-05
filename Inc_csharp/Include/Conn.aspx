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

OleDbConnection conn=null;
OleDbDataReader rs=null;  
OleDbDataReader rsx=null;  
OleDbDataReader rss=null;  
OleDbDataReader rst=null;  
OleDbDataReader rsd=null;  
OleDbDataReader tempRs=null;  
OleDbDataReader tempRs2=null; 
OleDbDataReader rsTemp=null;  
//|rs|rsx|rss|rst|rsd|tempRs|tempRs2|rsTemp|

/*
暂且只能操作Access
*/

System.Data.IDbCommand idCommand = null;
 
//打开数据库
OleDbConnection openConn(){
	try{ 
		if(checkFile(MDBPath)==false){
			eerr("Access数据库不存在", "<a href='/Inc_csharp/install.aspx'>点击创建数据库</a>");
		}
		string strcon = "Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source="+MDBPath;  
		conn = new OleDbConnection(strcon);  
		//conn.Open();  	
		idCommand = conn.CreateCommand();
		return conn;
		
	}finally{   
		conn.Close();
	}  
	return conn;
	
}
//返回总记录   Response.Write("|"+ rsRecordcount("select count(*) from " + db_PREFIX + "admin") +"|");
int rsRecordcount(string sql){
	string lCaseSql=lCase(sql);					//让sql小写
	int nLen = inStr(lCaseSql," from ")+6;
	string s2=mid(sql,nLen);
	string tableName=mid(s2,1,inStr(s2," ")-1);
	lCaseSql=lCase(s2);
	if(inStr(lCaseSql," order by ")>0){
		nLen = inStr(lCaseSql," order by ")-1;
		s2=mid(s2,1,nLen);
	}
	string sNewSql="select count(*) from " + s2;
	//eerr(tableName,sNewSql);
	idCommand.CommandText =sNewSql;
	string s=idCommand.ExecuteScalar().ToString();
	return cInt(s);
}
//记录总数 总记录，可判断是否有记录
int getRecordCount(string tableName, string addSql) {
	string sql="Select Count(*) From "+tableName+" " +addSql;
	return rsRecordcount(sql);
}

//获得表列表
string getTableList(){  
	try{
		if (conn.State == ConnectionState.Closed){
			conn.Open();
		}
		string c="";
		DataTable dt = conn.GetSchema("Tables");
		foreach (DataRow row in dt.Rows)
		{
			if(c!=""){
				c+=vbCrlf();
			}
			c+=row[2].ToString();
		}
		return c;
	}catch (Exception e){ 
		throw e;
	}finally { 
		
	}
}
//检测表是否存在
bool checkTable(string tableName){
	if(inStr("|"+ lCase(getTableList()) +"|","|"+ lCase(tableName) +"|")>0){
		return true;
	}else{
		return false;
	}
}
//判断表，并显示是否存在否
bool checkTable_show(string tableName){
	bool isTable=checkTable(tableName);
	if(isTable==true){
		Response.Write("表[" + tableName + "]已经存在<br>");	
	}else{
        Response.Write("创建表[" + tableName + "]成功√<br>");
	}
	return isTable;
}


 
//获得字段列表
string getFieldList(string tableName){
	try{
		string c="";
		if (conn.State == ConnectionState.Closed){
			conn.Open();
		}
		using (OleDbCommand cmd = new OleDbCommand())
		{
			cmd.CommandText = "SELECT TOP 1 * FROM [" + tableName + "]";
			cmd.Connection = conn;
			OleDbDataReader dr = cmd.ExecuteReader();
			for (int i = 0; i < dr.FieldCount; i++)
			{	
				if(c!=""){
					c+=",";
				}
				c+=dr.GetName(i);
			}
		}
		return c;
	}catch (Exception e){ 
		throw e; 
	}finally{		
	}
}
//获得字段配置列表
string getFieldConfigList(string tableName){
	try{
		string c=",";string s="";
		if (conn.State == ConnectionState.Closed){
			conn.Open();
		}
		using (OleDbCommand cmd = new OleDbCommand())
		{
			cmd.CommandText = "SELECT TOP 1 * FROM [" + tableName + "]";
			cmd.Connection = conn;
			OleDbDataReader dr = cmd.ExecuteReader();
			for (int i = 1; i < dr.FieldCount; i++)
			{	
			 
				s=lCase(dr.GetDataTypeName(i));
				if(inStr(s,"_")>0){
					s=mid(s,inStr(s,"_")+1);
				}
				//echo(dr.GetName(i),s);
				if(s=="wvarchar"){
					s="||";
				}else if(s=="i4" || s=="r8" || s=="bool"){		//r8为Float   i4为int
					s="|numb|0";
				}else if(s=="date"){
					s="|time|";
				}else if(s=="wlongvarchar"){
					s="|textarea|";
				}
				c+=dr.GetName(i)+s+",";
			}
		}
		return c;
	}catch (Exception e){ 
		throw e; 
	}finally{		
	}
}
//执行sql
bool connexecute(string sql){ 
	System.Data.IDbCommand MyCommand = conn.CreateCommand();
	//执行定义的SQL查询语句
	MyCommand.CommandText = sql;
	try{
		//打开数据库连接
		if (conn.State == ConnectionState.Closed){
			conn.Open();
		}
		MyCommand.ExecuteScalar(); 
		return true;
		
	}catch (Exception ex){
		Response.Write(ex.ToString());
		return false;
	}finally{ 
	}
}
bool checkSql(string sql){
	return connexecute(sql);
} 
//创建access数据库
bool createMdb(string accessPath){
	accessPath=handlePath(accessPath);
	accessPath=replace(accessPath,"\\","\\\\"); 
	ADOX.CatalogClass cat = new ADOX.CatalogClass();
    cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + accessPath + ";");  
	return true;
}
//检测SqlServr链接是否成功
bool checkSqlServer(string connStr){  
	try{
		//string str = "server='qds157513275.my3w.com,1433';database='qds157513275_db';uid='qds157513275';pwd='313801120'"; 
		SqlConnection con = new SqlConnection(connStr); 
		con.Open(); 
		con.Close();  
		return true;
		
	}catch (Exception ex){ 
		return false;
	}finally{ 
	}
}









public OleDbConnection getConn(){
	string connstr="Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=E:\\E盘\\WEB网站\\至前网站\\Admin\\Data\\oldData.mdb";
	OleDbConnection tempconn= new OleDbConnection(connstr);
	return(tempconn);
}

bool mySQL(){
	return conn3(getConn(),"select * from admin");
}
string mySQL2(){
	return conn2(getConn(),"select count(*) from admin");
}

bool mySS(){
	return connmySS(getConn(),"select * from admin");
}

bool mySQL4(){
		   
			 OleDbConnection mycon =null;  
            OleDbDataReader myReader=null;  
            try  
            {  
                string strcon = "Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=E:\\E盘\\WEB网站\\至前网站\\Admin\\Data\\oldData.mdb";  
                mycon = new OleDbConnection(strcon);  
                mycon.Open();  
                string sql = "select * from admin ";  
                OleDbCommand mycom = new OleDbCommand(sql, mycon);  
                myReader = mycom.ExecuteReader();  
                while (myReader.Read())  
                {  
                  // Response.Write(myReader.GetString(0)+" "+myReader.GetDouble(1)+" "+myReader.GetString(2)+" "+myReader.GetString(3)+" "+myReader.GetString(4));  
					Response.Write("我的名称是="+ myReader["username"] +"<hr>");
  
                }  
  
            }  
            finally   
            {  
                myReader.Close();  
                mycon.Close();  
                  
            }  
		   
		   return false;
}

//执行SQL     
bool connmySS(System.Data.IDbConnection conn, string strquery){

 
	/*
	string connstr="Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=E:\\E盘\\WEB网站\\至前网站\\Admin\\Data\\oldData.mdb";
	OleDbConnection tempconn= new OleDbConnection(connstr);
 	OleDbCommand cmd2 = new OleDbCommand();    
	*/
	
	OleDbConnection strConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=E:\\E盘\\WEB网站\\至前网站\\Admin\\Data\\oldData.mdb");
            //建立数据库引擎连接，注意数据表（后缀为.db）应放在DEBUG文件下
            OleDbDataAdapter myda = new OleDbDataAdapter("select * from admin",strConnection);
           //建立适配器，通过SQL语句去搜索数据库
            DataSet myds = new DataSet();
           //建立数据集
            myda.Fill(myds, "雇员");
           //用FILL的方式将适配器已经连接好的数据表填充到数据集MYDS这张表
           Response.Write("<hr>显示="+myds.Tables["id"]+"<hr>");
		   
		   
		   

		   
		   
	

	//操作mysql是成功的
	string str = "server='127.0.0.1';database='WebData';uid='sa';pwd='sa'"; 
	SqlConnection con = new SqlConnection(str); 
	con.Open();                                                                                               //打开连接
	string strsql = "select * from admin";                                                        //SQL查询语句 
	SqlCommand cmd = new SqlCommand(strsql, con);                    //初始化Command对象 
	SqlDataReader rd = cmd.ExecuteReader();                                      //初始化DataReader对象 
	while (rd.Read()) 
	{ 
		Response.Write(rd[0].ToString() + "、" + rd["username"].ToString() + "<hr>");                                         //通过索引获取列 
	}
			
	return false;
}    

//执行SQL     
public bool conn3(System.Data.IDbConnection conn, string strquery){
	//使用 CreateCommand() 方法生成 Command 对象
	System.Data.IDbCommand MyCommand = conn.CreateCommand();
	//执行定义的SQL查询语句
	MyCommand.CommandText = strquery;
	try{
		//打开数据库连接
		conn.Open();
		MyCommand.ExecuteScalar();
		return true;
		
	}catch (Exception ex){
		//输出错误异常Response.Write(ex.ToString());
		return false;
	}finally{
		//关闭数据库连接
		conn.Close();
	} 
}     

//执行SQL     
string conn2(System.Data.IDbConnection conn, string strquery){
	//使用 CreateCommand() 方法生成 Command 对象
        System.Data.IDbCommand MyCommand = conn.CreateCommand();
        //执行定义的SQL查询语句
        MyCommand.CommandText = strquery;
        try
        {
            //打开数据库连接
            conn.Open();
            //定义查询的结果信息
            String MyInfo = "测试连接成功！符合查询要求的记录共有：" + MyCommand.ExecuteScalar().ToString() + "条！";
            //输出查询结果信息
            return MyInfo;
			
		 
			
        }
        catch (Exception ex)
        {
            //输出错误异常
			return "";
        }
        finally
        {
            //关闭数据库连接
            conn.Close();
        }
}      
//获得表记录总数   
int getTableCount(System.Data.IDbConnection conn, string sql){
	//使用 CreateCommand() 方法生成 Command 对象
	System.Data.IDbCommand MyCommand = conn.CreateCommand(); 
	//执行定义的SQL查询语句
	MyCommand.CommandText = sql;
	try{
		//打开数据库连接
		conn.Open();		;
		return cInt(MyCommand.ExecuteScalar().ToString());
	}catch (Exception ex){
		//输出错误异常Response.Write(ex.ToString());
	}finally{
		//关闭数据库连接
		conn.Close();
	} 
	return 0;
}        

</script>
