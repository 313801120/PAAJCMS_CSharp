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
����ֻ�ܲ���Access
*/

System.Data.IDbCommand idCommand = null;
 
//�����ݿ�
OleDbConnection openConn(){
	try{ 
		if(checkFile(MDBPath)==false){
			eerr("Access���ݿⲻ����", "<a href='/Inc_csharp/install.aspx'>����������ݿ�</a>");
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
//�����ܼ�¼   Response.Write("|"+ rsRecordcount("select count(*) from " + db_PREFIX + "admin") +"|");
int rsRecordcount(string sql){
	string lCaseSql=lCase(sql);					//��sqlСд
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
//��¼���� �ܼ�¼�����ж��Ƿ��м�¼
int getRecordCount(string tableName, string addSql) {
	string sql="Select Count(*) From "+tableName+" " +addSql;
	return rsRecordcount(sql);
}

//��ñ��б�
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
//�����Ƿ����
bool checkTable(string tableName){
	if(inStr("|"+ lCase(getTableList()) +"|","|"+ lCase(tableName) +"|")>0){
		return true;
	}else{
		return false;
	}
}
//�жϱ�����ʾ�Ƿ���ڷ�
bool checkTable_show(string tableName){
	bool isTable=checkTable(tableName);
	if(isTable==true){
		Response.Write("��[" + tableName + "]�Ѿ�����<br>");	
	}else{
        Response.Write("������[" + tableName + "]�ɹ���<br>");
	}
	return isTable;
}


 
//����ֶ��б�
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
//����ֶ������б�
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
				}else if(s=="i4" || s=="r8" || s=="bool"){		//r8ΪFloat   i4Ϊint
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
//ִ��sql
bool connexecute(string sql){ 
	System.Data.IDbCommand MyCommand = conn.CreateCommand();
	//ִ�ж����SQL��ѯ���
	MyCommand.CommandText = sql;
	try{
		//�����ݿ�����
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
//����access���ݿ�
bool createMdb(string accessPath){
	accessPath=handlePath(accessPath);
	accessPath=replace(accessPath,"\\","\\\\"); 
	ADOX.CatalogClass cat = new ADOX.CatalogClass();
    cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + accessPath + ";");  
	return true;
}
//���SqlServr�����Ƿ�ɹ�
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
	string connstr="Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=E:\\E��\\WEB��վ\\��ǰ��վ\\Admin\\Data\\oldData.mdb";
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
                string strcon = "Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=E:\\E��\\WEB��վ\\��ǰ��վ\\Admin\\Data\\oldData.mdb";  
                mycon = new OleDbConnection(strcon);  
                mycon.Open();  
                string sql = "select * from admin ";  
                OleDbCommand mycom = new OleDbCommand(sql, mycon);  
                myReader = mycom.ExecuteReader();  
                while (myReader.Read())  
                {  
                  // Response.Write(myReader.GetString(0)+" "+myReader.GetDouble(1)+" "+myReader.GetString(2)+" "+myReader.GetString(3)+" "+myReader.GetString(4));  
					Response.Write("�ҵ�������="+ myReader["username"] +"<hr>");
  
                }  
  
            }  
            finally   
            {  
                myReader.Close();  
                mycon.Close();  
                  
            }  
		   
		   return false;
}

//ִ��SQL     
bool connmySS(System.Data.IDbConnection conn, string strquery){

 
	/*
	string connstr="Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=E:\\E��\\WEB��վ\\��ǰ��վ\\Admin\\Data\\oldData.mdb";
	OleDbConnection tempconn= new OleDbConnection(connstr);
 	OleDbCommand cmd2 = new OleDbCommand();    
	*/
	
	OleDbConnection strConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=E:\\E��\\WEB��վ\\��ǰ��վ\\Admin\\Data\\oldData.mdb");
            //�������ݿ��������ӣ�ע�����ݱ���׺Ϊ.db��Ӧ����DEBUG�ļ���
            OleDbDataAdapter myda = new OleDbDataAdapter("select * from admin",strConnection);
           //������������ͨ��SQL���ȥ�������ݿ�
            DataSet myds = new DataSet();
           //�������ݼ�
            myda.Fill(myds, "��Ա");
           //��FILL�ķ�ʽ���������Ѿ����Ӻõ����ݱ���䵽���ݼ�MYDS���ű�
           Response.Write("<hr>��ʾ="+myds.Tables["id"]+"<hr>");
		   
		   
		   

		   
		   
	

	//����mysql�ǳɹ���
	string str = "server='127.0.0.1';database='WebData';uid='sa';pwd='sa'"; 
	SqlConnection con = new SqlConnection(str); 
	con.Open();                                                                                               //������
	string strsql = "select * from admin";                                                        //SQL��ѯ��� 
	SqlCommand cmd = new SqlCommand(strsql, con);                    //��ʼ��Command���� 
	SqlDataReader rd = cmd.ExecuteReader();                                      //��ʼ��DataReader���� 
	while (rd.Read()) 
	{ 
		Response.Write(rd[0].ToString() + "��" + rd["username"].ToString() + "<hr>");                                         //ͨ��������ȡ�� 
	}
			
	return false;
}    

//ִ��SQL     
public bool conn3(System.Data.IDbConnection conn, string strquery){
	//ʹ�� CreateCommand() �������� Command ����
	System.Data.IDbCommand MyCommand = conn.CreateCommand();
	//ִ�ж����SQL��ѯ���
	MyCommand.CommandText = strquery;
	try{
		//�����ݿ�����
		conn.Open();
		MyCommand.ExecuteScalar();
		return true;
		
	}catch (Exception ex){
		//��������쳣Response.Write(ex.ToString());
		return false;
	}finally{
		//�ر����ݿ�����
		conn.Close();
	} 
}     

//ִ��SQL     
string conn2(System.Data.IDbConnection conn, string strquery){
	//ʹ�� CreateCommand() �������� Command ����
        System.Data.IDbCommand MyCommand = conn.CreateCommand();
        //ִ�ж����SQL��ѯ���
        MyCommand.CommandText = strquery;
        try
        {
            //�����ݿ�����
            conn.Open();
            //�����ѯ�Ľ����Ϣ
            String MyInfo = "�������ӳɹ������ϲ�ѯҪ��ļ�¼���У�" + MyCommand.ExecuteScalar().ToString() + "����";
            //�����ѯ�����Ϣ
            return MyInfo;
			
		 
			
        }
        catch (Exception ex)
        {
            //��������쳣
			return "";
        }
        finally
        {
            //�ر����ݿ�����
            conn.Close();
        }
}      
//��ñ��¼����   
int getTableCount(System.Data.IDbConnection conn, string sql){
	//ʹ�� CreateCommand() �������� Command ����
	System.Data.IDbCommand MyCommand = conn.CreateCommand(); 
	//ִ�ж����SQL��ѯ���
	MyCommand.CommandText = sql;
	try{
		//�����ݿ�����
		conn.Open();		;
		return cInt(MyCommand.ExecuteScalar().ToString());
	}catch (Exception ex){
		//��������쳣Response.Write(ex.ToString());
	}finally{
		//�ر����ݿ�����
		conn.Close();
	} 
	return 0;
}        

</script>
