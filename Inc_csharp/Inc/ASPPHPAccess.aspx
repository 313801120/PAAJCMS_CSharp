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
//ASP PHP���ݲ���ͨ���ļ�

//�ж�׷��Sql�Ǽ�Where ����And   sql = getWhereAnd(sql,addSql)        �޸���20141007 ��ǿ��
string getWhereAnd( string sql, string addSql){
    string LCaseAddSql=""; string addType=""; string s="";
    //׷��SQlΪ�����˳�
    if( aspTrim(addSql)== "" ){ return sql ; }
    if( inStr(lCase(sql), " where ") > 0 ){
        addType= " And ";
    }else{
        addType= " Where ";
    }
    if( addSql != "" ){
        addSql= aspTrim(addSql);
        LCaseAddSql= lCase(addSql);
        if( left(LCaseAddSql, 6)== "order " || left(LCaseAddSql, 6)== "group " ){
            return sql + " " + addSql ; //�Ľ�����ӿո���Ϊǰ���Ѿ�ɾ����20160115
        }else if( left(LCaseAddSql, 6)== "where " ){
            addSql= mid(addSql, 7,-1);
        }else if( left(LCaseAddSql, 4)== "and " ){
            addSql= mid(addSql, 5,-1);
        }
        //��where �Ľ�   20160623
        s= lCase(addSql);
        if( s != "and" && s != "or" && s != "where" ){
            sql= sql + addType + addSql;
        }
    }
    return sql;
}
//�����ѯ Or �� And        �����޸���20140703
string orAndSearch(string addSql, string seectField, string searchValue){
    string[] splStr; string s=""; string c="";
    searchValue= regExp_Replace(searchValue, " or ", " Or ");
    searchValue= regExp_Replace(searchValue, " and ", " And ");
    if( inStr(searchValue, " Or ") > 0 ){
        splStr= aspSplit(searchValue, " Or ");
        foreach(var eachs in splStr){
            s=eachs;
            if( s != "" ){
                if( c != "" ){ c= c + " Or " ;}
                c= c + " " + seectField + " Like '%" + s + "%'";
            }
        }
    }else if( inStr(searchValue, " And ") > 0 ){
        splStr= aspSplit(searchValue, " And ");
        foreach(var eachs in splStr){
            s=eachs;
            if( s != "" ){
                if( c != "" ){ c= c + " And " ;}
                c= c + " " + seectField + " Like '%" + s + "%'";
            }
        }
    }else if( searchValue != "" ){
        splStr= aspSplit(searchValue, " And ");
        foreach(var eachs in splStr){
            s=eachs;
            if( s != "" ){
                if( c != "" ){ c= c + " And " ;}
                c= c + " " + seectField + " Like '%" + s + "%'";
            }
        }
    }
    if( c != "" ){
        if( inStr(lCase(addSql), " where ")== 0 ){
            c= " Where " + c;
        }else{
            c= " And " + c;
        }
        addSql= addSql + c;
    }
    return addSql;
}



//��õ�ǰid��ǰҳ�� Ĭ��ÿҳ��ʾ10�� 20160716
string getThisIdPage(string tableName, string id, int nPageSize){
    int nCount=0;
    //��php��������Լ������
    string getThisIdPage= "1";
    if( id== "" ){
        return getThisIdPage;
    }
    //    if nPageSize = "" then        nPageSize = 10     end if 		'Ϊ���Ͻ�����Ҫ������Ժ�Ҫ��.netc������ͨ�������Բ�����
    nCount= nGetRecordCount( tableName ," where id<=" + id);
    getThisIdPage= cStr(getCountPage(cInt(nCount), nPageSize));
    //call echo("tableName=" & tableName & "id=" & id &",ncount=" & ncount,npagesize & "               ," & getThisIdPage)
    return getThisIdPage;
}
//����ܼ�¼
int nGetRecordCount(string tableName,string addsql){
    int nGetRecordCount=0;
    rsx = new OleDbCommand("select * from " + tableName + " " + addsql, conn).ExecuteReader();

    if( rsx.Read() ){
        nGetRecordCount=cInt(rsRecordcount("select count(*) from " + tableName + " " + addsql));
    }
    return nGetRecordCount;
}


//����SqlServer�����﷨(AccessתSqlServer)
//SqlServer��Ĵ���Ҫ��Ƚϸߣ��ո�ֻ��Ϊ''��������""���о�����ֵ���ַ���������
string handleSqlServer(string content){//���պ���
    return "";
}
</script>
