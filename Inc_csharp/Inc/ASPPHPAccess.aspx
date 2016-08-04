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
//ASP PHP数据操作通用文件

//判断追加Sql是加Where 还是And   sql = getWhereAnd(sql,addSql)        修改于20141007 加强版
string getWhereAnd( string sql, string addSql){
    string LCaseAddSql=""; string addType=""; string s="";
    //追加SQl为空则退出
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
            return sql + " " + addSql ; //改进必需加空格，因为前面已经删除了20160115
        }else if( left(LCaseAddSql, 6)== "where " ){
            addSql= mid(addSql, 7,-1);
        }else if( left(LCaseAddSql, 4)== "and " ){
            addSql= mid(addSql, 5,-1);
        }
        //对where 改进   20160623
        s= lCase(addSql);
        if( s != "and" && s != "or" && s != "where" ){
            sql= sql + addType + addSql;
        }
    }
    return sql;
}
//多个查询 Or 与 And        二次修改于20140703
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



//获得当前id当前页数 默认每页显示10条 20160716
string getThisIdPage(string tableName, string id, int nPageSize){
    int nCount=0;
    //在php会出错，所以加上这个
    string getThisIdPage= "1";
    if( id== "" ){
        return getThisIdPage;
    }
    //    if nPageSize = "" then        nPageSize = 10     end if 		'为了严谨，不要这个，以后要在.netc下运行通过，所以不用怕
    nCount= nGetRecordCount( tableName ," where id<=" + id);
    getThisIdPage= cStr(getCountPage(cInt(nCount), nPageSize));
    //call echo("tableName=" & tableName & "id=" & id &",ncount=" & ncount,npagesize & "               ," & getThisIdPage)
    return getThisIdPage;
}
//获得总记录
int nGetRecordCount(string tableName,string addsql){
    int nGetRecordCount=0;
    rsx = new OleDbCommand("select * from " + tableName + " " + addsql, conn).ExecuteReader();

    if( rsx.Read() ){
        nGetRecordCount=cInt(rsRecordcount("select count(*) from " + tableName + " " + addsql));
    }
    return nGetRecordCount;
}


//处理SqlServer创建语法(Access转SqlServer)
//SqlServer表的创建要求比较高，空格只能为''而不能用""还有就是数值与字符类型区别
string handleSqlServer(string content){//留空函数
    return "";
}
</script>
