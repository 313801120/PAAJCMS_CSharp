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
//Form

//表单提交
void formSubmit(){
    string tableName=""; string[] splStr; string s=""; string fieldName=""; string fieldContent=""; string fieldList=""; string YZM="";
    conn=openConn();
    conn.Open();

    splStr= aspSplit(cStr(Request.Form), "&");
    tableName= rf("TableName");
    YZM= aspTrim(rf("YZM"));

    if( YZM != "" ){
        if( cStr(Session["YZM"]) != YZM ){
            javascript("返回", "验证码不正确", "");
            Response.End();
        }
    }

    fieldList= lCase(getFieldList(tableName));
    //Call Echo("FieldList",FieldList)
    //Call Echo("TableName", TableName)
    rs = new OleDbCommand("Select * From [" + tableName + "]", conn).ExecuteReader();


    foreach(var eachs in splStr){
        s=eachs;
        fieldName= lCase(mid(s, 1, inStr(s, "=") - 1));
        //FieldContent = Mid(S,InStr(S,"=")+1)
        fieldContent= rf(fieldName);
        if( inStr("," + fieldList + ",", "," + fieldName + ",") > 0 ){

        }
        //Call Echo(FieldName,FieldContent & "," & unescape(FieldContent))
    }

    //Call Echo("DialogTitle",Rf("DialogTitle"))
    //Call Die("留言内容")

    javascript("返回", "提交" + rf("DialogTitle") + "成功", "");
}

//获得POST字段名称列表 20160226
string getFormFieldList(){
    string s=""; string c=""; string[] splStr; string fieldName="";
    splStr= aspSplit(cStr(Request.Form), "&");
    foreach(var eachs in splStr){
        s=eachs;
        fieldName= lCase(mid(s, 1, inStr(s, "=") - 1));
        if( c != "" ){ c= c + "|" ;}
        c= c + fieldName;
    }
    return c;
}

</script>
