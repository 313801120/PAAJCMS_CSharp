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
//Form

//���ύ
void formSubmit(){
    string tableName=""; string[] splStr; string s=""; string fieldName=""; string fieldContent=""; string fieldList=""; string YZM="";
    conn=openConn();
    conn.Open();

    splStr= aspSplit(cStr(Request.Form), "&");
    tableName= rf("TableName");
    YZM= aspTrim(rf("YZM"));

    if( YZM != "" ){
        if( cStr(Session["YZM"]) != YZM ){
            javascript("����", "��֤�벻��ȷ", "");
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
    //Call Die("��������")

    javascript("����", "�ύ" + rf("DialogTitle") + "�ɹ�", "");
}

//���POST�ֶ������б� 20160226
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
