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
//����վ����

//�����ַ��滻
string specialStrReplace( string content){
    content= replace(content, "\\|", "[$�����ַ�A]$");
    content= replace(content, "\\-", "[$�����ַ�B]$");
    content= replace(content, "\\,", "[$�����ַ�C]$");
    content= replace(content, "\\'", "[$�����ַ�D]$");
    content= replace(content, "\\\"", "[$�����ַ�E]$");
    return content;
}
//�������ַ��滻
string unSpecialStrReplace( string content, string startStr){
    content= replace(content, "[$�����ַ�A]$", startStr + "|");
    content= replace(content, "[$�����ַ�B]$", startStr + "-");
    content= replace(content, "[$�����ַ�C]$", startStr + ",");
    content= replace(content, "[$�����ַ�D]$", startStr + "'");
    content= replace(content, "[$�����ַ�E]$", startStr + "\"");
    return content;
}

//��Ŀ���ʹ��� ��ҳ|�ı�|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����
string handleColumnType(string columnName){
    string s="";
    switch ( columnName ){
        case "��ҳ" : s= "home";break;
        case "�ı�" : s= "text";break;
        case "��Ʒ" : s= "product";break;
        case "����" : s= "news";break;
        case "��Ƶ" : s= "video";break;
        case "����" : s= "download";break;
        case "����" : s= "case";break;
        case "����" : s= "message";break;
        case "����" : s= "feedback";break;
        case "��Ƹ" : s= "job";break;
        case "����" : s= "order";
        break;
    }
    return s;
}

</script>
