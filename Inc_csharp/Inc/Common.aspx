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
//���ú�����ȫ (2013,9,27Option Explicit)

//�ɹ���Զ���ǻ����� 2013,10,4,��

//��ʾ�����Ի���20150312
void msgBox( string content){
    content= replace(replace(content, chr(10).ToString(), "\\n"), chr(13).ToString(), "\\n");
    Response.Write("<script>alert('" + content + "');</"+"script>");
}
//��ʾ20150729
void MBInfo(string title, string content){
    msgBox(title);
}
//��Queststring��ֵ

//ASP�Դ�����ת
void rr(string url){
    Response.Redirect(url);
}
//�滻Request.Form����
string rf(string str){
    return cStr(Request.Form[str]);
}
//��ô�ֵ
string rq(string str){
    return cStr(Request.QueryString[str]);
}
//��ô�ֵ
string rfq(string str){
    return cStr(Request[str]);
}
string ADSqlRf(string inputName){
    return replace(cStr(Request.Form[inputName]), "'", "''");
}
//�����Access���ݿ�ֵ   ��������
string ADSql(string valueStr){
    return replace(valueStr, "'", "''");
}
//�滻Response.Write����
void rw(string str){
    Response.Write(str);
}
//������ݼӻ���
void rwBr(string str){
    Response.Write(str + vbCrlf());
}
//�滻Response.Write���� + Response.End()
void rwEnd(string str){
    Response.Write(str);
    Response.End();
}
//HTML����
void htmEnd(string str){
    rwEnd(str);
}
//�滻Response.Write����+Response.End()
void die(string str){
    rwEnd(str);
}
//�滻Response.Write����
void debug(string str){
    Response.Write("<div  style=\"border:solid 1px #000000;margin-bottom:2px;\">����" + str + "</div>" + vbCrlf());
}
//����
void trace(string str){
    debug(str);
}
//������ʾ��Ϣ
void echo(string word, string str){
    echoPrompt(word, str);
}
//������ʾ��Ϣ+��ɫ
void echoRed(string word, string str){
    Response.Write("<font color=red>" + word + "</font>��" + str + "<br>");
}
//������ʾ��Ϣ+��ɫ+��
void echoRedB(string word, string str){
    Response.Write("<b><font color=red>" + word + "</font>��" + str + "</b><br>");
}
//������ʾ��Ϣ+��ɫ
void echoYellow(string word, string str){
    Response.Write("<font color=#B38704>" + word + "</font>��" + str + "<br>");
}
//������ʾ��Ϣ+��ɫ+��
void echoYellowB(string word, string str){
    Response.Write("<b><font color=#B38704>" + word + "</font>��" + str + "</b><br>");
}
//��������
void echoPrompt(string word, string str){
    Response.Write("<font color=Green>" + word + "</font>��" + str + "<br>");
}
//��������
void echoStr(string word, string str){
    echoPrompt(word, str);
}
//��ӡ���� ��ӡPHP���õ�  ��ӡ [0] => aa [1] => bb [2] => cc

//������ʾ��Ϣ ��ɫ
void setColorEcho(string color, string word, string str){
    Response.Write("<font color=" + color + ">" + word + "</font>��" + str + "<br>");
}
//������ʾ��Ϣ��ͣ
void eerr(string word, string str){
    //Response.Write(TypeName(Word) & "-" & TypeName(Str)):Response.End()
    Response.Write("<font color=red>" + word + "</font>��" + str);
    Response.End();
}
//������������
void doEvents(){
    Response.Flush();
}
//����:ASP���IIF �磺IIf(1 = 2, "a", "b")

//Hr
void hr(){
    Response.Write("<hr size='1' color='#666666'> ");
}
//BR 20160517
void br(){
    Response.Write("<br/>");
}

//����ַ��� ���ñ���20141217
//Public Sub Echo(ByVal s) : Response.Write(s) : End Sub
//����ַ�����һ�����з�
void print( string s){
    Response.Write(s + vbCrlf()) ; Response.Flush();
}
//����ַ�����һ��html���з�
void println( string s){
    Response.Write(s + "<br />" + vbCrlf()) ; Response.Flush();
}
//����ַ�������HTML��ǩתΪ��ͨ�ַ�
void printHtml( string s){
    Response.Write(replace(replace(s, "<", "&lt;"), ">", "&gt;") + vbCrlf());
}
void printlnHtml( string s){
    Response.Write(replace(replace(s, "<", "&lt;"), ">", "&gt;") + "<br />" + vbCrlf());
}
//���������ֱ�����Ϊ�ַ���(Json��ʽ)
//Public Sub PrintString(ByVal s) : Response.Write(Str.ToString(s) & VbCrLf) : End Sub
//Public Sub PrintlnString(ByVal s) : Response.Write(Str.ToString(s)) & "<br />" & VbCrLf : End Sub
//���������ʽ�����ַ���
//Public Sub PrintFormat(ByVal s, ByVal f) : Response.Write(Str.Format(s, f)) & VbCrLf : End Sub
//Public Sub PrintlnFormat(ByVal s, ByVal f) : Response.Write(Str.Format(s, f)) & "<br />" & VbCrLf : End Sub
//����ַ�������ֹ��������
void printEnd( string s){
    Response.Write(s) ; Response.End();
}




//�ж��Ƿ�һ����һ������checked,���߷��ؿ�ֵ
string setIsChecked( string str, string str2){
    return IIF(str== str2,"checked='checked'","");
}
//�ж��Ƿ�һ����һ������selected,���߷��ؿ�ֵ
string setIsSelected( string str, string str2){
    return IIF(str== str2,"selected='selected'","");
}


void doError(string s, string msg){
    //On Error Resume Next
    int nRnd=0; string c="";

    nRnd= cInt(rnd() * 29252); //29252888
    c= "<br />";
    c= c + "<div style=\"width:100%; font-size:12px;;line-height:150%\">";
    c= c + "  <label onClick=\"ERRORDIV" + nRnd + ".style.display=(ERRORDIV" + nRnd + ".style.display=='none'?'':'none')\">";
    c= c + "  <span style=\"background-color:#820222;color:#FFFFFF;height:23px;font-size:14px;cursor:pointer\">�� ���� ��ʾ��Ϣ ERROR ��</span><br />";
    c= c + "  </label>";
    c= c + "  <div id=\"ERRORDIV" + nRnd + "\" style=\"width:100%;border:1px solid #820222;padding:5px;overflow:hidden;\">";
    c= c + " <span style=\"color:#FF0000;\">��������</span> " + s + "<br />";
    c= c + " <span style=\"color:#FF0000;\">������Ϣ</span> " + msg + "<br />";
    c= c + "  </div>";
    c= c + "</div>";
    c= c + "<br />";
    Response.Write(c);
    Response.End(); //��ֹ������ֹͣ
}

//ִ��ASP�ű�


</script>
