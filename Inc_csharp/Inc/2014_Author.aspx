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
//����������Ϣ
void aboutAuthor(){
    string c="";
    c= c + "<pre>" + vbCrlf();
    c= c + "���ߣ�С��" + vbCrlf();
    c= c + "��ϵ��ʽ" + vbCrlf();
    c= c + "QQ��313801120" + vbCrlf();
    c= c + "���䣺313801120@qq.com" + vbCrlf();
    c= c + "΢�ţ�mq313801120" + vbCrlf();
    c= c + "����Ⱥ35915100(Ⱥ�����м�����)" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "ҵ���س�" + vbCrlf();
    c= c + "��ͨASP,PHP,VB���򿪷�������������һ��ASP��վ��̨��VB���������" + vbCrlf();
    c= c + "��������HTML��DIV��CSS��JS" + vbCrlf();
    c= c + "����ʹ��Dreamweaver��Fireworks�� Flash��Photoshop�����" + vbCrlf();
    c= c + "�Ծ�PHP��Android�ȱ������" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "��������" + vbCrlf();
    c= c + "��ѧ����ǿ����֪ʶ���ܿ졢����������ѣ�������ս��" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "��������" + vbCrlf();
    c= c + "2007��1�� �� 2012��1�� �Ϻ���ӳ����" + vbCrlf();
    c= c + "�������ݣ���վ����" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "2013---2014���Ͼ���˼�²������޹�˾" + vbCrlf();
    c= c + "�������ݣ���վ����" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "2014---����΢ս���������޹�˾" + vbCrlf();
    c= c + "�������ݣ���վ��վ�������Լ���VB������һ����վ�������������" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "��վ������" + vbCrlf();
    c= c + "http://www.863health.com/" + vbCrlf();
    c= c + "http://www.wzl99.com/" + vbCrlf();
    c= c + "http://www.jfh6666.com/" + vbCrlf();

    c= c + "</pre>" + vbCrlf();
    Response.Write(c) ; Response.End();
}

//������Ϣ
string authorInfo(string fileInfo){
    return handleAuthorInfo(fileInfo, "asp");
}
//����������Ϣ
string handleAuthorInfo(string fileInfo, string sType){
    string c=""; string phpS=""; string aspS="";
    if( sType== "php" ){
        phpS= "/";
    }else{
        aspS= "'";
    }
    c= aspS + phpS + "************************************************************" + vbCrlf();
    if( fileInfo != "" ){ c= c + aspS + "  �ļ���" + fileInfo + vbCrlf() ;}
    c= c + aspS + "���ߣ��쳾���� (��ͨASP/VB/PHP/JS/Flash��������������ϵ����)" + vbCrlf();
    c= c + aspS + "��Ȩ��Դ���빫����������;�������ʹ�á� " + vbCrlf();
    c= c + aspS + "������" + format_Time(now(), 2) + vbCrlf();
    c= c + aspS + "��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com" + vbCrlf();
    c= c + aspS + "����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���" + vbCrlf();
    c= c + aspS + "*                                    Powered by PAAJCMS " + vbCrlf();
    c= c + aspS + "************************************************************" + phpS + vbCrlf();
    return c;
}


string authorInfo2(){
    string c="";
    c= "                '''" + vbCrlf();
    c= c + "               (0 0)" + vbCrlf();
    c= c + "   +-----oOO----(_)------------+" + vbCrlf();
    c= c + "   |                           |" + vbCrlf();
    c= c + "   |    ������һ��������       |" + vbCrlf();
    c= c + "   |    QQ:313801120           |" + vbCrlf();
    c= c + "   |    sharembweb.com         |" + vbCrlf();
    c= c + "   |                           |" + vbCrlf();
    c= c + "   +------------------oOO------+" + vbCrlf();
    c= c + "              |__|__|" + vbCrlf();
    c= c + "               || ||" + vbCrlf();
    c= c + "              ooO Ooo" + vbCrlf();

    return c;
}
</script>
