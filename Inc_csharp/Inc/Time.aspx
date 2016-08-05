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
//Time ʱ������� (2013,9,27)

//ʱ�䴦��
string format_Time(string timeStr, int nType){
    string y=""; string m=""; string d=""; string h=""; string mi=""; string s="" ;string c="";int nWeek=0;
    if( isDate(timeStr)== false ){ return ""; }
    y= cStr(year(timeStr));
    m= cStr(month(timeStr));
    if( len(m)== 1 ){ m= "0" + m ;}
    d= cStr(day(timeStr)); //��vb.net��Ҫ������  D = CStr(CDate(timeStr).Day)
    //��
    nWeek=fix(day(timeStr)/7);
    if( day(timeStr) % 7>0 ){
        nWeek=nWeek+1;
    }

    if( len(d)== 1 ){ d= "0" + d ;}
    h= cStr(hour(timeStr));
    if( len(h)== 1 ){ h= "0" + h ;}
    mi= cStr(minute(timeStr));
    if( len(mi)== 1 ){ mi= "0" + mi ;}
    s= cStr(second(timeStr));
    if( len(s)== 1 ){ s= "0" + s;}
    switch ( nType ){
        case 1:
        //yyyy-mm-dd hh:mm:ss
        c= y + "-" + m + "-" + d + " " + h + ":" + mi + ":" + s ;break;
        case 2:
        //yyyy-mm-dd
        c= y + "-" + m + "-" + d ;break;
        case 3:
        //hh:mm:ss
        c= h + ":" + mi + ":" + s ;break;
        case 4:
        //yyyy��mm��dd��
        c= y + "��" + m + "��" + d + "��" ;break;
        case 5:
        //yyyymmdd
        c= y + m + d ;break;
        case 6:
        //yyyymmddhhmmss
        c= y + m + d + h + mi + s ;break;
        case 7:
        //mm-dd
        c= m + "-" + d ;break;
        case 8:
        //yyyy��mm��dd��
        c= y + "��" + m + "��" + d + "��" + " " + h + ":" + mi + ":" + s ;break;
        case 9:
        //yyyy��mm��dd��Hʱmi��S�� ����
        c= y + "��" + m + "��" + d + "��" + " " + h + "ʱ" + mi + "��" + s + "�룬" + getDayStatus(h, 1) ;break;
        case 10:
        //yyyy��mm��dd��Hʱ
        c= y + "��" + m + "��" + d + "��" + h + "ʱ" ;break;
        case 11:
        //yyyy��mm��dd��Hʱmi��S��
        c= y + "��" + m + "��" + d + "��" + " " + h + "ʱ" + mi + "��" + s + "��" ;break;
        case 12:
        //yyyy��mm��dd��Hʱmi��
        c= y + "��" + m + "��" + d + "��" + " " + h + "ʱ" + mi + "��" ;break;
        case 13:
        //yyyy��mm��dd��Hʱmi�� ����
        c= m + "��" + d + "��" + " " + h + ":" + mi + " " + getDayStatus(h, 0) ;break;
        case 14:
        //yyyy��mm��dd��
        c= y + "/" + m + "/" + d ;break;
        case 15:
        //yyyy��mm�� ��1��
        c= y + "��" + m + "�� ��" + nWeek + "��";
        break;
    }
    return c;
}
//��õ�ǰʱ�ڻ����Լ���

//��õ�ǰ��״̬
string getDayStatus(string h, int nType){
    string c="";int nH=0;
    if( left(h, 1)== "0" ){
        h= right(h, 1);
    }
    nH= cInt(h);
    if( nH >= 0 && nH <= 5 ){
        c= "�賿";
    }else if( nH >= 6 && nH <= 8 ){
        c= "����";
    }else if( nH >= 9 && nH <= 12 ){
        c= "����";
    }else if( nH >= 13 && nH <= 18 ){
        c= "����";
    }else if( nH >= 19 && nH <= 24 ){
        c= "����";
    }else{
        c= "��ҹ";
    }
    if( nType== 1 ){ c= "<b>" + c + "</b>" ;}
    return c;
}
//ʱ�����
string printTimeValue( int nV){

    int nVTemp=0; string c=""; int n=0;
    nVTemp=nV;
    if( nV >= 86400 ){
        n= nV % 86400;
        nV= fix(nV / 60 / 60 / 24); //�������ǳ�24������60����Ϊһ����24��Сʱ��С�������ɵX  20150119
        c= c + nV + "��";
        nV= n;
    }
    if( nV >= 3600 ){
        n= nV % 3600;
        nV= fix(nV / 60 / 60);
        c= c + nV + "Сʱ";
        nV= n;
    }

    if( nV >= 60 ){
        n= nV % 60;
        nV= fix(nV / 60);
        c= c + nV + "��";
        nV= n;
    }
    if( nV > 0 ){
        c= c + nV + "��";
    }

    return c;
}
//������ʱ  (�ʴ���Сʱ�������ʾ)
string printAskTime(int nV){

    string c=""; int n=0;
    if( nV >= 3600 ){
        n= nV % 3600;
        nV= fix(nV / 60 / 60);
        c= c + nV + "Сʱ";
        nV= n;
        return c ;
    }
    if( nV >= 60 ){
        n= nV % 60;
        nV= fix(nV / 60);
        c= c + nV + "����";
        nV= n;
        return c ;
    }
    if( nV >= 0 ){
        c= c + nV + "����";
        return c ;
    }
    return "";
}
//���ʱ��
int getTimerSet(){
    return calculationTimer(pubTimer,now());
}
//����ʱ��
int calculationTimer(System.DateTime startTime,System.DateTime endTime){
    int n=0;
    //n = formatNumber((timer() - pubTimer) * 1000, 2, - 1) / 1000
    //calculationTimer = toNumber(n, 3)
    return dateDiff("s", startTime,endTime);
}

//���ʱ��
string getTimer(){
    return "<br>------------------------------------<br>ҳ��ִ��ʱ�䣺" + getTimerSet() + " ��";
}
//���ʱ��
string vbGetTimer(){
    string vbGetTimer= vbCrlf() + "------------------------------------" + vbCrlf() + "����ʱ�䣺" + calculationTimer(pubTimer,now());
    return "";
}
//��õ�����
string vbEchoTimer(){
    string vbEchoTimer= "����ʱ�䣺" + calculationTimer(pubTimer,now()) + " ��";
    return "";
}
//���ʱ�������
string vbRunTimer(System.DateTime startTime){
    string vbRunTimer= "����ʱ�䣺" + calculationTimer(pubTimer,now()) + " ��";
    return "";
}



//���ʱ��
string sAddTime(string timeObj,string sType,int nValue){
    string s="";
    //��
    if( sType=="s" ){
        s=timeObj+0.00001*nValue;
        //����
    }else if( sType=="n" ){
        s=timeObj+0.00060*nValue;
        //Сʱ
    }else if( sType=="h" ){
        s=timeObj+0.00001*nValue*60*60;
        //��
    }else if( sType=="d" ){
        s=timeObj+0.00001*nValue*60*60*24;
        //��
    }else if( sType=="w" ){
        s=timeObj+0.00001*nValue*60*60*24*7;
        //��
    }else if( sType=="m" ){
        s=timeObj+0.00001*nValue*60*60*24*30;
        //����
    }else if( sType=="q" ){
        s=timeObj+0.00001*nValue*60*60*24*120;
        //��
    }else if( sType=="y" ){
        s=timeObj+0.00001*nValue*60*60*24*366;
    }
    return s;
}
</script>
