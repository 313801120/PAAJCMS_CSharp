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
//Js

//Asp����������� 20160624
string jsCodeConfusion(string content){
    string[] splStr; int i=0; bool yesJs; bool yesWord; string sx=""; string s=""; string wc=""; string zc=""; string s1=""; string aspCode=""; string upWord="";
    string upWordn=""; string tempS=""; string dimList="";
    bool yesFunction; //�����Ƿ�Ϊ��
    bool isStartFunction; //��ʼ���� Ŀ����Ϊ����function default ����������û��()   20150218
    string wcType=""; //�����ı����ͣ��� " �� '
    bool isAddToSYH; //�Ƿ��ۼ�˫����
    string beforeStr=""; string afterStr=""; string endCode=""; int nSYHCount=0;

    isStartFunction= false; //Ĭ�Ͽ�ʼ����Ϊ��
    //If nType="" Then  nType  = 0
    yesJs= false; //��ASP Ĭ��Ϊ��
    yesFunction= false; //�Ǻ��� Ĭ��Ϊ��
    yesWord= false; //�ǵ��� Ĭ��Ϊ��
    nSYHCount= 0; //˫����Ĭ��Ϊ0
    splStr= aspSplit(content, vbCrlf()); //�ָ���
    //ѭ������
    foreach(var eachs in splStr){
        s=eachs;
        //ѭ��ÿ���ַ�
        for( i= 1 ; i<= len(s); i++){
            sx= mid(s, i, 1);
            //Asp��ʼ
            if( sx== "<" && wc== "" ){ //����ı�����Ϊ�� WcΪ������� ��"<%" �ų� �޸���20140412
                if( mid(s, i + 1, 6)== "script" ){
                    yesJs= true; //ASPΪ��
                    i= i + 1; //��1�����ܼ�2��Ҫ��Ȼ<%function Test() �ͽ�ȡ����
                    sx= mid(s, i, 1);
                    aspCode= aspCode + "<";
                }
                //ASP����
            }else if( sx== "<" && mid(s, i + 1, 8)== "/script>" && wc== "" ){ //WcΪ�������
                yesJs= false; //ASPΪ��
                i= i + 1; //���ܼ�2��ֻ�ܼ�1����Ϊ���ﶨ��ASPΪ�٣���������һ����ʾ����� 'ASP����Ϊ��
                sx= mid(s, i, 8);
                aspCode= aspCode + "/script>";
            }
            if( yesJs== true ){

                beforeStr= right(replace(mid(s, 1, i - 1), " ", ""), 1); //��һ���ַ�
                afterStr= left(replace(mid(s, i + 1,-1), " ", ""), 1); //��һ���ַ�
                endCode= mid(s, i + 1,-1); //��ǰ�ַ���������� һ��
                //�����ı�
                if((sx== "\"" || sx== "'" && wcType== "") || sx== wcType || wc != "" ){
                    isAddToSYH= true;
                    //����һ�ּ򵥵ķ�����������(20150914)
                    if( isAddToSYH== true && beforeStr== "\\" ){

                        if( len(wc) >= 1 ){
                            if( isStrTransferred(wc)== true ){ //Ϊת���ַ�Ϊ��
                                //call echo(wc,isStrTransferred(wc))
                                isAddToSYH= false;
                            }
                        }else{
                            isAddToSYH= false;
                        }
                        //call echo(wc,isAddToSYH)
                    }
                    if( wc== "" ){
                        wcType= sx;
                    }

                    //˫�����ۼ�
                    if( sx== wcType && isAddToSYH== true ){ nSYHCount= nSYHCount + 1 ;}//�ų���һ���ַ�Ϊ\���ת���ַ�(20150914)


                    //�ж��Ƿ�"�����
                    if( nSYHCount % 2== 0 && beforeStr != "\\" ){
                        if( mid(s, i + 1, 1) != wcType ){
                            wc= wc + sx;
                            aspCode= aspCode + wc; //�д����ۼ�
                            //call echo("wc",wc)
                            nSYHCount= 0 ; wc= ""; //���
                            wcType= "";
                        }else{
                            wc= wc + sx;
                        }
                    }else{
                        wc= wc + sx;
                    }

                }else if( sx== "'" ){ //ע�����˳�
                    aspCode= aspCode + mid(s, i,-1);
                    break;
                    //��ĸ
                }else if( checkABC(sx)== true ||(sx== "_" && zc != "") || zc != "" ){
                    zc= zc + sx;
                    s1= lCase(mid(s + " ", i + 1, 1));
                    if( inStr("abcdefghijklmnopqrstuvwxyz0123456789", s1)== 0 && (s1== "_" && zc != "") ){//����ж�
                        tempS= mid(s, i + 1,-1);

                        if( inStr("|function|sub|", "|" + lCase(zc) + "|")>0 ){
                            //������ʼ
                            if( yesFunction== false && lCase(upWord) != "end" ){
                                yesFunction= true;
                                dimList= getFunDimName(tempS);
                                isStartFunction= true;
                            }else if( yesFunction== true && lCase(upWord)== "end" ){//�����һ������
                                yesFunction= false;
                            }
                        }else if( yesFunction== true && lCase(zc)== "var" ){
                            dimList= dimList + "," + getVarName(tempS);
                        }else if( yesFunction== true ){
                            //�ų���������ÿһ������
                            if( isStartFunction== false ){
                                zc= replaceDim2(dimList, zc);
                            }
                            isStartFunction= false;
                        }
                        upWord= zc; //��ס��ǰ����
                        aspCode= aspCode + zc;
                        zc= "";
                    }
                }else{
                    aspCode= aspCode + sx;
                }
            }else{
                aspCode= aspCode + sx;
            }
            doEvents();
        }
        aspCode= aspRTrim(aspCode); //ȥ���ұ߿ո�
        aspCode= aspCode + vbCrlf(); //Asp����
        doEvents();
    }
    return aspCode;
}


//ɾ��JSע�� 20160602
string delJsNote(string content){
    string[] splStr; string s=""; string c=""; bool isMultiLineNote; string s2="";
    isMultiLineNote= false; //����ע��Ĭ��Ϊ��
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        s2= phpTrim(s);
        if( isMultiLineNote== true ){
            if( len(s2) >= 2 ){
                if( right(s2, 2)== "*/" ){
                    isMultiLineNote= false;
                }
            }
            s= "";
        }else{
            if( left(s2, 2)== "/*" ){
                if( right(s2, 2) != "*/" ){
                    isMultiLineNote= true;
                }
                s= "";
            }else if( left(s2, 2)== "//" ){
                s= "";
            }
        }
        c= c + s + vbCrlf();
    }
    return c;
}

//JSת�������ñ���



//Զ����վ��Աͳ��2010330
//<script>document.writeln("<script src=\'http://127.0.0.1/web_soft/R.Asp?act=Stat&GoToUrl="+escape(document.referrer)+"&ThisUrl="+escape(window.location.href)+"&screen="+escape(window.screen.width+"x"+window.screen.height)+"&co="+escape(document.cookie)+" \'><\/script>");<'/script>
string showStatJSCode(string url){
    return "<script>document.writeln(\"<script src=\\'" + url + "act=Stat&GoToUrl=\"+escape(document.referrer)+\"&ThisUrl=\"+escape(window.location.href)+\"&screen=\"+escape(window.screen.width+\"x\"+window.screen.height)+\"&co=\"+escape(document.cookie)+\" \\'><\\/script>\");</"+"script>";
}


//Js��ʱ��ת Timing = ��ʱ ʱ��ⶨ ����Call Rw("�˺Ż��������" & JsTiming("����", 5))
string jsTiming(string url, int nSeconds){
    string c="";
    c= c + "<span id=mytimeidboyd>����ʱ</span>" + vbCrlf();
    c= c + "<script type=\"text/javascript\">" + vbCrlf();
    c= c + "//����Config" + vbCrlf();
    c= c + "var coutnumb" + vbCrlf();
    c= c + "coutnumb=" + nSeconds + "" + vbCrlf();
    c= c + "" + vbCrlf();
    c= c + "//��ʱ��ת" + vbCrlf();
    c= c + "function Countdown(){" + vbCrlf(); //Countdown=��������
    c= c + "    coutnumb-=1" + vbCrlf();
    c= c + "    mytimeidboyd.innerHTML=\"����ʱ<font color=#000000>\"+coutnumb+\"</font>\"" + vbCrlf();
    c= c + "    if(coutnumb<1){    " + vbCrlf();

    if( url== "back" || url== "����" ){ //��ActionΪback�Ƿ�����ҳ
        c= c + "        history.back();" + vbCrlf();
    }else{
        c= c + "        location.href='" + url + "';" + vbCrlf();
    }


    c= c + "    }else{" + vbCrlf();
    c= c + "        setTimeout(\"Countdown()\",1000);" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}setTimeout(\"Countdown()\",1)" + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    return c;
}
//JS���� Call Javascript("����", "�����ɹ�", "")
string javascript(string action, string msg, string url){
    if( msg != "" ){ msg= "alert('" + msg + "');" ;}//��Msg��Ϊ���򵯳���Ϣ
    if( action== "back" || action== "����" ){ //��ActionΪback�Ƿ�����ҳ
        Response.Write("<script>" + msg + "history.back();</"+"script>");
    }else if( url != "" ){ //��Url��Ϊ��
        Response.Write("<script>" + msg + "location.href='" + url + "';</"+"script>"); //��תUrlҳ
    }else{
        Response.Write("<script>" + msg + "</"+"script>");
    }
    Response.End();
    return "";
}
//����Ajax����ʵ��
string createAjax(){
    string c="";
    c= "<script language=\"javascript\">" + vbCrlf();
    c= c + "//AjAX XMLHTTP����ʵ��" + vbCrlf();
    c= c + "function createAjax() { " + vbCrlf();
    c= c + "    var _xmlhttp;" + vbCrlf();
    c= c + "    try {    " + vbCrlf();
    c= c + "        _xmlhttp=new ActiveXObject(\"Microsoft.XMLHTTP\");    //IE�Ĵ�����ʽ" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "    catch (e) {" + vbCrlf();
    c= c + "        try {" + vbCrlf();
    c= c + "            _xmlhttp=new XMLHttpRequest();    //FF��������Ĵ�����ʽ" + vbCrlf();
    c= c + "        }" + vbCrlf();
    c= c + "        catch (e) {" + vbCrlf();
    c= c + "            _xmlhttp=false;        //�������ʧ�ܣ�������false" + vbCrlf();
    c= c + "        }" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "    return _xmlhttp;    //����xmlhttp����ʵ��" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "//Ajax" + vbCrlf();
    c= c + "function Ajax(URL,ShowID) {  " + vbCrlf();
    c= c + "    var xmlhttp=createAjax();" + vbCrlf();
    c= c + "    if (xmlhttp) {" + vbCrlf();
    c= c + "        URL+= \"&n=\"+Math.random() " + vbCrlf();
    c= c + "        xmlhttp.open('post', URL, true);//��������" + vbCrlf();
    c= c + "        xmlhttp.setRequestHeader(\"cache-control\",\"no-cache\"); " + vbCrlf();
    c= c + "        xmlhttp.setRequestHeader(\"Content-Type\", \"application/x-www-form-urlencoded\");         " + vbCrlf();
    c= c + "        xmlhttp.onreadystatechange=function() {        " + vbCrlf();
    c= c + "            if (xmlhttp.readyState==4 && xmlhttp.status==200) {     " + vbCrlf();
    c= c + "                document.getElementById(ShowID).innerHTML = \"�������\"// unescape(xmlhttp.responseText); " + vbCrlf();
    c= c + "            }" + vbCrlf();
    c= c + "            else {                " + vbCrlf();
    c= c + "                document.getElementById(ShowID).innerHTML = \"���ڼ�����...\"" + vbCrlf();
    c= c + "            }" + vbCrlf();
    c= c + "        }" + vbCrlf();
    //c=c & "alert(document.all.TEXTContent.value)" & vbcrlf
    c= c + "        xmlhttp.send(\"Content=\"+escape(document.all.TEXTContent.value)+\"\");    " + vbCrlf();
    c= c + "        //alert(\"�������\");" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "function GetIDHTML(Root){" + vbCrlf();
    c= c + "    alert(document.all[Root].innerHTML)" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    return c;
}
//JS���߱༭
string onLineEditJS(){
    string c="";
    c= c + "<script language=\"javascript\">" + vbCrlf();
    c= c + "//��ʾ��༭���ݣ������޸ģ�ASP����������� ������2013,10,5" + vbCrlf();
    c= c + "function TestInput(Root){ " + vbCrlf();
    c= c + "    var TempContent" + vbCrlf();
    c= c + "    TempContent=\"\" " + vbCrlf();
    c= c + "    " + vbCrlf();
    c= c + "    document.all[Root].title=\"\"" + vbCrlf();
    c= c + "    if(document.all[Root].innerHTML.indexOf(\"<TEXTAREA\")==-1){" + vbCrlf();
    c= c + "            TempContent=document.all[Root].innerHTML" + vbCrlf();
    c= c + "            TempContent=TempContent.replace(/<BR><BR>/g,\"<BR>\");     " + vbCrlf();
    c= c + "            TempContent=TempContent.replace(/<BR>/g,\"\\n\");     " + vbCrlf();
    c= c + "            if(TempContent==\"&nbsp;\"){TempContent=\"\"}" + vbCrlf();
    c= c + "            document.all[Root].innerHTML=\"<textarea name=TEXT\"+Root+\" style='width:50%;height:50%' onblur=if(this.value!=''){document.all.\"+Root+\".title='����ɱ༭';document.all.\"+Root+\".innerHTML=ReplaceNToBR(this.value)}else{document.all.\"+Root+\".innerHTML='&nbsp;'};>\" + TempContent + \"</textarea>\";" + vbCrlf();
    c= c + "            document.all[\"TEXT\"+Root].focus();" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "function ReplaceNToBR(Content){" + vbCrlf();
    c= c + "    return Content.replace(/\\n/g,\"<BR>\")" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    return c;
}
//���߱༭
string editTXT(string content, string jsId){
    content= IIF(content== "", "&nbsp;", content);
    return "<span id='" + jsId + "' onClick=\"TestInput('" + jsId + "');\" title='����ɱ༭'>" + content + "</span>";
}
//���߱༭  (����)
string onLineEdit(string content, string jsId){
    return editTXT(content, jsId);
}
//****************************************************
//��������JSGoTo
//��  �ã���ʾ�ı�
//ʱ  �䣺2013��12��14��
//��  ����Url
//*       SetTime
//����ֵ���ַ���
//��  �ԣ�Call Echo("���Ժ��� JSGoTo", JSGoTo("", "",""))
//****************************************************
string jsGoTo(string title, string url, int nSetTime){
    string c="";
    if( title== "" ){ title= "��ӳɹ�" ;}
    //if nSetTime = "" then nSetTime = 4                                                'Ĭ��Ϊ4��
    c= c + "<script>" + vbCrlf();
    c= c + "//ͨ�ö�ʱ�� �磺MyTimer('Show', 'alert(1+1)', 5)" + vbCrlf();
    c= c + "var StopTimer = \"\"" + vbCrlf();
    c= c + "function MyTimer(ID, ActionStr,TimeNumb){" + vbCrlf();
    c= c + "    if(StopTimer == \"ֹͣ\" || StopTimer == \"ֹͣ��ʱ��\"){" + vbCrlf();
    c= c + "        StopTimer = \"\"" + vbCrlf();
    c= c + "        return false" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "    TimeNumb--" + vbCrlf();
    c= c + "    document.all[ID].innerHTML = \"����ʱ��\" + TimeNumb" + vbCrlf();
    c= c + "    if(TimeNumb<1){" + vbCrlf();
    c= c + "        setTimeout(ActionStr,100);" + vbCrlf();
    c= c + "    }else{" + vbCrlf();
    c= c + "        setTimeout(\"MyTimer('\"+ID+\"', '\"+ActionStr+\"',\"+TimeNumb+\")\",1000);" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "function GotoURL(){" + vbCrlf();
    c= c + "    location.href='" + url + "'" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    c= c + "<div id=\"Show\">Loading...</div><script>MyTimer('Show', 'GotoURL()', " + nSetTime + ")</"+"script>" + vbCrlf();
    return c;
}

//JSͼƬ����
string jsPhotoScroll(string id, string width, string height){
    string c="";
    c= c + "<script type=\"text/javascript\">" + vbCrlf();
    c= c + "    var marqueeB = new Marquee(\"" + id + "\")    " + vbCrlf();
    c= c + "    marqueeB.Direction =2;" + vbCrlf();
    c= c + "    marqueeB.Step = 1;" + vbCrlf();
    c= c + "    marqueeB.Width = " + width + ";" + vbCrlf();
    c= c + "    marqueeB.Height = " + height + ";" + vbCrlf();
    c= c + "    marqueeB.Timer = 1;" + vbCrlf();
    c= c + "    marqueeB.DelayTime = 0;" + vbCrlf();
    c= c + "    marqueeB.WaitTime = 0;" + vbCrlf();
    c= c + "    marqueeB.ScrollStep = 20;" + vbCrlf();
    c= c + "    marqueeB.Start();    " + vbCrlf();
    c= c + "</"+"script>" + vbCrlf();
    return c;
}
//ͼƬ����������ݲ��ã�
string photoLeftScroll(string demo, string demo1, string demo2){
    string c="";
    c= c + "<!--ͼƬ�����ַ�����-->" + vbCrlf();
    c= c + "<script language=\"javascript\">" + vbCrlf();
    c= c + "var speed=30" + vbCrlf();
    c= c + "" + demo2 + ".innerHTML=" + demo1 + ".innerHTML" + vbCrlf();
    c= c + "function Marquee(){" + vbCrlf();
    c= c + "    if(" + demo2 + ".offsetWidth-" + demo + ".scrollLeft<=0)" + vbCrlf();
    c= c + "        " + demo + ".scrollLeft-=" + demo1 + ".offsetWidth" + vbCrlf();
    c= c + "    else{" + vbCrlf();
    c= c + "        " + demo + ".scrollLeft++" + vbCrlf();
    c= c + "    }" + vbCrlf();
    c= c + "}" + vbCrlf();
    c= c + "var MyMar=setInterval(Marquee,speed)" + vbCrlf();
    c= c + "" + demo + ".onmouseover=function() {clearInterval(MyMar)}" + vbCrlf();
    c= c + "" + demo + ".onmouseout=function() {MyMar=setInterval(Marquee,speed)}" + vbCrlf();
    c= c + "</"+"script> " + vbCrlf();
    return c;
}
</script>
