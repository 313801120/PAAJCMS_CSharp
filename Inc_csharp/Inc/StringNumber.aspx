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

//�ж��Ƿ�Ϊ�ַ�ת��
bool isStrTransferred(string content){
    string[] splStr; int i=0; string s=""; int nCount=0;
    nCount= 0;
    bool isStrTransferred= false;
    for( i= 0 ; i<= len(content) - 1 ; i++){
        s= mid(content, len(content) - i, 1);
        if( s== "\\" ){
            nCount= nCount + 1;
        }else{
            isStrTransferred= IIF(nCount % 2== 1, true, false);
            return isStrTransferred;
        }
    }
    return isStrTransferred;
}

//������ʣ���Ϸ�������õ� 20150601
int[] getBL(int nSetWidth, int nSetHeight, int nDanFuXianZhi){
    int[] arrNSplStr=aspNArray(3);
    int nWidthZheFu= 1; //������
    int nHeightZheFu= 1; //������
    int nBFB=0; //�ٷֱ�
    int nXXFBX=0; //ÿ���ٷֱȣ���ΪҪ�ж������ܳ���10
    if( nSetWidth < 0 ){
        nSetWidth= nSetWidth * - 1;
        nWidthZheFu= -1;
    }
    if( nSetHeight < 0 ){
        nSetHeight= nSetHeight * - 1;
        nHeightZheFu= -1;
    }
    if( nSetWidth > nSetHeight ){
        nBFB= formatNumber(nSetWidth / nSetHeight, 2); ////���� �ٷֱ�
        arrNSplStr[0]= nBFB;
        arrNSplStr[1]= 1;
    }else{
        nBFB= formatNumber(nSetHeight / nSetWidth, 2); ////�߿� �ٷֱ�
        arrNSplStr[0]= 1;
        arrNSplStr[1]= nBFB;
    }
    //ÿ������ָ��ֵ������
    //if nBFB>=nDanFuXianZhi then
    nXXFBX= formatNumber(nDanFuXianZhi / nBFB, 2);
    arrNSplStr[0]= arrNSplStr[0] * nXXFBX;
    arrNSplStr[1]= arrNSplStr[1] * nXXFBX;
    //end if

    arrNSplStr[0]= arrNSplStr[0] * nWidthZheFu;
    arrNSplStr[1]= arrNSplStr[1] * nHeightZheFu;
    arrNSplStr[2]= nBFB;
    arrNSplStr[3]= getCountPage(nSetWidth, arrNSplStr[0]);
    //Call echo("page count ҳ��", arrNSplStr(3))
    arrNSplStr[3]= getCountStep(nSetWidth, nSetHeight, arrNSplStr[0], arrNSplStr[1], arrNSplStr[3]);

    return arrNSplStr;
}

//����ܲ���
int getCountStep(int nWidthStep, int nHeightStep, int nWidthBL, int nHeightBL, int nCountPage){
    int i=0;
    int getCountStep= 0;
    if( nWidthStep < 0 ){
        nWidthStep= nWidthStep * - 1;
    }
    if( nHeightStep < 0 ){
        nHeightStep= nHeightStep * - 1;
    }
    if( nWidthBL < 0 ){
        nWidthBL= nWidthBL * - 1;
    }
    if( nHeightBL < 0 ){
        nHeightBL= nHeightBL * - 1;
    }
    for( i= nCountPage - 10 ; i<= nCountPage; i++){
        //call echo(i & "��nWidthBL*i>=nWidthStep",nWidthBL*i &">="&nWidthStep    & "   |  " & nHeightBL*i &">="& nHeightStep)
        if( nWidthBL * i >= nWidthStep || nHeightBL * i >= nHeightStep ){
            getCountStep= i;
            //call echo("getCountStep",getCountStep)
        }
    }
    return getCountStep;
}


//������ĺ�������
string getChina(string content){
    int i=0; string c=""; int j=0; string s="";
    for( i= 1 ; i<= len(content); i++){
        j= asc(mid(content, i, 1));
        s= mid(content, i, 1);
        //�Ǻ����ۼ�
        if( j < 0 ){
            if((j <= -22033 && j >= -24158)== false ){
                c= c + s;
            }
        }
    }
    return c;
}
//�ж��Ƿ�������
bool isChina(string content){
    int i=0; int j=0; string s="";
    bool isChina= false;
    for( i= 1 ; i<= len(content); i++){
        j= asc(mid(content, i, 1));
        s= mid(content, i, 1);
        //�Ǻ����ۼ�
        if( j < 0 ){
            if((j <= -22033 && j >= -24158)== false ){
                isChina= true;
                return isChina;
            }
        }
    }
    return isChina;
}
//�ж��Ƿ������� (����)
bool checkChina(string content){
    return isChina(content);
}

//PHP��Randʹ��20150212


//�������棬Ϊʲô����Ϊ����д���������
int phpRnd(int nMinimum, int nMaximum){
    return phpRand(nMinimum, nMaximum);
}


//ɾ���ظ�����  20141220
string deleteRepeatStr(string content, string sSplType){
    string[] splStr; string s=""; string c="";
    c= "";
    splStr= aspSplit(content, sSplType);
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            if( inStr(sSplType + c + sSplType, sSplType + s + sSplType)== 0 ){
                c= c + s + sSplType;
            }
        }
    }
    if( c != "" ){ c= left(c, len(c) - len(sSplType)); }
    return c;
}

//�滻����N�� 20141220
string replaceN(string content, string yunStr, string replaceStr, int nNumb){
    int i=0;string sNumb="";
    sNumb= handleNumber(cStr(nNumb));			//Ϊ�˸�.netc��
    if( sNumb== "" ){
        nNumb= 1;
    }else{
        nNumb= cInt(sNumb);
    }
    for( i= 1 ; i<= nNumb; i++){
        content= replace(content, yunStr, replaceStr);
    }

    return content;
}

//asp���ڲ�0����   ���ñ���20141216
string fillZero(string content){
    string fillZero="";
    if( len(content)== 1 ){
        fillZero= "0" + content;
    }else{
        fillZero= content;
    }
    return fillZero;
}

//���ִ�Сд�滻�����ߣ�С�ƣ�д��20140925 �÷�Response.Write(CaseInsensitiveReplace("112233aabbbccddee","b","С����ţ"))
string caseInsensitiveReplace(string content, string sCheck, string sReplace){
    int nStartLen=0; int nEndLen=0; string lowerCase=""; string startStr=""; string endStr=""; string c=""; int i=0;
    c= "";
    if( lCase(sCheck)== lCase(sReplace) ){
        string caseInsensitiveReplace= content;
    }
    lowerCase= lCase(content);
    for( i= 1 ; i<= 99; i++){
        if( inStr(lowerCase, sCheck) > 0 ){
            nStartLen= inStr(lowerCase, sCheck) - 1;
            startStr= left(content, nStartLen);
            nEndLen= nStartLen + len(sCheck) + 1;
            endStr= mid(content, nEndLen,-1);
            content= startStr + sReplace + endStr;
            //Call Echo(StartLen,EndLen)
            //Call Echo(StartStr,EndStr)
            //Call Echo("Content",Content)
            lowerCase= lCase(content);
        }else{
            break;
        }
    }
    return content;
}

//������������ (2013,10,1)
int[] array_Sort(int[] arrnArray){
    int i=0; int j=0; int nMinmaxSlot=0; int nMinmax=0; int nTemp=0;
    for( i= uBound(arrnArray) ; i>= 0 ; i--){
        nMinmax= arrnArray[i];
        nMinmaxSlot= 0;
        for( j= 1 ; j<= i; j++){
            if( arrnArray[j] > nMinmax ){
                nMinmax= arrnArray[j];
                nMinmaxSlot= j;
            }
        }
        if( nMinmaxSlot != i ){
            nTemp= arrnArray[nMinmaxSlot];
            arrnArray[nMinmaxSlot]= arrnArray[i];
            arrnArray[i]= nTemp;
        }
    }
    return arrnArray;
}

//����Zip��С


////���������
string getRnd( int nCount){

    char charS; int i=0; string c="";
    for( i= 1 ; i<= nCount; i++){
        if( i % 2== 0 ){
            charS= chr((57 - 48) * rnd() + 48); //0~9
        }else if( i % 3== 0 ){
            charS= chr((90 - 65) * rnd() + 65); //A~Z
        }else{
            charS= chr((122 - 97) * rnd() + 97); //a~z
        }
        c= c + charS;
    }
    return c;
}

//������������js(20150826)
string mathRandom(){
    int i=0; string c="";
    c= "";

    for( i= 1 ; i<= 16; i++){
        c= c + cInt(cInt(rnd() * 9));
    }
    return "0." + c;
}



//���ָ��λ�����A��Z�ַ�
string getRndAZ(int nCount){
    string zd=""; int i=0; string s=""; string c="";
    c= "" ; zd= "";

    zd= "abcdefghijklmnopqrstuvwxyz" + uCase(zd);
    for( i= 1 ; i<= nCount; i++){
        s= mid(zd, phpRnd(1, len(zd)), 1);
        c= c + s;
    }
    return c;
}

//�������� ���������棩
string getRand( int nCount){
    return getRnd(nCount);
}

//��������N�� InputStr����ֵ  nMultiplier������php����һ��2014 12 02
//��� nMultiplier ������ΪС�ڵ���0���������ؿ��ַ�����
string copyStrNumb( string inputStr, int nMultiplier){
    int i=0; string s="";
    if( nMultiplier > 0 ){
        s= inputStr;
        for( i= 1 ; i<= nMultiplier - 1; i++){
            inputStr= inputStr + s;
        }
    }else{
        inputStr= "";
    }
    return inputStr;
}

//��������N��  PHP�ﺯ��
string str_Repeat( string inputStr, int nMultiplier){
    return copyStrNumb(inputStr, nMultiplier);
}

//���������
string copyStr(string inputStr, int nMultiplier){
    return copyStrNumb(inputStr, nMultiplier);
}

//���ݼ�Tab
string contentAddTab( string content, int nNumb){
    return copyStr("    ", nNumb) + join(aspSplit(content, vbCrlf()), vbCrlf() + copyStr("    ", nNumb));
}

//ɾ�����ָ���ַ�20150228 Content=DeleteEndStr(Content,2)
string deleteEndStr(string content, int nLen){
    if( content != "" ){ content= left(content, len(content) - nLen); }
    return content;
}


//StringNumber (2013,9,27)
int toNumber( int n, int nD){
    return formatNumber(n, nD, - 1);
}

//���������
string handleNumber( string content){
    int i=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", s) > 0 ){
            c= c + s;
        }
    }
    return c;
}

//�ַ�������ȡ���� 20150507
string strDrawInt( string content){
    return handleNumber(content);
}

//��������� ���ַ�������-����
string getFirstNegativeNumber( string content){
    int i=0; string s=""; string c="";
    c= "";
    content= aspTrim(content);
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( s== "-" && c== "" ){
            c= c + s;
        }else if( inStr("0123456789", s) > 0 ){
            c= c + s;
        }
    }
    if( c== "" ){ c= "0";}
    return c;
}

//����Ƿ�Ϊ��������
string checkNumberType( string content){
    return handleNumber(content);
}

//����ַ�����Ϊ��������
bool checkStrIsNumberType( string content){
    int i=0; string s="";
    bool checkStrIsNumberType= true;
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", s)== 0 ){
            checkStrIsNumberType= false;
            return checkStrIsNumberType;
        }
    }
    return checkStrIsNumberType;
}

//�������������
string handleNumberType( string content){
    int i=0; string s=""; string c="";
    c= "";
    content= aspTrim(content);
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( i== 1 && inStr("+-*/", left(content, 1)) > 0 ){
            c= c + s;
        }else if( i > 1 && s== "." ){
            c= c + s;
        }else if( inStr("0123456789", s) > 0 ){
            c= c + s;
        }
    }
    return c;
}

//������� ֻ����������� ���ҵ�һ����������Ϊ��0     20150322
string getNumber( string content){
    int i=0; string s=""; string c="";
    c= "";
    content= aspTrim(content);
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", s) > 0 ){
            if( c== "" && s== "0" ){ //���Ľ�����Ϊ�������Ӳ������ˣ�����ô����20150322
            }else{
                c= c + s;
            }
        }
    }
    return c;
}

//����Ƿ�Ϊ����
bool checkNumb(string s){
    bool checkNumb= false;
    if( inStr("0123456789.", s) > 0 ){
        checkNumb= true;
    }
    return checkNumb;
}

//�����С��������
string getDianNumb( string content){
    int i=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789.", s) > 0 ){
            c= c + s;
        }
    }
    return c;
}

//�����ҳ��
int getCountPage(int nCount, int nPageSize){
    //�Ѹ���ת����ȷ���м���20150502
    int nCountPage=0;
    if( nCount < 0 ){
        nCount= nCount * - 1;
    }
    if( nPageSize < 0 ){
        nPageSize= nPageSize * - 1;
    }
    nCountPage= fix(nCount / nPageSize);
    if( nCount % nPageSize > 0 ){ nCountPage= nCountPage + 1 ;}
    return nCountPage;
}

//��ô����ҳ��
int getPageNumb(int nRecordCount, int nPageSize){
    int n=0;
    n= cInt(nRecordCount / nPageSize);
    if( nRecordCount % nPageSize > 0 ){
        n= n + 1;
    }
    return n;
}

//�����òɼ���ҳ��
string getCaiHandleCountPage(string content){
    content= delHtml(content);
    content= handleNumber(content);
    string getCaiHandleCountPage= "";
    if( len(content) < 10 ){
        getCaiHandleCountPage= right(content, 1);
    }else if( len(content) < 200 ){
        getCaiHandleCountPage= right(content, 2);
    }
    return getCaiHandleCountPage;
}

//��òɼ�������ҳ�� 20150312
string getCaiSortCountPage( string content){
    int i=0; string s="";
    string getCaiSortCountPage= "";
    content= delHtml(content);
    content= handleNumber(content);
    for( i= 1 ; i<= 30; i++){
        s= mid(content, 1, len(i));
        if( s== cStr(i) ){
            getCaiSortCountPage= cStr(i);
            //Call Echo(i,s)
            content= right(content, len(content) - len(i));
        }
    }
    return getCaiSortCountPage;
}

//�������С֮�� Between the minimum and maximum
int minMaxBetween(int nMin, int nMax, int nValue){
    //nMin = CInt(nMin)                                                         '��С��
    //nMax = CInt(nMax)                                                         '�����
    //nValue = CInt(nValue)                                                     '��ǰ��
    int minMaxBetween= nMin;
    if( nMin > nMax ){
        minMaxBetween= nMax;
    }else if( nValue > nMin ){
        minMaxBetween= nValue;
        if( nValue > nMax ){
            minMaxBetween= nMax;
        }
    }
    return minMaxBetween;
}

//��������ļ�����������  (��FSO�ļ����Ѿ������������20141220)
string getStrFileType(string fileName){
    string c="";
    c= "";
    if( inStr(fileName, ".") > 0 ){
        c= lCase(mid(fileName, inStrRev(fileName, ".") + 1,-1));
        if( inStr(c, "?") > 0 ){
            c= mid(c, 1, inStr(c, "?") - 1);
        }
    }
    return c;
}

//���ַ�����ת����������
int val( string s){
    int val= cInt(s);
    if( s + ""== "" || is_numeric(s) ){
        val= 0;
    }
    return val;
}

//�����ַ������N��byte


//ѭ�������� AddIndent(Content,"    ")
string addIndent(string content, string indentStr){
    string[] splStr; string s=""; string c="";
    c= "";
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        c= c + indentStr + s + vbCrlf();
    }
    return trimVbCrlf(c);
}

//�������ǰ�ַ� 2014 12 12(������Ϊ����ǿ�л��ֿ�����)
string getNumberBeforeStr(string content){
    int i=0; string s=""; string c="";
    c= "";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789", s) > 0 ){ break; }
        c= c + s;
    }
    return c;
}

//�������� 20141212
//�÷�response.write makePassword(6)
string makePassword( int nMaxLen){
    string strNewPass="" ;int n=0;
    int nWhatsNext=0; int nUpper=0; int nLower=0; int nCounter=0;

    strNewPass= "";
    for( nCounter= 1 ; nCounter<= nMaxLen; nCounter++){
        nWhatsNext= cInt((1 - 0 + 1) * rnd() + 0);
        if( nWhatsNext== 0 ){
            nUpper= 90;
            nLower= 65;
        }else{
            nUpper= 57;
            nLower= 48;
        }
        n=cInt( (nUpper - nLower + 1) * rnd() + nLower);
        strNewPass= strNewPass + chr(n).ToString();
    }
    return strNewPass;
}

//����˵������������ַ�����������Сд��ĸ�����֣����������ϣ������ڸ����롣 20141212
//����˵����nMin--��������С���ȣ�ends--��������󳤶�
//�÷�'Response.Write rndCode(20, 330)
string rndCode( int nMin, int nMax){
    int nRndLen=0; int i=0;

    string rndCode= "";
    nRndLen= cInt(nMin * rnd() + nMax - nMin);
    for( i= 1 ; i<= nRndLen; i++){

        rndCode= rndCode + chr(cInt(127 * rnd() + 1).ToString());
    }
    return rndCode;
}

//�������ֻ����� ûʲô���壬���ñ��˵�  20141217
//��:CAll Rw(GetRandomPhoneNumber(41))
string getRandomPhoneNumber(int nCount){
    string s=""; string rndnum=""; int j=0; string c="" ;int n=0;
    c= "" ; rndnum= "";
    j= 1;
    while( j < nCount){

        while( len(rndnum) < 9 ){//����������ĸ���
            n= cInt((57 - 48) * rnd() + 48);
            rndnum= rndnum + chr(n).ToString();
        }
        c= c + 13 + rndnum + vbCrlf();
        rndnum= "";
        j= j + 1;
    }
    if( c != "" ){ c= left(c, len(c) - 2); }
    return c;
}

//����ַ�����


//����ת�ַ���

//�Ƴ�����(20151022)
string remoteNumber(string content){
    int i=0; string s=""; string c="";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( inStr("0123456789.", s)== 0 ){
            c= c + s;
        }
    }
    return c;
}


//================================================= �ж��������ַ� start
//��������ָ���ַ�
bool handleHaveStr(string content, string zd){
    string s=""; int i=0;
    bool handleHaveStr= false;
    for( i= 1 ; i<= len(zd); i++){
        s= mid(zd, i, 1);
        if( inStr(content, s) > 0 ){
            handleHaveStr= true;
            return handleHaveStr;
        }
    }
    return handleHaveStr;
}
//��Сд(20151224)
bool haveLowerCase(string content){
    return handleHaveStr(content, "abcdefghijklmnopqrstuvwxyz");
}
//�д�д(20151224)
bool haveUpperCase(string content){
    return handleHaveStr(content, "ABCDEFGHIJKLMNOPQRSTUVWXYZ");
}
//������(20151224)
bool haveNumber(string content){
    return handleHaveStr(content, "0123456789");
}
//�к���(20151224)
bool haveChina(string content){
    int i=0; int j=0;
    bool haveChina= false;
    for( i= 1 ; i<= len(content); i++){
        j= asc(mid(content, i, 1));
        //�Ǻ����ۼ�
        if( j < 0 ){
            if((j <= -22033 && j >= -24158)== false ){
                haveChina= true;
                return haveChina;
            }
        }
    }
    return haveChina;
}
//================================================= �ж��������ַ� end




//*************************************************
//��������EncodeJP
//��  �ã������ձ�Ƭ��������Access����ʧЧ��bug
//*************************************************
string encodeJP(string str){
    if( str== "" ){
        str= replace(str, "��", "&#12460;");
        str= replace(str, "��", "&#12462;");
        str= replace(str, "��", "&#12464;");
        str= replace(str, "��", "&#12450;");
        str= replace(str, "��", "&#12466;");
        str= replace(str, "��", "&#12468;");
        str= replace(str, "��", "&#12470;");
        str= replace(str, "��", "&#12472;");
        str= replace(str, "��", "&#12474;");
        str= replace(str, "��", "&#12476;");
        str= replace(str, "��", "&#12478;");
        str= replace(str, "��", "&#12480;");
        str= replace(str, "��", "&#12482;");
        str= replace(str, "��", "&#12485;");
        str= replace(str, "��", "&#12487;");
        str= replace(str, "��", "&#12489;");
        str= replace(str, "��", "&#12496;");
        str= replace(str, "��", "&#12497;");
        str= replace(str, "��", "&#12499;");
        str= replace(str, "��", "&#12500;");
        str= replace(str, "��", "&#12502;");
        str= replace(str, "��", "&#12502;");
        str= replace(str, "��", "&#12503;");
        str= replace(str, "��", "&#12505;");
        str= replace(str, "��", "&#12506;");
        str= replace(str, "��", "&#12508;");
        str= replace(str, "��", "&#12509;");
        str= replace(str, "��", "&#12532;");
    }
    return str;
}
</script>
