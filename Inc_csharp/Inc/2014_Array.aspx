<%
/************************************************************
作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
版权：源代码免费公开，各种用途均可使用。 
创建：2016-08-05
联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
*                                    Powered by PAAJCMS 
************************************************************/
%>
<script runat="server" language="c#">
//数据处理函数库


//dim rsDataArray(99,2),s,c,i,j,k
//call handleDataArray(rsDataArray,"add","asp","asp是active server pages")
//call handleDataArray(rsDataArray,"add","asp","asp是active server pages")
//call handleDataArray(rsDataArray,"add","asp","asp是active server pages")
//call handleDataArray(rsDataArray,"del","asp","asp是active server pages")
//call handleDataArray(rsDataArray,"edit","asp","eeeeeeeeeee")
//call handleDataArray(rsDataArray,"打印","","")
//操作二维数组 当数据库来用 20160727


//内容名称排序
string contentNameSort(string content, string sType){
    string[] splStr; string[] arrayStr=aspArray(99); string fileName=""; bool isOther; string otherStr=""; string sid=""; int nIndex=0; string c=""; string s=""; int i=0; string left1="";
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            fileName= getStrFileName(s);
            isOther= true;
            left1= left(fileName, 1);
            if( inStr(fileName, "、") > 0 ){
                sid= replace(left(fileName, 2), "、", "");
                if( isNumber(sid) ){
                    nIndex= cInt(sid);
                    arrayStr[nIndex]= arrayStr[nIndex] + s + vbCrlf();
                    isOther= false;
                }
            }

            if( inStr(sType, left1)== 0 && isOther== true ){
                otherStr= otherStr + s + vbCrlf();
            }
        }
    }
    for( i= 0 ; i<= uBound(arrayStr); i++){
        c= c + arrayStr[i];
    }
    return c + otherStr;
}


//删除内容有#号列表(20150818)
string remoteContentJingHao(string content, string sSplType){
    string[] splStr; string s=""; string c="";
    splStr= aspSplit(content, sSplType);
    foreach(var eachs in splStr){
        s=eachs;
        if( left(phpTrim(s), 1) != "#" && left(phpTrim(s), 1) != "_" ){
            if( c != "" ){ c= c + sSplType ;}
            c= c + s;
        }
    }
    return c;
}
//删除数组有#号或这个_号列表(20150818)
string[] remoteArrayJingHao(string[] splStr){
    string s=""; string c=""; string sSplType="";
    sSplType= "[|-|_]";
    foreach(var eachs in splStr){
        s=eachs;
        if( left(phpTrim(s), 1) != "#" && left(phpTrim(s), 1) != "_" ){
            if( c != "" ){ c= c + sSplType ;}
            c= c + s;
        }
    }
    splStr= aspSplit(c, sSplType);
    return splStr;
}

//每个字符加指定值
string getEachStrAddValue(string content, string valueStr){
    int i=0; string s=""; string c="";
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        c= c + s + valueStr;
    }
    return c;
}
//获得值在数组里位置 20150708
int getValueInArrayID(string[] splStr, string valueStr){
    int i=0;
    int getValueInArrayID= -1;
    for( i= 0 ; i<= uBound(splStr); i++){
        if( splStr[i]== valueStr ){
            getValueInArrayID= i;
            break;
        }
    }
    return getValueInArrayID;
}
//判断值是否在数组里
bool checkValueInArray(string[] splStr, string valueStr){
    int i=0;
    bool checkValueInArray= false;
    for( i= 0 ; i<= uBound(splStr); i++){
        if( splStr[i]== valueStr ){
            checkValueInArray= true;
            break;
        }
    }
    return checkValueInArray;
}

//删除重复数组  20141220
string[] deleteRepeatArray(string[] splStr){
    string sSplType=""; string s=""; string c="";
    sSplType= "[|-|_]";
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            if( inStr(sSplType + c + sSplType, sSplType + s + sSplType)== 0 ){
                c= c + s + sSplType;
            }
        }
    }
    if( c != "" ){ c= left(c, len(c) - len(sSplType)); }
    splStr= aspSplit(c, sSplType);
    return splStr;
}
//删除重复内容 自定分割类型 20150724
string deleteRepeatContent(string content, string sSplType){
    string[] splStr; string s=""; string c="";
    splStr= aspSplit(content, sSplType);
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            if( inStr(sSplType + c + sSplType, sSplType + s + sSplType)== 0 ){
                if( c != "" ){ c= c + sSplType ;}
                c= c + s;
            }
        }
    }
    return c;
}

//获得数据总数
int getArrayCount(string content, string sSpl){
    string[] splStr;
    int getArrayCount= 0;
    if( inStr(content, sSpl) > 0 ){
        splStr= aspSplit(content, sSpl);
        getArrayCount= uBound(splStr) + 1;
    }
    return getArrayCount;
}
//随机显示内容 randomShow("1,2,3,4,5,6,7,8,9", ",", 2)
string randomShow(string content, string sSplType, int nSwitch){
    string[] splStr; string s=""; string c=""; int n=0; int i=0;

    for( i= 1 ; i<= nSwitch; i++){
        splStr= aspSplit(content, sSplType);
        foreach(var eachs in splStr){
            s=eachs;
            n= cInt(rnd() * 100);
            if( n > 50 ){
                c= c + s + sSplType;
            }else{
                c= s + sSplType + c;
            }
        }
        if( c != "" ){ c= left(c, len(c) - len(sSplType)); }
        content= c;
        c= "";
    }
    //Call Echo("C",C)
    return content;
}
//数组随机显示 ArrayRandomShow("1,2,3,4,5,6,7,8,9", 2)
string[] arrayRandomShow( string[] splStr, int nSwitch){
    string s=""; string c=""; int n=0; int i=0; string sSplType="";
    sSplType= "[|-|_]";

    for( i= 1 ; i<= nSwitch; i++){
        foreach(var eachs in splStr){
            s=eachs;
            n= cInt(rnd() * 100);
            if( n > 50 ){
                c= c + s + sSplType;
            }else{
                c= s + sSplType + c;
            }
            //Call Echo(S,N)
        }
        if( c != "" ){ c= left(c, len(c) - len(sSplType)); }
        splStr= aspSplit(c, sSplType) ; c= "";
    }
    return splStr;
}
//打印数组内容
void printArray(string[] splStr){
    int i=0; string s="";
    for( i= 0 ; i<= uBound(splStr); i++){
        s= splStr[i];
        echo(i, s);
    }
}
//显示数组内容  (辅助上面)
void echoArray(string[] splStr){
    printArray(splStr);
}
//返回打乱后的字符串。Shuffle=洗牌  2014 12 02
string str_Shuffle(string content){
    int i=0; string s=""; string c=""; int n=0;

    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        if( c== "" ){
            c= s;
        }else if( len(c)== 1 ){
            n= cInt(rnd() * 100);
            if( n > 50 ){
                c= c + s;
            }else{
                c= s + c;
            }
        }else{
            n= cInt(rnd() * len(c)) + 1;
            c= mid(c, 1, n) + s + mid(c, n + 1,-1);
        }
    }
    return c;
}
//字符打乱
string characterUpset(string content){
    return str_Shuffle(content);
}
//将字符串转换为数组   PHP里用到，暂时留着   把内容与几个字符分割成数组  如  abcefg,2     0=ab 1=ce 2=fg
string[] str_Split(string content, int nSplitLength){
    int i=0; string s=""; string c=""; int n=0; string[] arrStr=aspArray(99); int nArr=0;
    if( nSplitLength <= 0 ){ nSplitLength= 1 ;}
    n= 0 ; nArr= 0;
    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        c= c + s;
        n= n + 1;
        if( n== nSplitLength ){
            arrStr[nArr]= c;
            c= "" ; n= 0;
            nArr= nArr + 1;
        }
    }
    return arrStr;
}
//移除空值数组
string removeNullValueArray(string content, string sSplType){
    return handleArray(content, sSplType, "nonull");
}
//移除重复数组
string removeRepeatValueArray(string content, string sSplType){
    return handleArray(content, sSplType, "norepeat");
}
//处理数组
string handleArray(string content, string sSplType, string sType){
    string[] splStr; string s=""; string c=""; bool isTrue;
    sType= "|" + lCase(sType) + "|";
    splStr= aspSplit(content, sSplType);
    foreach(var eachs in splStr){
        s=eachs;
        isTrue= true;
        if( inStr(sType, "|nonull|") > 0 && isTrue== true ){
            if( s== "" ){ isTrue= false ;}
        }
        if( inStr(sType, "|norepeat|") > 0 && isTrue== true ){
            if( inStr(sSplType + c + sSplType, sSplType + s + sSplType) > 0 ){ isTrue= false ;}
        }
        if( isTrue== true ){ c= c + s + sSplType ;}
    }
    if( c != "" ){ c= left(c, len(c) - len(sSplType)); }
    return c;
}
//处理转字符(20151124)
string arrayToString(string[] splStr, string addtoStr){
    string s=""; string c="";
    foreach(var eachs in splStr){
        s=eachs;
        if( s != "" ){
            c= c + s + addtoStr;
        }
    }
    return c;
}
//测试数组数据 20141217
void testArrayData(){

    int[] nArrayData= {3, 2, 4, 1, 6, 0};
    responseArray(nArrayData, "原来顺序");
    responseArray(nArraySelectSort(nArrayData), "选择排序");
    responseArray(nArrayQuickSort(nArrayData), "快速排序");
    responseArray(nArrayInsertSort(nArrayData), "插入排序");
    responseArray(nArrayBubbleSort(nArrayData), "冒泡排序");
    responseArray(nArrayReQuickSort(nArrayData), "反序排序");
    Response.Write("<b>最 大 值：</b>" + nGetMaxArray(nArrayData) + "<hr>");
    Response.Write("<b>最 小 值：</b>" + nGetMinArray(nArrayData) + "<hr>");
}
//===================================
//输出数组
//===================================
void responseArray(int[] nArrayData, string str){
    string s=""; int i=0;
    s= "";
    Response.Write("<b>" + str + "：</b>");
    for( i= 0 ; i<= uBound(nArrayData); i++){
        s= s + nArrayData[i] + ",";
    }
    s= left(s, len(s) - 1);
    Response.Write(s);
    Response.Write("<hr>");
}
//===================================
//选择排序
//===================================
int[] nArraySelectSort(int[] nArrayData){
    int i=0; int j=0; int nK=0;
    int nBound=0; int nTemp=0;
    nBound= uBound(nArrayData);

    for( i= 0 ; i<= nBound - 1; i++){
        nK= i;
        for( j= i + 1 ; j<= nBound; j++){
            if( nArrayData[nK] > nArrayData[j] ){
                nK= j;
            }
        }
        nTemp= nArrayData[i];
        nArrayData[i]= nArrayData[nK];
        nArrayData[nK]= nTemp;
    }
    return nArrayData;
}
//===================================
//快速排序
//===================================
int[] nArrayQuickSort(int[] nArrayData){
    int i=0; int j=0;
    int nBound=0; int nTemp=0;
    nBound= uBound(nArrayData);
    for( i= 0 ; i<= nBound - 1; i++){
        for( j= i + 1 ; j<= nBound; j++){
            if( nArrayData[i] > nArrayData[j] ){
                nTemp= nArrayData[i];
                nArrayData[i]= nArrayData[j];
                nArrayData[j]= nTemp;
            }
        }
    }
    return nArrayData;
}
//===================================
//冒泡排序
//===================================
int[] nArrayBubbleSort(int[] nArrayData){
    int nBound=0;
    nBound= uBound(nArrayData);
    bool isSorted; int i=0;int j=0; int nTemp=0;
    isSorted= false;
    //    do while nBound > 0 and isSorted = false		不用这种
    for( j=1 ; j<= 9999	; j++){
        isSorted= true;
        for( i= 0 ; i<= nBound - 1; i++){
            if( nArrayData[i] > nArrayData[i + 1] ){
                nTemp= nArrayData[i];
                nArrayData[i]= nArrayData[i + 1];
                nArrayData[i + 1]= nTemp;
                isSorted= false;
            }
        }
        nBound= nBound - 1;
        if( isSorted==false ){
            break;
        }
    }

    return nArrayData;
}
//===================================
//插入排序
//===================================
int[] nArrayInsertSort(int[] nArrayData){
    int nBound=0;
    nBound= uBound(nArrayData);
    int i=0; int j=0; int nT=0;int nTemp=0;

    for( i= 1 ; i<= nBound; i++){
        nTemp= nArrayData[i];
        j= i;
        while( nT < nArrayData[j - 1] && j > 0){
            nArrayData[j]= nArrayData[j - 1];
            j= j - 1;
        }
        nArrayData[j]= nTemp;
    }
    return nArrayData;
}
//===================================
//快速排序-反序排列
//===================================
int[] nArrayReQuickSort(int[] nArrayData){
    int i=0; int nBound=0;
    nArrayData= nArrayQuickSort(nArrayData);
    int[] nArrayTemp= nArrayQuickSort(nArrayData);
    nBound= uBound(nArrayData);
    for( i= 0 ; i<= nBound; i++){
        nArrayData[i]= nArrayTemp[nBound - i];
    }
    return nArrayData;
}
//数组反向
int[] nArrayReverse(int[] nArrayData){
    return nArrayReQuickSort(nArrayData);
}

//===================================
//求数组最大值
//===================================
int nGetMaxArray(int[] nArrayData){
    int i=0; int j=0; int nBound=0; int nTemp=0;
    nArrayData= nArrayQuickSort(nArrayData);
    nBound= uBound(nArrayData);
    for( i= 0 ; i<= nBound ; i++){
        for( j= i + 1 ; j<= nBound; j++){
            if( nArrayData[j] > nArrayData[i] ){
                nTemp= nArrayData[i];
                nArrayData[i]= nArrayData[j];
                nArrayData[j]= nTemp;
            }
        }
    }
    return nArrayData[0];
}
//===================================
//求数组最小值
//===================================
int nGetMinArray(int[] nArrayData){
    int i=0; int j=0; int nBound=0; int nTemp=0;
    nArrayData= nArrayQuickSort(nArrayData);
    nBound= uBound(nArrayData);
    for( i= 0 ; i<= nBound; i++){
        for( j= i + 1 ; j<= nBound; j++){
            if( nArrayData[j] > nArrayData[i] ){
                nTemp= nArrayData[i];
                nArrayData[i]= nArrayData[j];
                nArrayData[j]= nTemp;
            }
        }
    }
    return nArrayData[nBound];
}




</script>
