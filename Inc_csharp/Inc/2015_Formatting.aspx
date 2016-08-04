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
//格式化20150212


//Html格式化 简单版 加强于2014 12 09，20150709
string htmlFormatting(string content){
    return handleHtmlFormatting(content, false, 0, "");
}

//处理格式化
string handleHtmlFormatting( string content, bool isMsgBox, int nErrLevel, string action){
    string[] splStr; string s=""; string tempS=""; string lCaseS=""; string c=""; string left4Str=""; string left5Str=""; string left6Str=""; string left7Str=""; string left8Str="";
    int nLevel=0; //级别
    string elseS=""; string elseLable="";

    string[] levelArray=aspArray(299); string keyWord="";
    string lableName=""; //标签名称
    bool isJavascript; //为javascript
    bool isTextarea; //为表单文本域<textarea
    bool isPre; //为pre
    isJavascript= false; //默认javascript为假
    isTextarea= false; //表单文件域为假
    isPre= false; //默认pre为假
    nLevel= 0; //级别数

    action= "|" + action + "|"; //动作
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        tempS= s ; elseS= s;
        s= trimVbCrlf(s) ; lCaseS= lCase(s);
        //判断于20150710
        if((left(lCaseS, 8)== "<script " || left(lCaseS, 8)== "<script>") && inStr(s, "</"+"script>")== 0 && isJavascript== false ){
            isJavascript= true;
            c= c + phpTrim(tempS) + vbCrlf();
        }else if( isJavascript== true ){

            if( left(lCaseS, 9)== "</"+"script>" ){
                isJavascript= false;
                c= c + phpTrim(tempS) + vbCrlf(); //最后清除两边空格
            }else{
                c= c + tempS + vbCrlf(); //为js则显示原文本  不处理清空两边空格phptrim(tempS)
            }

            //表单文本域判断于20151019
        }else if((left(lCaseS, 10)== "<textarea " || left(lCaseS, 10)== "<textarea>") && inStr(s, "</textarea>")== 0 && isTextarea== false ){
            isTextarea= true;
            c= c + phpTrim(tempS) + vbCrlf();
        }else if( isTextarea== true ){
            c= c + phpTrim(tempS) + vbCrlf();
            if( left(lCaseS, 11)== "</textarea>" ){
                isTextarea= false;
            }
            //表单文本域判断于20151019
        }else if((left(lCaseS, 5)== "<pre " || left(lCaseS, 5)== "<pre>") && inStr(s, "</pre>")== 0 && isPre== false ){
            isPre= true;
            c= c + phpTrim(tempS) + vbCrlf();
        }else if( isPre== true ){
            c= c + tempS + vbCrlf();
            if( left(lCaseS, 6)== "</pre>" ){
                isPre= false;
            }


        }else if( s != "" && isJavascript== false && isTextarea== false ){
            left4Str= "|" + left(lCaseS, 4) + "|" ; left5Str= "|" + left(lCaseS, 5) + "|" ; left6Str= "|" + left(lCaseS, 6) + "|";
            left7Str= "|" + left(lCaseS, 7) + "|" ; left8Str= "|" + left(lCaseS, 8) + "|";

            keyWord= ""; //关键词初始清空
            lableName= ""; //标签名称
            if( inStr("|<ul>|<ul |<li>|<li |<dt>|<dt |<dl>|<dl |<dd>|<dd |<tr>|<tr |<td>|<td |", left4Str) > 0 ){
                keyWord= left4Str;
                lableName= mid(left4Str, 3, 2);
            }else if( inStr("|<div>|<div |", left5Str) > 0 ){
                keyWord= left5Str;
                lableName= mid(left5Str, 3, 3);
            }else if( inStr("|<span>|<span |<form>|<form |", left6Str) > 0 ){
                keyWord= left6Str;
                lableName= mid(left6Str, 3, 4);

            }else if( inStr("|<table>|<table |<tbody>|<tbody |", left7Str) > 0 ){
                keyWord= left7Str;
                lableName= mid(left7Str, 3, 5);

            }else if( inStr("|<center>|<center |", left8Str) > 0 ){
                keyWord= left8Str;
                lableName= mid(left8Str, 3, 6);
            }
            keyWord= aspTrim(replace(replace(keyWord, "<", ""), ">", ""));
            //call echo(KeyWord,lableName)
            //开始
            if( keyWord != "" ){
                s= copyStr("    ", nLevel) + s;
                if( right(lCaseS, 3 + len(lableName)) != "</" + lableName + ">" && inStr(lCaseS, "</" + lableName + ">")== 0 ){
                    nLevel= nLevel + 1;
                    if( nLevel >= 0 ){
                        levelArray[nLevel]= keyWord;
                    }
                }
            }else if( inStr("|</ul>|</li>|</dl>|</dt>|</dd>|</tr>|</td>|", "|" + left(lCaseS, 5) + "|") > 0 || inStr("|</div>|", "|" + left(lCaseS, 6) + "|") > 0 || inStr("|</span>|</form>|", "|" + left(lCaseS, 7) + "|") > 0 || inStr("|</table>|</tbody>|", "|" + left(lCaseS, 8) + "|") > 0 || inStr("|</center>|", "|" + left(lCaseS, 9) + "|") > 0 ){
                nLevel= nLevel - 1;
                s= copyStr("    ", nLevel) + s;
            }else{
                s= copyStr("    ", nLevel) + s;
                //最后是结束标签则减一级
                if( right(lCaseS, 6)== "</div>" ){
                    if( checkHtmlFormatting(lCaseS)== false ){
                        s= left(s, len(s) - 6);
                        nLevel= nLevel - 1;
                        s= s + vbCrlf() + copyStr("    ", nLevel) + "</div>";
                    }
                }else if( right(lCaseS, 7)== "</span>" ){
                    if( checkHtmlFormatting(lCaseS)== false ){
                        s= left(s, len(s) - 7);
                        nLevel= nLevel - 1;
                        s= s + vbCrlf() + copyStr("    ", nLevel) + "</span>";
                    }
                }else if( inStr("|</ul>|</dt>|<dl>|<dd>|", left5Str) > 0 ){
                    s= left(s, len(s) - 5);
                    nLevel= nLevel - 1;
                    s= s + vbCrlf() + copyStr("    ", nLevel) + right(lCaseS, 5);
                }


                //对   aaa</li>   这种进处理   20160106
                elseS= phpTrim(lCase(elseS));
                if( inStr(elseS, "</") > 0 ){
                    elseLable= mid(elseS, inStr(elseS, "</"),-1);
                    if( inStr("|</ul>|</li>|</dl>|</dt>|</dd>|</tr>|</td>|</div>|</span>|<form>|", "|" + elseLable + "|") > 0 && nLevel > 0 ){
                        nLevel= nLevel - 1;
                    }
                }
                //call echo("s",replace(s,"<","&lt;"))


            }
            //call echo("",ShowHtml(temps)
            c= c + s + vbCrlf();
        }else if( s== "" ){
            if( inStr(action, "|delblankline|")== 0 && inStr(action, "|删除空行|")== 0 ){//删除空行
                c= c + vbCrlf();
            }
        }
    }
    string handleHtmlFormatting= c;
    nErrLevel= nLevel; //获得错误级别
    if( nLevel != 0 && isMsgBox==true ){
        echo("HTML标签有错误", nLevel);
    }
    //Call Echo("nLevel",nLevel & "," & levelArray(nLevel))                '显示错误标题20150212
    return handleHtmlFormatting;
}

//处理闭合HTML标签(20150902)  比上面的更好用 配合上面  修复<script>部分20160719home
string formatting(string content, string action){
    int i=0; string endStr=""; string s=""; string c=""; string labelName=""; string startLabel=""; string endLabel=""; string endLabelStr=""; int nLevel=0; bool isYes; string parentLableName=""; int nTempI=0; string tempS="";
    string sNextLableName=""; //下一个标题名称
    bool isA; //是否为A链接
    bool isTextarea; //是否为多行输入文本框
    bool isScript; //脚本语言
    bool isStyle; //Css层叠样式表
    bool isPre; //是否为pre
    startLabel= "<";
    endLabel= ">";
    nLevel= 0;
    action= "|" + action + "|"; //层级
    isA= false ; isTextarea= false ; isScript= false ; isStyle= false ; isPre= false;
    content= replace(replace(content, vbCrlf(), chr(10).ToString()), vbTab(), "    ");

    for( i= 1 ; i<= len(content); i++){
        s= mid(content, i, 1);
        endStr= mid(content, i,-1);
        isYes= false;
        if( s== "<" ){
            if( inStr(endStr, ">") > 0 ){
                tempS= mid(endStr, 1, inStr(endStr, ">"));
                nTempI= i + len(tempS) - 1;
                tempS= mid(tempS, 2, len(tempS) - 2);
                if( right(tempS, 1)== "/" ){
                    tempS= phpTrim(left(tempS, len(tempS) - 1));
                }
                endStr= right(endStr, len(endStr) - len(tempS) - 2); //最后字符减去当前标签  -2是因为它有<>二个字符
                //注意之前放在labelName下面
                labelName= mid(tempS, 1, inStr(tempS + " ", " ") - 1);
                labelName= lCase(labelName);

                //call echo("labelName",labelName)
                if( labelName== "a" ){
                    isA= true;
                }else if( labelName== "/a" ){
                    isA= false;
                }else if( labelName== "textarea" ){
                    isTextarea= true;
                }else if( labelName== "/textarea" ){
                    isTextarea= false;
                }else if( labelName== "script" ){
                    isScript= true;
                }else if( labelName== "/script" ){
                    isScript= false;
                }else if( labelName== "style" ){
                    isStyle= true;
                }else if( labelName== "/style" ){
                    isStyle= false;
                }else if( labelName== "pre" ){
                    isPre= true;
                }else if( labelName== "/pre" ){
                    isPre= false;
                }else if( isScript== true ){
                    isYes= true;
                }
            }

            if( isYes== false ){
                //call echo("","11111111111")
                s= tempS;
                i= nTempI;

                endLabelStr= endLabel;
                sNextLableName= getHtmlLableName(endStr, 0);
                //不为压缩HTML
                if( inStr(action, "|ziphtml|")== 0 && isPre== false ){
                    if( isA== false ){
                        if( inStr("|a|strong|u|i|s|script|", "|" + labelName + "|")== 0 && "/" + labelName != sNextLableName && inStr("|/a|/strong|/u|/i|/s|/script|", "|" + sNextLableName + "|")== 0 ){
                            endLabelStr= endLabelStr + chr(10).ToString();
                        }
                    }
                }
                //单标签最后加个 /   20160615
                if( inStr("|br|hr|img|input|param|meta|link|", "|" + labelName + "|") > 0 ){
                    s= s + " /";
                }

                s= startLabel + s + endLabelStr;
                //不为压缩HTML
                if( inStr(action, "|ziphtml|")== 0 && isPre== false ){
                    //处理这个            aaaaa</span>
                    if( isA== false && isYes== false && left(labelName, 1)== "/" && labelName != "/script" && labelName != "/a" ){
                        //排除这种    <span>天天发团</span>     并且判断上一个字段不等于vbcrlf换行
                        if( "/" + parentLableName != labelName && right(aspTrim(c), 1) != chr(10).ToString() ){
                            s= chr(10).ToString() + s;
                        }
                    }
                }
                parentLableName= labelName;
                isYes= true;
            }
        }else if( s != "" ){
            isYes= false;
            //call echo("isPre",isPre)
            if( isPre== false ){
                if( s== chr(10).ToString() ){
                    if( isTextarea== false && isScript== false && isStyle== false ){
                        s= "";
                    }else if( isScript== true ){
                        if( inStr(action, "|zipscripthtml|") > 0 ){
                            s= " ";
                        }
                    }else if( isStyle== true ){
                        if( inStr(action, "|zipstylehtml|") > 0 ){
                            s= "";
                        }
                    }else if( isTextarea== true ){
                        if( inStr(action, "|ziptextareahtml|") > 0 ){
                            s= "";
                        }
                    }else{
                        s= chr(10).ToString();
                    }
                    //Right(Trim(c), 1) = ">")   为在压缩时用到
                }else if((right(aspTrim(c), 1)== chr(10).ToString() || right(aspTrim(c), 1)== ">") && phpTrim(s)== "" && isTextarea== false && isScript== false ){
                    s= "";
                }
            }
        }
        c= c + s;
    }
    c= replace(c, chr(10).ToString(), vbCrlf());
    return c;
}

string zipHTML(string c){
    return formatting(c, "ziphtml|zipscripthtml|zipstylehtml"); //ziphtml|zipscripthtml|zipstylehtml|ziptextareahtml
}

//检测HTML标签是否成对出现 如（<div><ul><li>aa</li></ul></div></div>）
bool checkHtmlFormatting( string content){
    string[] splStr; string s=""; string c=""; string[] splxx; int nLable=0; string lableStr="";
    content= lCase(content);
    splStr= aspSplit("ul|li|dt|dd|dl|div|span", "|");
    foreach(var eachs in splStr){
        s=eachs;
        s= phpTrim(s);
        if( s != "" ){
            nLable= 0;
            lableStr= "<" + s + " ";
            if( inStr(content, lableStr) > 0 ){
                splxx= aspSplit(content, lableStr);
                nLable= nLable + uBound(splxx);
            }
            lableStr= "<" + s + ">";
            if( inStr(content, lableStr) > 0 ){
                splxx= aspSplit(content, lableStr);
                nLable= nLable + uBound(splxx);
            }
            lableStr= "</" + s + ">";
            if( inStr(content, lableStr) > 0 ){
                splxx= aspSplit(content, lableStr);
                nLable= nLable - uBound(splxx);
            }
            //call echo(ShowHtml(lableStr),nLable)
            if( nLable != 0 ){
                bool checkHtmlFormatting= false;
                return checkHtmlFormatting;
            }
        }
    }
    return true;
}


</script>
