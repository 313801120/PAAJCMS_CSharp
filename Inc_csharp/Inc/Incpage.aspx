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
//dim nCount,nPageSize,maxpage,page,x,i,PageControl


//页控制  记录总数  每页显示数  当前面 (2015117)   webPageControl(59,12,1,"http://www.baidu.com")
string webPageControl(int nRecrodCount, int nPageSize, string sPage, string configPageUrl, string action){
    string c=""; int nCountPage=0; int i=0; int nDisplay=0; int nDispalyOK=0; string sTemp=""; string cPages=""; string url=""; string selStr="";int nPage=0;string s="";
    int nPreviousPage=0; int nNextPage=0; //定义上一页，下一页
    bool isDisplayTip; //是否显示提示翻页信息
    isDisplayTip= true;

    string sPageStart=""; string sPageEnd=""; string sHomePage=""; string sHomePageFocus=""; string sUpPage=""; string sUpPageFocus=""; string sNextPage=""; string sNextPageFocus=""; string sForPage=""; string sForPageFocus=""; string sTailPage=""; string sTailPageFocus="";
    if( action != "" ){
        sPageStart= getStrCut(action, "[sPageStart]", "[/sPageStart]", 2); //页头部分
        sPageEnd= getStrCut(action, "[sPageEnd]", "[/sPageEnd]", 2); //页尾部分
        sHomePage= getStrCut(action, "[sHomePage]", "[/sHomePage]", 2); //首页
        sHomePageFocus= getStrCut(action, "[sHomePageFocus]", "[/sHomePageFocus]", 2); //首页交点
        sUpPage= getStrCut(action, "[sUpPage]", "[/sUpPage]", 2); //上一页
        sUpPageFocus= getStrCut(action, "[sUpPageFocus]", "[/sUpPageFocus]", 2); //上一页交点
        sNextPage= getStrCut(action, "[sNextPage]", "[/sNextPage]", 2); //下一页
        sNextPageFocus= getStrCut(action, "[sNextPageFocus]", "[/sNextPageFocus]", 2); //下一页交点
        sForPage= getStrCut(action, "[sForPage]", "[/sForPage]", 2); //循环页
        sForPageFocus= getStrCut(action, "[sForPageFocus]", "[/sForPageFocus]", 2); //循环页交点
        sTailPage= getStrCut(action, "[sTailPage]", "[/sTailPage]", 2); //最后页
        sTailPageFocus= getStrCut(action, "[sTailPageFocus]", "[/sTailPageFocus]", 2); //最后页交点


    }
    //页头部分
    if( sPageStart== "" ){
        sPageStart= "<ul class=\"pagecontrolwrap\">" + vbCrlf() + "<li class=\"pageinfo\">共[$nRecrodCount$]条 [$nPage$]/[$nCountPage$]页</li>" + vbCrlf();
    }
    //页尾部分
    if( sPageEnd== "" ){
        sPageEnd= "</ul><div class=\"clear\"></div>" + vbCrlf();
    }
    //首页
    if( sHomePage== "" ){
        sHomePage= "<li><a href=\"[$url$]\">首页</a></li>" + vbCrlf();
    }
    //首页交点
    if( sHomePageFocus== "" ){
        sHomePageFocus= "<li class=\"pageli\">首页</li>" + vbCrlf();
    }
    //上一页
    if( sUpPage== "" ){
        sUpPage= "<li><a href=\"[$url$]\">上一页</a></li>" + vbCrlf();
    }
    //上一页交点
    if( sUpPageFocus== "" ){
        sUpPageFocus= "<li class=\"pageli\">上一页</li>" + vbCrlf();
    }
    //下一页
    if( sNextPage== "" ){
        sNextPage= "<li><a href=\"[$url$]\">下一页</a></li>" + vbCrlf();
    }
    //下一页交点
    if( sNextPageFocus== "" ){
        sNextPageFocus= "<li class=\"pageli\">下一页</li>" + vbCrlf();
    }
    //循环页
    if( sForPage== "" ){
        sForPage= "<li><a href=\"[$url$]\">[$i$]</a></li>" + vbCrlf();
    }
    //循环页交点
    if( sForPageFocus== "" ){
        sForPageFocus= "<li class=\"pagefocus\">[$i$]</li>" + vbCrlf();
    }
    //最后页
    if( sTailPage== "" ){
        sTailPage= "<li><a href=\"[$url$]\">末页</a></li>" + vbCrlf();
    }
    //最后页交点
    if( sTailPageFocus== "" ){
        sTailPageFocus= "<li class=\"pageli\">末页</li>" + vbCrlf();
    }
    //测试时用到20160630
    if( 1== 2 ){
        c= "[sPageStart]" + vbCrlf() + sPageStart + "[/sPageStart]" + vbCrlf() + vbCrlf();
        c= c + "[sHomePage]" + vbCrlf() + sHomePage + "[/sHomePage]" + vbCrlf() + vbCrlf();
        c= c + "[sHomePageFocus]" + vbCrlf() + sHomePageFocus + "[/sHomePageFocus]" + vbCrlf() + vbCrlf();

        c= c + "[sUpPage]" + vbCrlf() + sUpPage + "[/sUpPage]" + vbCrlf() + vbCrlf();
        c= c + "[sUpPageFocus]" + vbCrlf() + sUpPageFocus + vbCrlf() + "[/sUpPageFocus]" + vbCrlf();


        c= c + "[sForPage]" + vbCrlf() + sForPage + "[/sForPage]" + vbCrlf() + vbCrlf();
        c= c + "[sForPageFocus]" + vbCrlf() + sForPageFocus + "[/sForPageFocus]" + vbCrlf() + vbCrlf();


        c= c + "[sNextPage]" + vbCrlf() + sNextPage + "[/sNextPage]" + vbCrlf() + vbCrlf();
        c= c + "[sNextPageFocus]" + vbCrlf() + sNextPageFocus + "[/sNextPageFocus]" + vbCrlf() + vbCrlf();

        c= c + "[sTailPage]" + vbCrlf() + sTailPage + "[/sTailPage]" + vbCrlf() + vbCrlf();
        c= c + "[sTailPageFocus]" + vbCrlf() + sTailPageFocus + "[/sTailPageFocus]" + vbCrlf() + vbCrlf();
        c= c + "[sPageEnd]" + vbCrlf() + sPageEnd + "[/sPageEnd]" + vbCrlf();
        rwEnd("[page]" + vbCrlf() + vbCrlf() + c + vbCrlf() + "[/page]");
    }
    //配置页为空则
    if( configPageUrl== "" ){
        configPageUrl= getUrlAddToParam(getUrl(), "?page=[id]", "replace");
    }

    nDisplay= 6; //显示数
    nDispalyOK= 0; //显示成功数
    s= handleNumberType(sPage);
    if( s== "" ){
        nPage= 1;
    }else{
        nPage= cInt(s);
    }
    //获得总页数
    nCountPage= getCountPage(nRecrodCount, nPageSize);



    nPreviousPage= nPage - 1;
    nNextPage= nPage + 1;

    //处理上一页
    if( nPreviousPage <= 0 ){
        nPreviousPage= -1;
    }
    //处理下一页
    if( nNextPage > nCountPage ){
        nNextPage= -1;
    }

    //页开始
    c= sPageStart;
    //首页
    if( nPage > 1 ){
        c= c + replace(sHomePage, "[$url$]", replace(configPageUrl, "[id]", ""));
    }else if( isDisplayTip== true ){
        c= c + sHomePageFocus;
    }
    //上一页
    if( nPreviousPage != -1 ){
        sTemp= cStr(nPreviousPage);
        if( nPreviousPage <= 1 ){
            sTemp= "";
        }
        c= c + replace(sUpPage, "[$url$]", replace(configPageUrl, "[id]", sTemp));
    }else if( isDisplayTip== true ){
        c= c + sUpPageFocus;
    }


    int n=0;
    if( nPage==0 ){
        nPage=1;
    }
    n=(nPage - 3);
    //call echo("n=" & n, "nPage=" & nPage)

    //翻页循环
    for( i= n ; i<= nCountPage; i++){
        if( i >= 1 ){
            nDispalyOK= nDispalyOK + 1;
            //call echo(i,nPage)
            if( i== nPage ){
                c= c + replace(sForPageFocus, "[$i$]", i);
            }else{
                sTemp= cStr(i);
                if( i <= 1 ){
                    sTemp= "";
                }
                c= c + replace(replace(sForPage, "[$url$]", replace(configPageUrl, "[id]", sTemp)), "[$i$]", i);
            }
            if( nDispalyOK > nDisplay ){
                break;
            }
        }
    }
    //下一页
    if( nCountPage > nPage ){
        c= c + replace(sNextPage, "[$url$]", replace(configPageUrl, "[id]", nNextPage));
    }else if( isDisplayTip== true ){
        c= c + sNextPageFocus;
    }
    //末页
    if( nCountPage > nPage ){
        c= c + replace(sTailPage, "[$url$]", replace(configPageUrl, "[id]", nCountPage));
    }else if( isDisplayTip== true ){
        c= c + sTailPageFocus;
    }

    c= c + sPageEnd;


    c= replaceValueParam(c, "nRecrodCount", nRecrodCount);
    c= replaceValueParam(c, "nPage", nPage);
    if( nCountPage== 0 ){
        nCountPage= 1;
    }
    c= replaceValueParam(c, "nCountPage", nCountPage);

    if( inStr(c, "[$page-select-openlist$]") > 0 ){
        for( i= 1 ; i<= nCountPage; i++){
            url= replace(configPageUrl, "[id]", i);
            selStr= "";
            if( i== nPage ){
                selStr= " selected";
            }
            cPages= cPages + "<option value=\"" + url + "\"" + selStr + ">" + i + "</option>" + vbCrlf();
        }
        c= replace(c, "[$page-select-openlist$]", cPages);
    }

    return c + vbCrlf();
}


//获得Rs页数

</script>
