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
//dim nCount,nPageSize,maxpage,page,x,i,PageControl


//ҳ����  ��¼����  ÿҳ��ʾ��  ��ǰ�� (2015117)   webPageControl(59,12,1,"http://www.baidu.com")
string webPageControl(int nRecrodCount, int nPageSize, string sPage, string configPageUrl, string action){
    string c=""; int nCountPage=0; int i=0; int nDisplay=0; int nDispalyOK=0; string sTemp=""; string cPages=""; string url=""; string selStr="";int nPage=0;string s="";
    int nPreviousPage=0; int nNextPage=0; //������һҳ����һҳ
    bool isDisplayTip; //�Ƿ���ʾ��ʾ��ҳ��Ϣ
    isDisplayTip= true;

    string sPageStart=""; string sPageEnd=""; string sHomePage=""; string sHomePageFocus=""; string sUpPage=""; string sUpPageFocus=""; string sNextPage=""; string sNextPageFocus=""; string sForPage=""; string sForPageFocus=""; string sTailPage=""; string sTailPageFocus="";
    if( action != "" ){
        sPageStart= getStrCut(action, "[sPageStart]", "[/sPageStart]", 2); //ҳͷ����
        sPageEnd= getStrCut(action, "[sPageEnd]", "[/sPageEnd]", 2); //ҳβ����
        sHomePage= getStrCut(action, "[sHomePage]", "[/sHomePage]", 2); //��ҳ
        sHomePageFocus= getStrCut(action, "[sHomePageFocus]", "[/sHomePageFocus]", 2); //��ҳ����
        sUpPage= getStrCut(action, "[sUpPage]", "[/sUpPage]", 2); //��һҳ
        sUpPageFocus= getStrCut(action, "[sUpPageFocus]", "[/sUpPageFocus]", 2); //��һҳ����
        sNextPage= getStrCut(action, "[sNextPage]", "[/sNextPage]", 2); //��һҳ
        sNextPageFocus= getStrCut(action, "[sNextPageFocus]", "[/sNextPageFocus]", 2); //��һҳ����
        sForPage= getStrCut(action, "[sForPage]", "[/sForPage]", 2); //ѭ��ҳ
        sForPageFocus= getStrCut(action, "[sForPageFocus]", "[/sForPageFocus]", 2); //ѭ��ҳ����
        sTailPage= getStrCut(action, "[sTailPage]", "[/sTailPage]", 2); //���ҳ
        sTailPageFocus= getStrCut(action, "[sTailPageFocus]", "[/sTailPageFocus]", 2); //���ҳ����


    }
    //ҳͷ����
    if( sPageStart== "" ){
        sPageStart= "<ul class=\"pagecontrolwrap\">" + vbCrlf() + "<li class=\"pageinfo\">��[$nRecrodCount$]�� [$nPage$]/[$nCountPage$]ҳ</li>" + vbCrlf();
    }
    //ҳβ����
    if( sPageEnd== "" ){
        sPageEnd= "</ul><div class=\"clear\"></div>" + vbCrlf();
    }
    //��ҳ
    if( sHomePage== "" ){
        sHomePage= "<li><a href=\"[$url$]\">��ҳ</a></li>" + vbCrlf();
    }
    //��ҳ����
    if( sHomePageFocus== "" ){
        sHomePageFocus= "<li class=\"pageli\">��ҳ</li>" + vbCrlf();
    }
    //��һҳ
    if( sUpPage== "" ){
        sUpPage= "<li><a href=\"[$url$]\">��һҳ</a></li>" + vbCrlf();
    }
    //��һҳ����
    if( sUpPageFocus== "" ){
        sUpPageFocus= "<li class=\"pageli\">��һҳ</li>" + vbCrlf();
    }
    //��һҳ
    if( sNextPage== "" ){
        sNextPage= "<li><a href=\"[$url$]\">��һҳ</a></li>" + vbCrlf();
    }
    //��һҳ����
    if( sNextPageFocus== "" ){
        sNextPageFocus= "<li class=\"pageli\">��һҳ</li>" + vbCrlf();
    }
    //ѭ��ҳ
    if( sForPage== "" ){
        sForPage= "<li><a href=\"[$url$]\">[$i$]</a></li>" + vbCrlf();
    }
    //ѭ��ҳ����
    if( sForPageFocus== "" ){
        sForPageFocus= "<li class=\"pagefocus\">[$i$]</li>" + vbCrlf();
    }
    //���ҳ
    if( sTailPage== "" ){
        sTailPage= "<li><a href=\"[$url$]\">ĩҳ</a></li>" + vbCrlf();
    }
    //���ҳ����
    if( sTailPageFocus== "" ){
        sTailPageFocus= "<li class=\"pageli\">ĩҳ</li>" + vbCrlf();
    }
    //����ʱ�õ�20160630
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
    //����ҳΪ����
    if( configPageUrl== "" ){
        configPageUrl= getUrlAddToParam(getUrl(), "?page=[id]", "replace");
    }

    nDisplay= 6; //��ʾ��
    nDispalyOK= 0; //��ʾ�ɹ���
    s= handleNumberType(sPage);
    if( s== "" ){
        nPage= 1;
    }else{
        nPage= cInt(s);
    }
    //�����ҳ��
    nCountPage= getCountPage(nRecrodCount, nPageSize);



    nPreviousPage= nPage - 1;
    nNextPage= nPage + 1;

    //������һҳ
    if( nPreviousPage <= 0 ){
        nPreviousPage= -1;
    }
    //������һҳ
    if( nNextPage > nCountPage ){
        nNextPage= -1;
    }

    //ҳ��ʼ
    c= sPageStart;
    //��ҳ
    if( nPage > 1 ){
        c= c + replace(sHomePage, "[$url$]", replace(configPageUrl, "[id]", ""));
    }else if( isDisplayTip== true ){
        c= c + sHomePageFocus;
    }
    //��һҳ
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

    //��ҳѭ��
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
    //��һҳ
    if( nCountPage > nPage ){
        c= c + replace(sNextPage, "[$url$]", replace(configPageUrl, "[id]", nNextPage));
    }else if( isDisplayTip== true ){
        c= c + sNextPageFocus;
    }
    //ĩҳ
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


//���Rsҳ��

</script>
