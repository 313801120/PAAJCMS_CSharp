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
DateTime pubTimer=System.DateTime.Now; 

string WEBCOLUMNTYPE="��ҳ|�ı�|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����"; 		//��վ��Ŀ�����б�
string WEB_VIEWURL="/index.aspx";                //ǰ̨��ʾ
string WEB_ADMINURL ="/admin/index.aspx";				//��̨��ʾ
string EDITORTYPE="aspx"; 		//�༭�����ͣ���ASP,��PHP,��jSP,��.NET 
string webDir="";				//��վĿ¼
string[,] ModuleReplaceArray=new string[2, 3];
bool isOnCacheHtml=false;					//�Ƿ�������html
 
string readBlockList ="";		//
string replaceDimList="";		//׷�Ӵ�ɾ��
string funDim =""; 
string rootFunDim="";
string webKeywords="";
string webDescription="";
string cfg_webtemplate=""; 
string ROOT_PATH="/";
string MDBPath = ""; 

string databaseType = ""; //���ݿ����ͣ�Access��SqlServer֮��
string sqlServerHostIP = "localhost"; //SQL����IP��ַ��������������2008�汾ʹ��
string sqlServerUsername = "sa"; //SQL���ݿ��û���
string sqlServerPassword = ""; //SQL���ݿ��û�����
string sqlServerDatabaseName = "webdata"; //SQL���ݿ���
string db_PREFIX = "asdf3a23xy_"; //��ǰ׺

//=========




string[] sArraySysColor= {"#999", "#666"};

bool makeHtmlFileToLCase= true; //����HTML�ļ�תСд
bool isWebLabelClose= true; //�պϱ�ǩ(20150831)
bool handleisCache= false; //�����Ƿ�����

bool isCacheTip= false; //�Ƿ���������ʾ
string adminDir= "/Admin/"; //��̨Ŀ¼


bool openErrorLog= true; //����������־
string openWriteSystemLog= "|txt|database|"; //����дϵͳ��־ txtд���ı� databaseд�����ݿ�
bool isTestEcho= true; //���ز��Ի���
string webVersion= "PAAJCMS v1.5"; //��վ�汾


//dim WEB_CACHEFile : WEB_CACHEFile=webDir & "/web/"& EDITORTYPE &"cachedata.txt"                                '�����ļ�
string WEB_CACHEFile= "[adminDir]/[EDITORTYPE]cachedata.txt"; //�����ļ�
string WEB_CACHEContent= ""; //�����ļ�����

string language= "#en-us"; //en-us  | zh-cn | zh-tw
string[] moduleReplaceArray= {"", ""}; //�滻����20160630







//ǰ���õ�
string code=""; string pubCode=""; //html����
string templateName=""; //ģ������
string cfg_webSiteUrl=""; string cfg_webTemplate=""; string cfg_webImages=""; string cfg_webCss=""; string cfg_webJs=""; string cfg_webTitle=""; string cfg_webKeywords=""; string cfg_webDescription=""; string cfg_webSiteBottom=""; string cfg_flags="";
string glb_columnName=""; string glb_columnId=""; string glb_id=""; string glb_columnType=""; string glb_columnENType=""; string glb_table=""; string glb_detailTitle=""; string glb_flags="";
//dim webTemplate                                                                 '��վģ��·��
string glb_url=""; string glb_filePath=""; //��ǰ������ַ,���ļ�·��
bool glb_isonhtml; //�Ƿ����ɾ�̬��ҳ
string glb_locationType=""; //λ������

string glb_bodyContent=""; //��������
string glb_articleAuthor=""; //��������
string glb_articleAdddatetime=""; //��������ʱ��
string glb_articlehits=""; //���������
string glb_upArticle=""; //��һƪ����
string glb_downArticle=""; //��һƪ����
string glb_aritcleRelatedTags=""; //���±�ǩ��
string glb_aritcleSmallImage=""; string glb_aritcleBigImage=""; //����Сͼ�����´�ͼ
string glb_searchKeyWord=""; //�����ؼ���
string cacheHtmlFilePath=""; //����html�ļ�·��
string positionEndStr=""; //�����β��׷������

bool isMakeHtml; //�Ƿ�������ҳ










//*****************************************  ����
//�����滻����
string batchReplaceL(string content, string str){
    string[] splStr; string s=""; int i=0;
    splStr= aspSplit(str, "|*|");
    for( i= 0 ; i<= uBound(splStr); i++){
        s= splStr[i];
        if( s != "" ){
            content= replaceL(content, s);
        }
    }
    return content;
}
//�滻����
string replaceL(string content, string str){
    return replace(content, str, setL(str));
}
//����
string setL(string str){
    string c="";
    c= str;
    if( language== "en-us" ){
        c= languageEN(c);
    }
    return c;
}
//����������  c=handleDisplayLanguage(c,"loginok")
string handleDisplayLanguage(string c, string sType){
    //�����ֱ��ת���ˣ���Ҫ��һ��һ��ת�ˣ�
    if( language== "zh-tw" ){
        string handleDisplayLanguage= simplifiedTransfer(c);
        return handleDisplayLanguage;
    }
    if( sType== "login" ){

        c= batchReplaceL(c, "�벻Ҫ���������ַ�|*|������ȷ|*|�û�����������ĸ|*|�û�����������ĸ|*|�����û���Ϊ��|*|�����������ĸ|*|��������Ϊ��");
        c= batchReplaceL(c, "��¼��̨|*|����Ա��¼|*|��������ǹ���Ա|*|������ֹͣ���ĵ�½��Ϊ|*|�û���|*|��");
        c= batchReplaceL(c, "��&nbsp;��|*|����|*|������|*|�� ¼|*|��¼|*|�� ��|*|����");


    }else if( sType== "loginok" ){
        c= batchReplaceL(c, "��̨��ͼ|*|�������|*|��������Ա|*|��ǰλ��|*|����Ա��Ϣ|*|�޸�����|*|���·ÿ���Ϣ|*|�����Ŷ�|*|��Ȩ����|*|������֧���Ŷ�");
        c= batchReplaceL(c, "���������޸�ģʽ|*|ϵͳ��Ϣ|*|��ѿ�Դ��|*|��Ȩ��Ϣ|*|����������|*|�������汾|*|����Ⱥ|*|�������|*|��¼��̨");
        c= batchReplaceL(c, "�û���|*|��ǰ׺|*|����|*|�˳�|*|����|*|��ҳ|*|Ȩ��|*|�˿�|*|����|*|����|*|��|*|�쳾����");
        c= batchReplaceL(c, "ϵͳ����|*|�ҵ����|*|��Ŀ|*|ģ��|*|��Ա|*|����|*|��������");

        c= batchReplaceL(c, "վ������|*|��վͳ��|*|����|*|��̨������־|*|��̨����Ա|*|��վ��Ŀ|*|������Ϣ|*|����|*|����ͳ��|*|��ҳ|*|��������|*|��Ƹ����");
        c= batchReplaceL(c, "����|*|����|*|��Ա����|*|���۴�|*|��վ����|*|��վģ��|*|��̨�˵�|*|ִ��|*|��վ");


    }
    return c;
}

//ΪӢ��
string languageEN(string str){
    string c="";
    if( str== "��¼�ɹ������ڽ����̨..." ){
        c= "Login successfully, is entering the background...";
    }else if( str== "�˺��������<br>��¼����Ϊ " ){
        c= "Account password error <br> login ";
    }else if( str== "��¼�ɹ������ڽ����̨..." ){
        c= "Login successfully, is entering the background...";
    }else if( str== "�˳��ɹ�" ){
        c= "Exit success";
    }else if( str== "�˳��ɹ������ڽ����¼����..." ){
        c= "Quit successfully, is entering the login screen...";
    }else if( str== "���������ɣ����ڽ����̨����..." ){
        c= "Clear buffer finish, is entering the background interface...";
    }else if( str== "��ʾ��Ϣ" ){
        c= "Prompt info";
    }else if( str== "������������û���Զ���ת����������" ){
        c= "If your browser does not automatically jump, please click here";
    }else if( str== "����ʱ" ){
        c= "Countdown ";
    }else if( str== "��̨��ͼ" ){
        c= "Admin map";
    }else if( str== "�������" ){
        c= "Clear buffer";
    }else if( str== "��������Ա" ){
        c= "Super administrator";
    }else if( str== "��ǰλ��" ){
        c= "current location";
    }else if( str== "����Ա��Ϣ" ){
        c= "Admin info";
    }else if( str== "�޸�����" ){
        c= "Modify password";
    }else if( str== "�û���" ){
        c= "username";
    }else if( str== "��ǰ׺" ){
        c= "Table Prefix";
    }else if( str== "���������޸�ģʽ" ){
        c= "online modification";
    }else if( str== "ϵͳ��Ϣ" ){
        c= "system info";
    }else if( str== "��Ȩ��Ϣ" ){
        c= "Authorization information";
    }else if( str== "��ѿ�Դ��" ){
        c= "Free open source";
    }else if( str== "����������" ){
        c= "Server name";
    }else if( str== "�������汾" ){
        c= "Server version";
    }else if( str== "���·ÿ���Ϣ" ){
        c= "visitor info";
    }else if( str== "�����Ŷ�" ){
        c= "team info";
    }else if( str== "��Ȩ����" ){
        c= "copyright";
    }else if( str== "������֧���Ŷ�" ){
        c= "Develop and support team";
    }else if( str== "����Ⱥ" ){
        c= "Exchange group";
    }else if( str== "�������" ){
        c= "Related links";
    }else if( str== "ϵͳ����" ){
        c= "System";
    }else if( str== "�ҵ����" ){
        c= "My panel";
    }else if( str== "��Ŀ" ){
        c= "Column";
    }else if( str== "ģ��" ){
        c= "Template";
    }else if( str== "��Ա" ){
        c= "Member";
    }else if( str== "����" ){
        c= "Generation";
    }else if( str== "��������" ){
        c= "More settings";


    }else if( str== "��¼��̨" ){
        c= "Login management background";
    }else if( str== "����Ա��¼" ){
        c= "Administrator login ";
    }else if( str== "��������ǹ���Ա" ){
        c= "If you are not an administrator";
    }else if( str== "������ֹͣ���ĵ�½��Ϊ" ){
        c= "Please stop your login immediately";
    }else if( str== "��&nbsp;��" || str== "����" ){
        c= "password";
    }else if( str== "������" ){
        c= "Please input";
    }else if( str== "�� ¼" || str== "��¼" ){
        c= "login";
    }else if( str== "�� ��" || str== "����" ){
        c= "reset";
    }else if( str== "�벻Ҫ���������ַ�" ){
        c= "Please do not enter special characters";
    }else if( str== "������ȷ" ){
        c= "Input correct";
    }else if( str== "�û�����������ĸ" ){
        c= "Username use ";
    }else if( str== "�����û���Ϊ��" ){
        c= "Your username is empty";
    }else if( str== "�����������ĸ" ){
        c= "Passwords use ";
    }else if( str== "��������Ϊ��" ){
        c= "Your password is empty";
    }else if( str== "վ������" ){
        c= "Site configuration";
    }else if( str== "��վͳ��" ){
        c= "Website statistics";
    }else if( str== "��̨������־" ){
        c= "Admin log ";
    }else if( str== "��̨����Ա" ){
        c= "Background manager";
    }else if( str== "��վ��Ŀ" ){
        c= "Website column";
    }else if( str== "������Ϣ" ){
        c= "Classified information";
    }else if( str== "����ͳ��" ){
        c= "Search statistics";
    }else if( str== "��ҳ" ){
        c= "Single page";
    }else if( str== "��������" ){
        c= "Friendship link";
    }else if( str== "��Ƹ����" ){
        c= "Recruitment management";
    }else if( str== "����" ){
        c= "Feedback management";
    }else if( str== "����" ){
        c= "message management";
    }else if( str== "��Ա����" ){
        c= "Member allocation";
    }else if( str== "���۴�" ){
        c= "Bidding words";
    }else if( str== "��վ����" ){ 
        c= "Website layout";
    }else if( str== "��վģ��" ){
        c= "Website module";
    }else if( str== "��̨�˵�" ){
        c= "Background menu";
    }else if( str== "��վ" ){
        c= "Template website ";

    }else if( str== "11111" ){
        c= "1111111";



    }else if( str== "ִ��" ){
        c= "implement ";
    }else if( str== "����" ){
        c= "comment ";
    }else if( str== "����" ){
        c= "generate ";
    }else if( str== "Ȩ��" ){
        c= "jurisdiction";
    }else if( str== "����" ){
        c= "Help";
    }else if( str== "�˳�" ){
        c= "sign out";
    }else if( str== "����" ){
        c= "hello";
    }else if( str== "��ҳ" ){
        c= "home";
    }else if( str== "�˿�" ){
        c= "port";
    }else if( str== "����" ){
        c= "official website";
    }else if( str== "����" ){
        c= "Emai";
    }else if( str== "�쳾����" ){
        c= "Cloud";

    }else if( str== "��" ){
        c= " edition";

    }else{
        c= str;
    }
    return c;
}
</script>
