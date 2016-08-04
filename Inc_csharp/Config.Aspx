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
DateTime pubTimer=System.DateTime.Now; 

string WEBCOLUMNTYPE="首页|文本|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单"; 		//网站栏目类型列表
string WEB_VIEWURL="/index.aspx";                //前台显示
string WEB_ADMINURL ="/admin/index.aspx";				//后台显示
string EDITORTYPE="aspx"; 		//编辑器类型，是ASP,或PHP,或jSP,或.NET 
string webDir="";				//网站目录
string[,] ModuleReplaceArray=new string[2, 3];
bool isOnCacheHtml=false;					//是否开启缓冲html
 
string readBlockList ="";		//
string replaceDimList="";		//追加待删除
string funDim =""; 
string rootFunDim="";
string webKeywords="";
string webDescription="";
string cfg_webtemplate=""; 
string ROOT_PATH="/";
string MDBPath = ""; 

string databaseType = ""; //数据库类型，Access与SqlServer之分
string sqlServerHostIP = "localhost"; //SQL主机IP地址（服务器名）在2008版本使用
string sqlServerUsername = "sa"; //SQL数据库用户名
string sqlServerPassword = ""; //SQL数据库用户密码
string sqlServerDatabaseName = "webdata"; //SQL数据库名
string db_PREFIX = "asdf3a23xy_"; //表前缀

//=========




string[] sArraySysColor= {"#999", "#666"};

bool makeHtmlFileToLCase= true; //生成HTML文件转小写
bool isWebLabelClose= true; //闭合标签(20150831)
bool handleisCache= false; //缓冲是否处理了

bool isCacheTip= false; //是否开启缓冲提示
string adminDir= "/Admin/"; //后台目录


bool openErrorLog= true; //开启错误日志
string openWriteSystemLog= "|txt|database|"; //开启写系统日志 txt写入文本 database写入数据库
bool isTestEcho= true; //开关测试回显
string webVersion= "PAAJCMS v1.5"; //网站版本


//dim WEB_CACHEFile : WEB_CACHEFile=webDir & "/web/"& EDITORTYPE &"cachedata.txt"                                '缓冲文件
string WEB_CACHEFile= "[adminDir]/[EDITORTYPE]cachedata.txt"; //缓冲文件
string WEB_CACHEContent= ""; //缓冲文件内容

string language= "#en-us"; //en-us  | zh-cn | zh-tw
string[] moduleReplaceArray= {"", ""}; //替换数组20160630







//前端用到
string code=""; string pubCode=""; //html代码
string templateName=""; //模板名称
string cfg_webSiteUrl=""; string cfg_webTemplate=""; string cfg_webImages=""; string cfg_webCss=""; string cfg_webJs=""; string cfg_webTitle=""; string cfg_webKeywords=""; string cfg_webDescription=""; string cfg_webSiteBottom=""; string cfg_flags="";
string glb_columnName=""; string glb_columnId=""; string glb_id=""; string glb_columnType=""; string glb_columnENType=""; string glb_table=""; string glb_detailTitle=""; string glb_flags="";
//dim webTemplate                                                                 '网站模板路径
string glb_url=""; string glb_filePath=""; //当前链接网址,和文件路径
bool glb_isonhtml; //是否生成静态网页
string glb_locationType=""; //位置类型

string glb_bodyContent=""; //主体内容
string glb_articleAuthor=""; //文章作者
string glb_articleAdddatetime=""; //文章添加时间
string glb_articlehits=""; //文章浏览器
string glb_upArticle=""; //上一篇文章
string glb_downArticle=""; //下一篇文章
string glb_aritcleRelatedTags=""; //文章标签组
string glb_aritcleSmallImage=""; string glb_aritcleBigImage=""; //文章小图与文章大图
string glb_searchKeyWord=""; //搜索关键词
string cacheHtmlFilePath=""; //缓冲html文件路径
string positionEndStr=""; //面包线尾部追加内容

bool isMakeHtml; //是否生成网页










//*****************************************  语言
//批量替换语言
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
//替换语言
string replaceL(string content, string str){
    return replace(content, str, setL(str));
}
//语言
string setL(string str){
    string c="";
    c= str;
    if( language== "en-us" ){
        c= languageEN(c);
    }
    return c;
}
//处理显语言  c=handleDisplayLanguage(c,"loginok")
string handleDisplayLanguage(string c, string sType){
    //繁体就直接转换了，不要再一个一个转了，
    if( language== "zh-tw" ){
        string handleDisplayLanguage= simplifiedTransfer(c);
        return handleDisplayLanguage;
    }
    if( sType== "login" ){

        c= batchReplaceL(c, "请不要输入特殊字符|*|输入正确|*|用户名可以用字母|*|用户名可以用字母|*|您的用户名为空|*|密码可以用字母|*|您的密码为空");
        c= batchReplaceL(c, "登录后台|*|管理员登录|*|如果您不是管理员|*|请立即停止您的登陆行为|*|用户名|*|版");
        c= batchReplaceL(c, "密&nbsp;码|*|密码|*|请输入|*|登 录|*|登录|*|重 置|*|重置");


    }else if( sType== "loginok" ){
        c= batchReplaceL(c, "后台地图|*|清除缓冲|*|超级管理员|*|当前位置|*|管理员信息|*|修改密码|*|最新访客信息|*|开发团队|*|版权所有|*|开发与支持团队");
        c= batchReplaceL(c, "进入在线修改模式|*|系统信息|*|免费开源版|*|授权信息|*|服务器名称|*|服务器版本|*|交流群|*|相关链接|*|登录后台");
        c= batchReplaceL(c, "用户名|*|表前缀|*|帮助|*|退出|*|您好|*|首页|*|权限|*|端口|*|邮箱|*|官网|*|版|*|红尘云孙");
        c= batchReplaceL(c, "系统管理|*|我的面板|*|栏目|*|模板|*|会员|*|生成|*|更多设置");

        c= batchReplaceL(c, "站点配置|*|网站统计|*|生成|*|后台操作日志|*|后台管理员|*|网站栏目|*|分类信息|*|评论|*|搜索统计|*|单页|*|友情链接|*|招聘管理");
        c= batchReplaceL(c, "反馈|*|留言|*|会员配置|*|竞价词|*|网站布局|*|网站模块|*|后台菜单|*|执行|*|仿站");


    }
    return c;
}

//为英文
string languageEN(string str){
    string c="";
    if( str== "登录成功，正在进入后台..." ){
        c= "Login successfully, is entering the background...";
    }else if( str== "账号密码错误<br>登录次数为 " ){
        c= "Account password error <br> login ";
    }else if( str== "登录成功，正在进入后台..." ){
        c= "Login successfully, is entering the background...";
    }else if( str== "退出成功" ){
        c= "Exit success";
    }else if( str== "退出成功，正在进入登录界面..." ){
        c= "Quit successfully, is entering the login screen...";
    }else if( str== "清除缓冲完成，正在进入后台界面..." ){
        c= "Clear buffer finish, is entering the background interface...";
    }else if( str== "提示信息" ){
        c= "Prompt info";
    }else if( str== "如果您的浏览器没有自动跳转，请点击这里" ){
        c= "If your browser does not automatically jump, please click here";
    }else if( str== "倒计时" ){
        c= "Countdown ";
    }else if( str== "后台地图" ){
        c= "Admin map";
    }else if( str== "清除缓冲" ){
        c= "Clear buffer";
    }else if( str== "超级管理员" ){
        c= "Super administrator";
    }else if( str== "当前位置" ){
        c= "current location";
    }else if( str== "管理员信息" ){
        c= "Admin info";
    }else if( str== "修改密码" ){
        c= "Modify password";
    }else if( str== "用户名" ){
        c= "username";
    }else if( str== "表前缀" ){
        c= "Table Prefix";
    }else if( str== "进入在线修改模式" ){
        c= "online modification";
    }else if( str== "系统信息" ){
        c= "system info";
    }else if( str== "授权信息" ){
        c= "Authorization information";
    }else if( str== "免费开源版" ){
        c= "Free open source";
    }else if( str== "服务器名称" ){
        c= "Server name";
    }else if( str== "服务器版本" ){
        c= "Server version";
    }else if( str== "最新访客信息" ){
        c= "visitor info";
    }else if( str== "开发团队" ){
        c= "team info";
    }else if( str== "版权所有" ){
        c= "copyright";
    }else if( str== "开发与支持团队" ){
        c= "Develop and support team";
    }else if( str== "交流群" ){
        c= "Exchange group";
    }else if( str== "相关链接" ){
        c= "Related links";
    }else if( str== "系统管理" ){
        c= "System";
    }else if( str== "我的面板" ){
        c= "My panel";
    }else if( str== "栏目" ){
        c= "Column";
    }else if( str== "模板" ){
        c= "Template";
    }else if( str== "会员" ){
        c= "Member";
    }else if( str== "生成" ){
        c= "Generation";
    }else if( str== "更多设置" ){
        c= "More settings";


    }else if( str== "登录后台" ){
        c= "Login management background";
    }else if( str== "管理员登录" ){
        c= "Administrator login ";
    }else if( str== "如果您不是管理员" ){
        c= "If you are not an administrator";
    }else if( str== "请立即停止您的登陆行为" ){
        c= "Please stop your login immediately";
    }else if( str== "密&nbsp;码" || str== "密码" ){
        c= "password";
    }else if( str== "请输入" ){
        c= "Please input";
    }else if( str== "登 录" || str== "登录" ){
        c= "login";
    }else if( str== "重 置" || str== "重置" ){
        c= "reset";
    }else if( str== "请不要输入特殊字符" ){
        c= "Please do not enter special characters";
    }else if( str== "输入正确" ){
        c= "Input correct";
    }else if( str== "用户名可以用字母" ){
        c= "Username use ";
    }else if( str== "您的用户名为空" ){
        c= "Your username is empty";
    }else if( str== "密码可以用字母" ){
        c= "Passwords use ";
    }else if( str== "您的密码为空" ){
        c= "Your password is empty";
    }else if( str== "站点配置" ){
        c= "Site configuration";
    }else if( str== "网站统计" ){
        c= "Website statistics";
    }else if( str== "后台操作日志" ){
        c= "Admin log ";
    }else if( str== "后台管理员" ){
        c= "Background manager";
    }else if( str== "网站栏目" ){
        c= "Website column";
    }else if( str== "分类信息" ){
        c= "Classified information";
    }else if( str== "搜索统计" ){
        c= "Search statistics";
    }else if( str== "单页" ){
        c= "Single page";
    }else if( str== "友情链接" ){
        c= "Friendship link";
    }else if( str== "招聘管理" ){
        c= "Recruitment management";
    }else if( str== "反馈" ){
        c= "Feedback management";
    }else if( str== "留言" ){
        c= "message management";
    }else if( str== "会员配置" ){
        c= "Member allocation";
    }else if( str== "竞价词" ){
        c= "Bidding words";
    }else if( str== "网站布局" ){ 
        c= "Website layout";
    }else if( str== "网站模块" ){
        c= "Website module";
    }else if( str== "后台菜单" ){
        c= "Background menu";
    }else if( str== "仿站" ){
        c= "Template website ";

    }else if( str== "11111" ){
        c= "1111111";



    }else if( str== "执行" ){
        c= "implement ";
    }else if( str== "评论" ){
        c= "comment ";
    }else if( str== "生成" ){
        c= "generate ";
    }else if( str== "权限" ){
        c= "jurisdiction";
    }else if( str== "帮助" ){
        c= "Help";
    }else if( str== "退出" ){
        c= "sign out";
    }else if( str== "您好" ){
        c= "hello";
    }else if( str== "首页" ){
        c= "home";
    }else if( str== "端口" ){
        c= "port";
    }else if( str== "官网" ){
        c= "official website";
    }else if( str== "邮箱" ){
        c= "Emai";
    }else if( str== "红尘云孙" ){
        c= "Cloud";

    }else if( str== "版" ){
        c= " edition";

    }else{
        c= str;
    }
    return c;
}
</script>

