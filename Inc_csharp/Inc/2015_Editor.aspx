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
//��ñ༭������
string getEditorStr(string inputName,string sType){
    string C="";
    C= "<script charset=\"utf-8\" src=\"Keditor/kindeditor.js\"></"+"script> " + vbCrlf();
    C= C + "<script charset=\"utf-8\" src=\"Keditor/lang/zh_CN.js\"></"+"script>  " + vbCrlf();
    C= C + "<script language=\"javascript\">" + vbCrlf();
    C= C + "KindEditor.ready(function(K) {" + vbCrlf();
    C= C + "	var editor1 = K.create('textarea[name=\""+ inputName +"\"]', {		//�������ı���" + vbCrlf();
    C= C + "		cssPath : 'Keditor/plugins/code/prettify.css'," + vbCrlf();
    C= C + "		uploadJson : 'Keditor/"+ sType +"/upload_json."+ sType +"'," + vbCrlf();
    C= C + "		fileManagerJson : 'Keditor/"+ sType +"/file_manager_json."+ sType +"'," + vbCrlf();
    C= C + "		allowFileManager : true," + vbCrlf();
    C= C + "		afterCreate : function() {" + vbCrlf();
    C= C + "			var self = this;" + vbCrlf();
    C= C + "			K.ctrl(document, 13, function() {" + vbCrlf();
    C= C + "				self.sync();" + vbCrlf();
    C= C + "				K('form[name=example]')[0].submit();" + vbCrlf();
    C= C + "			});" + vbCrlf();
    C= C + "			K.ctrl(self.edit.doc, 13, function() {" + vbCrlf();
    C= C + "				self.sync();" + vbCrlf();
    C= C + "				K('form[name=example]')[0].submit();" + vbCrlf();
    C= C + "			});" + vbCrlf();
    C= C + "		}" + vbCrlf();
    C= C + "	});" + vbCrlf();
    C= C + "	//prettyPrint();				//��Ϊ������������Ը���ע�͵� 2013,12,12" + vbCrlf();
    C= C + "});" + vbCrlf();
    C= C + "</"+"script>" + vbCrlf();

    return C;
}
//�ϴ��ļ����
string displayUploadDialog(string returnInputName,string sType){
    string displayUploadDialog="";
    if( sType=="asp" ){
        displayUploadDialog="<iframe style='top:2px' src='"+ adminDir +"upload_Photo.asp?PhotoUrlID=1&returnInputName="+ returnInputName +"' frameborder=0 scrolling='No' width=340 height=25></iframe>";
    }else if( sType=="php" ){
        displayUploadDialog="<iframe style='top:2px' src='upfile.php?returnInputName="+ returnInputName +"' frameborder=0 scrolling='No' width=340 height=25></iframe>";
    }
    return displayUploadDialog;
}
</script>