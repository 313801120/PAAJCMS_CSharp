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
//��������   ?act=saveData
string saveData(string sType){

    if( cStr(Session["yzm"])=="" ){
        eerr("��ʾ","��֤��ʧЧ");
    }

    //if instr("|"& getFormFieldList() &"|","|yzm|") then
    if( cStr(Session["yzm"])!=cStr(Request.Form["yzm"]) ){
        eerr("��ʾ","��֤�����");
    }
    Session["yzm"]="";			//�����֤��


    //������������
    if( sType== "articlecomment" ){
        autoSavePostData("", "tablecomment", "tablename||ArticleDetail,adddatetime|now,itemid||"+ cStr(Request["itemid"]) +",adddatetime,ip||"+ getIP());
        echo("��ʾ", "�����ύ�ɹ����ȴ�����Ա���");

    }else if( sType== "feedback" ){
        if( cStr(Request.Form["guestname"])=="" ){
            eerr("��ʾ","����Ϊ��");
        }
        autoSavePostData("", "feedback", "isthrough|numb|0,adddatetime|now,ip||"+ getIP() +",columnid||" + cStr(Request.QueryString["columnid"]));
        echo("��ʾ", "�����ύ�ɹ����ȴ�����Ա���");
    }else if( sType== "guestbook" ){
        if( cStr(Request.Form["guestname"])=="" ){
            eerr("��ʾ","����Ϊ��");
        }
        autoSavePostData("", "guestbook", "isthrough|numb|0,adddatetime|now,ip||"+ getIP() +",columnid||" + cStr(Request.QueryString["columnid"]));
        echo("��ʾ", "�����ύ�ɹ����ȴ�����Ա���");

    }else if( sType== "articledetail" ){
        autoSavePostData("", "articledetail", "title|bodycontent,adddatetime|now,ip||"+ getIP());
        echo("��ʾ", "�����ύ�ɹ�");
    }
    Response.End();
    return "";
}
//�Զ�����POST���ݵ���
string autoSavePostData(string sID, string tableName, string fieldNameList){
    string sql="";
    sql=getPostSql(sID, tableName, fieldNameList);
    //���SQL
    if( checkSql(sql)== false ){
        errorLog("������ʾ��<hr>sql=" + sql + "<br>");
        return "";
    }
    //conn.execute(sql)			'checksql��һ�����Ѿ�ִ���˲���Ҫ��ִ����20160410
    return "";
}
//���Post���ͱ�����SQL��� 20160309
string getPostSql(string sID, string tableName, string fieldNameList){
    string valueStr=""; string editValueStr=""; string sql="";
    string[] splStr; string[] splxx; string s=""; string fieldList="";
    string fieldName=""; string defaultFieldValue=""; //�ֶ�����
    string fieldSetType=""; //�ֶ���������
    string fieldValue=""; //�ֶ�ֵ

    string systemFieldList=""; //���ֶ��б�
    systemFieldList= getHandleFieldList(db_PREFIX + tableName, "�ֶ������б�");

    string postFieldList=""; //post�ֶ��б�
    string[] splPost; string fieldContent=""; string fieldConfig="";
    postFieldList= getFormFieldList();
    //�Ժ��ٰ����������������ִ������³�һ�ֿ����в���
    splPost= aspSplit(postFieldList, "|");
    foreach(var eachfieldName in splPost){
        fieldName=eachfieldName;
        fieldContent= cStr(Request.Form[fieldName]);
        if( inStr(systemFieldList, "," + fieldName + "|") > 0 && inStr("," + fieldList + ",", "," + fieldName + ",")== 0 ){
            //Ϊ�Զ����
            if( inStr(fieldNameList, "," + fieldName + "|") > 0 ){
                fieldConfig= mid(fieldNameList, inStr(fieldNameList, "," + fieldName + "|") + 1,-1);
            }else{
                fieldConfig= mid(systemFieldList, inStr(systemFieldList, "," + fieldName + "|") + 1,-1);
            }
            fieldConfig= mid(fieldConfig, 1, inStr(fieldConfig, ",") - 1);
            //call echo("config",fieldConfig)
            //call echo(fieldName,fieldContent)
            //call echo("fieldConfig",fieldConfig)
            splxx= aspSplit(fieldConfig + "|||", "|");
            fieldName= splxx[0]; //�ֶ�����
            fieldSetType= splxx[1]; //�ֶ���������
            defaultFieldValue= splxx[2]; //Ĭ���ֶ�ֵ
            fieldValue= ADSqlRf(fieldName); //�������棬��Ϊ��������'����
            //call echo("fieldValue",fieldValue)
            //�������벻����
            if( fieldValue!="#NO******NO#" ){
                //md5����
                if( fieldSetType== "md5" ){
                    fieldValue= myMD5(fieldValue);
                }

                if( fieldSetType== "yesno" ){
                    if( fieldValue== "" ){
                        fieldValue= defaultFieldValue;
                    }
                    //��Ϊ�������ͼӵ�����
                }else if( fieldSetType== "numb" ){
                    if( fieldValue== "" ){
                        fieldValue= defaultFieldValue;
                    }

                }else if( fieldName== "flags" ){
                    //PHP���÷�
                    if( EDITORTYPE== "php" ){



                    }else{
                        fieldValue= "|" + arrayToString(aspSplit(fieldValue, ", "), "|");
                    }


                    fieldValue= "'" + fieldValue + "'";

                    //Ϊʱ��
                }else if( fieldSetType== "time" || fieldSetType== "now" ){
                    if( fieldValue== "" ){
                        fieldValue= cStr(now());	//��.net��
                    }
                    fieldValue= "'" + fieldValue + "'";
                    //Ϊʱ��
                }else if( fieldSetType== "date" ){
                    if( fieldValue== "" ){
                        fieldValue= aspDate();
                    }
                    fieldValue= "'" + fieldValue + "'";

                }else{
                    fieldValue= "'" + fieldValue + "'";
                }

                fieldValue=unescape(fieldValue);			//����20160418

                if( valueStr != "" ){
                    valueStr= valueStr + ",";
                    editValueStr= editValueStr + ",";
                }
                valueStr= valueStr + fieldValue;
                editValueStr= editValueStr + fieldName + "=" + fieldValue;
            }
            if( fieldList != "" ){
                fieldList= fieldList + ",";
            }
            fieldList= fieldList + fieldName;


        }
    }
    //�Զ����ֶ��Ƿ���Ҫд��Ĭ��ֵ  �е�
    splStr= aspSplit(fieldNameList, ",");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "|") > 0 ){
            splxx= aspSplit(s + "|||", "|");
            fieldName= splxx[0]; //�ֶ�����
            fieldSetType= splxx[1]; //�ֶ���������
            fieldValue= splxx[2]; //Ĭ���ֶ�ֵ

            if( inStr(systemFieldList, "," + fieldName + "|") > 0 && inStr("," + fieldList + ",", "," + fieldName + ",")== 0 ){

                if( fieldSetType== "date" && fieldValue=="" ){
                    fieldValue= aspDate();
                }else if( (fieldSetType== "time" || fieldSetType== "now") && fieldValue=="" ){
                    fieldValue=cStr(now());			//��.net��
                }
                if( fieldSetType != "yesno" && fieldSetType != "numb" ){
                    fieldValue= "'" + fieldValue + "'";
                }
                if( fieldList != "" ){
                    fieldList= fieldList + ",";
                    valueStr= valueStr + ",";
                    editValueStr= editValueStr + ",";
                }
                fieldList= fieldList + fieldName;
                valueStr= valueStr + fieldValue;
                editValueStr= editValueStr + fieldName + "=" + fieldValue;
                //call echo(fieldName,fieldSetType)
            }
        }
    }

    if( sID== "" ){
        sql= "insert into " + db_PREFIX + "" + tableName + " (" + fieldList + ",updatetime) values(" + valueStr + ",'" + now() + "')";
    }else{
        sql= "update " + db_PREFIX + "" + tableName + " set " + editValueStr + ",updatetime='" + now() + "' where id=" + sID;
    }
    return sql;
}
</script>
