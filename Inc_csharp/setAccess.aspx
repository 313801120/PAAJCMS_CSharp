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
//��#stop#��1     ����Ϊ�����뵱ǰ������

//����callfile_setAccess�ļ�����
string callfile_setAccess(){
    switch ( cStr(Request["stype"]) ){
        case "backupDatabase" : backupDatabase() ;break;//�������ݿ�
        case "recoveryDatabase" : recoveryDatabase(); //�ָ����ݿ�

        break;
        default : eerr("setAccessҳ��û�ж���", cStr(Request["stype"]));
        break;
    }
    return "";
}

//�ָ����ݿ�
string recoveryDatabase(){
    string backupDir=""; string backupFilePath="";
    string content=""; string s=""; string[] splStr; string tableName="";
    handlePower("�ָ����ݿ�");
    backupDir= adminDir + "/Data/BackUpDateBases/";
    backupFilePath= backupDir + "/" + cStr(Request["databaseName"]);
    if( checkFile(backupFilePath)== false ){
        eerr("���ݿ��ļ�������", backupFilePath);
    }
    content= getFText(backupFilePath);
    splStr= aspSplit(content, "===============================" + vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        tableName= newGetStrCut(s, "table");
        if( tableName != "" ){
            connexecute("delete from " + db_PREFIX + tableName);
            echo(tableName, nImportTXTData(s, tableName, "���"));
        }
    }
    echo("�ָ����ݿ����", "");
    return "";
}

//�������ݿ�
string backupDatabase(){
    bool isUnifyToFile; string tableNameList=""; string databaseTableNameList=""; string fieldConfig=""; string fieldName=""; string fieldType=""; string[] splField; string fieldValue=""; int nLen=0; bool isOK;
    string[] splStr; string[] splxx; string tableName=""; string s=""; string c=""; string backupDir=""; string backupFilePath="";
    handlePower("�������ݿ�");
    tableNameList= lCase(cStr(Request["tableNameList"])); //�Զ��屸�����ݱ��б�
    isUnifyToFile= IIF(cStr(Request["isUnifyToFile"])== "1", true, false); //ͳһ�ŵ�һ���ļ���
    databaseTableNameList= lCase(db_PREFIX + "webcolumn" + vbCrlf() + getTableList()); //��db_PREFIX����ǰ�棬��Ϊ��������Ҫ�������ȡ
    nLen= len(db_PREFIX);

    //�����Զ�����б�
    if( tableNameList != "" ){
        splStr= aspSplit(tableNameList, "|");
        foreach(var eachtableName in splStr){
            tableName=eachtableName;
            if( inStr(vbCrlf() + databaseTableNameList + vbCrlf(), vbCrlf() + db_PREFIX + tableName + vbCrlf()) > 0 ){
                if( c != "" ){
                    c= c + vbCrlf();
                }
                c= c + db_PREFIX + tableName;
            }
        }
        if( c== "" ){
            eerr("�Զ��屸�ݱ���ȷ <a href=\"javascript:history.go(-1)\">�������</a>", tableNameList);
        }
        databaseTableNameList= c;
    }
    splStr= aspSplit(databaseTableNameList, vbCrlf());
    c= "";
    foreach(var eachtableName in splStr){
        tableName=eachtableName;
        tableName= aspTrim(tableName);
        isOK= true;
        //�ж�ǰ׺�Ƿ�һ��
        if( nLen > 0 ){
            if( mid(tableName, 1, nLen) != db_PREFIX ){
                isOK= false;
            }
        }
        if( isOK== true ){
            fieldConfig= lCase(getFieldConfigList(tableName));
            echo(tableName, fieldConfig);
            rs = new OleDbCommand("select * from " + tableName, conn).ExecuteReader();

            c= c + "��table��" + mid(tableName, len(db_PREFIX) + 1,-1) + vbCrlf();
            while( rs.Read()){
                splField= aspSplit(fieldConfig, ",");
                foreach(var eachs in splField){
                    s=eachs;
                    if( inStr(s, "|") > 0 ){
                        splxx= aspSplit(s, "|");
                        fieldName= splxx[0];
                        fieldType= splxx[1];
                        fieldValue= cStr(rs[fieldName]);
                        if( fieldType== "numb" ){
                            fieldValue= replace(replace(fieldValue, "True", "1"), "False", "0");
                        }
                        //��̨�˵�
                        if( tableName== db_PREFIX + "listmenu" && fieldName== "parentid" ){
                            fieldValue= getListMenuName(fieldValue);
                            //��վ��Ŀ
                        }else if( tableName== db_PREFIX + "webcolumn" && fieldName== "parentid" ){
                            fieldValue= getColumnName(fieldValue);
                        }
                        if( fieldValue != "" ){
                            if( inStr(fieldValue, vbCrlf()) > 0 ){
                                fieldValue= fieldValue + "��/" + fieldName + "��";
                            }
                            c= c + "��" + fieldName + "��" + fieldValue + vbCrlf();
                        }
                    }
                }
                c= c + "-------------------------------" + vbCrlf();
            }
            c= c + "===============================" + vbCrlf();
        }
    }
    backupDir= adminDir + "/Data/BackUpDateBases/";
    backupFilePath= backupDir + "/" + format_Time(now(), 4) + ".txt";
    createDirFolder(backupDir);
    deleteFile(backupFilePath); //ɾ���ɱ����ļ�
    createFile(backupFilePath, c); //���������ļ�
    hr();
    echo("backupDir", backupDir);
    echo("backupFilePath", backupFilePath);
    eerr("�������", "<a href='?act=displayLayout&templateFile=layout_manageDatabases.html&lableTitle=���ݿ�'>������� ���ݻָ�����</a>");
    return "";
}

//�������ݿ�����
void resetAccessData(){
    handlePower("�ָ�ģ������"); //����Ȩ�޴���
    conn=openConn();
    conn.Open();

    string[] splStr; int i=0; string s=""; string columnname=""; string title=""; int nCount=0; string webdataDir="";
    webdataDir= cStr(Request["webdataDir"]);
    if( webdataDir != "" ){
        if( checkFolder(webdataDir)== false ){
            eerr("��վ����Ŀ¼�����ڣ��ָ�Ĭ������δ�ɹ�", webdataDir);
        }
    }else{
        webdataDir= "/Data/WebData/";
    }

    //�޸���վ����
    nImportTXTData(getFText(webdataDir + "/website.txt"), "website", "�޸�");
    batchImportDirTXTData(webdataDir, db_PREFIX + "WebColumn" + vbCrlf() + getTableList()); //��webcolumn����Ϊwebcolumn�����µ������ݣ���Ϊ��̨��������Ҫ������20160711

    echo("��ʾ", "�ָ��������");
    rw("<hr><a href='../" + EDITORTYPE + "web." + EDITORTYPE + "' target='_blank'>������ҳ</a> | <a href=\"?\" target='_blank'>�����̨</a>");



    writeSystemLog("", "�ָ�Ĭ������" + db_PREFIX); //ϵͳ��־
}

//����������Ӧ����Ϣ
string batchImportDirTXTData(string webdataDir, string tableNameList){
    string folderPath=""; string tableName=""; string[] splStr; string content=""; string[] splxx; string filePath=""; string fileName=""; string handleTableNameList="";
    splStr= aspSplit(tableNameList, vbCrlf());
    foreach(var eachtableName in splStr){
        tableName=eachtableName;
        if( tableName != "" ){
            if( db_PREFIX != "" ){
                tableName= mid(tableName, len(db_PREFIX) + 1,-1);
            }
            tableName= aspTrim(lCase(tableName));
            //�жϱ� ���ظ�����
            if( inStr("|" + handleTableNameList + "|", "|" + tableName + "|")== 0 ){
                handleTableNameList= handleTableNameList + tableName + "|";

                folderPath= handlePath(webdataDir + "/" + tableName);
                if( checkFolder(folderPath)== true ){
                    connexecute("delete from " + db_PREFIX + tableName); //ɾ����ǰ��ȫ������
                    echo("tableName", tableName);
                    content= getDirAllFileList(folderPath, "txt");
                    splxx= aspSplit(content, vbCrlf());
                    foreach(var eachfilePath in splxx){
                        filePath=eachfilePath;
                        fileName= getFileName(filePath);
                        if( filePath != "" && inStr("_#", left(fileName, 1))== 0 ){
                            echo(tableName, filePath);
                            nImportTXTData(getFText(filePath), tableName, "���");
                            doEvents();
                        }
                    }
                }
            }
        }
    }
    return "";
}

//��������
int nImportTXTData(string content, string tableName, string sType){
    string fieldConfigList=""; string[] splList; string listStr=""; string[] splStr; string[] splxx; string s=""; string sql=""; int nOK=0;
    string fieldName=""; string fieldType=""; string fieldValue=""; string addFieldList=""; string addValueList=""; string updateValueList="";
    string fieldStr="";
    tableName= aspTrim(lCase(tableName)); //��
    //��������Ϊ�˴�GitHub����ʱ����vbcrlfת�� chr(10)  20160409
    if( inStr(content, vbCrlf())== 0 ){
        content= replace(content, chr(10).ToString(), vbCrlf());
    }
    fieldConfigList= lCase(getFieldConfigList(db_PREFIX + tableName));
    splStr= aspSplit(fieldConfigList, ",");
    splList= aspSplit(content, vbCrlf() + "-------------------------------");
    nOK= 0;
    foreach(var eachlistStr in splList){
        listStr=eachlistStr;
        addFieldList= ""; //����ֶ��б����
        addValueList= ""; //����ֶ��б�ֵ
        updateValueList= ""; //�޸��ֶ��б�

        s= lCase(newGetStrCut(listStr, "#stop#"));
        if( s != "1" && s != "true" ){
            foreach(var eachfieldStr in splStr){
                fieldStr=eachfieldStr;
                if( fieldStr != "" ){
                    splxx= aspSplit(fieldStr, "|");
                    fieldName= splxx[0];
                    fieldType= splxx[1];
                    if( inStr(listStr, "��" + fieldName + "��") > 0 ){
                        listStr= listStr + vbCrlf(); //�Ӹ�������Ϊ�������һ����������ӽ�ȥ 20160629
                        if( addFieldList != "" ){
                            addFieldList= addFieldList + ",";
                            addValueList= addValueList + ",";
                            updateValueList= updateValueList + ",";
                        }
                        addFieldList= addFieldList + fieldName;
                        //call echo(fieldName,fieldType)
                        //doevents
                        fieldValue= newGetStrCut(listStr, fieldName);
                        if( fieldType== "textarea" ){
                            fieldValue= contentTranscoding(fieldValue);
                            //�������Ϊ�����������Ͳ�Ҫ����true �� false ��  sqlserver��  20160803
                        }else if( fieldType== "yesno" || fieldType== "numb" ){
                            if( lCase(fieldValue)=="true" ){
                                fieldValue="1";
                            }else if( lCase(fieldValue)=="false" ){
                                fieldValue="0";
                            }

                        }
                        //call echo(tableName,fieldName)
                        //���´���
                        if((tableName== "articledetail" || tableName== "webcolumn") && fieldName== "parentid" ){
                            //call echo(tableName,fieldName)
                            //call echo("fieldValue",fieldValue)
                            fieldValue= getColumnId(fieldValue);
                            //call echo("fieldValue",fieldValue)
                            //��̨�˵�
                        }else if( tableName== "listmenu" && fieldName== "parentid" ){
                            fieldValue= getListMenuId(fieldValue);
                        }
                        if( fieldType== "date" && fieldValue== "" ){
                            fieldValue= aspDate();
                        }else if((fieldType== "time" || fieldType== "now") && fieldValue== "" ){
                            fieldValue= cStr(now());
                        }
                        if( fieldType != "yesno" && fieldType != "numb" ){
                            fieldValue= "'" + fieldValue + "'";
                            //Ĭ����ֵ����Ϊ0
                        }else if( fieldValue== "" ){
                            fieldValue= "0";
                        }
                        //call echo(fieldName,fieldType & "("& replace(left(fieldValue,22),"<","&lt;") &")" ) :doevents
                        addValueList= addValueList + fieldValue; //���ֵ
                        updateValueList= updateValueList + fieldName + "=" + fieldValue; //�޸�ֵ
                    }
                }
            }
            //�ֶ��б�Ϊ��
            if( addFieldList != "" ){
                if( sType== "�޸�" ){
                    sql= "update " + db_PREFIX + "" + tableName + " set " + updateValueList;
                }else{
                    sql= "insert into " + db_PREFIX + "" + tableName + " (" + addFieldList + ") values(" + addValueList + ")";
                }
                //���SQL
                if( checkSql(sql)== false ){
                    eerr("������ʾ", "<hr>sql=" + sql + "<br>");
                }
                nOK= nOK + 1;
            }else{
                nOK= nBatchImportColumnList(splStr, listStr, nOK, tableName);

            }
        }

    }
    return nOK;
}
//����������Ŀ�б� 20160716
int nBatchImportColumnList(string[] splField, string listStr, int nOK, string tableName){
    string[] splStr; string[] splxx; bool isColumn; string columnName=""; string s=""; string c=""; int nLen=0; string id=""; string[] parentIdArray=aspArray(99); string[] columntypeArray=aspArray(99); string[] flagsArray=aspArray(99); int nIndex=0; string fieldStr=""; string fieldName=""; string valueStr=""; int nCount=0;
    string fileName=""; string templatepath=""; string rowC="";
    isColumn= false;
    nCount= 0;
    listStr= replace(listStr, vbTab(), "    ");
    splStr= aspSplit(listStr, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        rowC= "";
        if( left(aspTrim(s), 1)== "#" ){
            //Ϊ#����ִ��
        }else if( s== "��#sub#��" ){
            isColumn= true;
        }else if( isColumn== true ){
            columnName= s;
            if( inStr(columnName, "��|��") > 0 ){
                columnName= mid(columnName, 1, inStr(columnName, "��|��") - 1);
            }
            columnName= aspRTrim(columnName);
            nLen= len(columnName);
            columnName= aspLTrim(columnName);
            nLen= nLen - len(columnName);
            nIndex= cInt(nLen / 4);
            if( columnName != "" ){
                parentIdArray[nIndex]= columnName;
                rowC= rowC + "��columnname��" + columnName + vbCrlf();
                foreach(var eachfieldStr in splField){
                    fieldStr=eachfieldStr;
                    splxx= aspSplit(fieldStr + "|", "|");
                    fieldName= splxx[0];
                    if( fieldName != "" && fieldName != "columnname" && inStr(s, fieldName + "='") > 0 ){
                        valueStr= getStrCut(s, fieldName + "='", "'", 2);
                        rowC= rowC + "��" + fieldName + "��" + valueStr + vbCrlf();

                        if( fieldName== "columntype" ){
                            columntypeArray[nIndex]= valueStr;
                        }else if( fieldName== "flags" ){
                            flagsArray[nIndex]= valueStr;

                        }else if( fieldName== "templatepath" ){
                            templatepath= valueStr;
                        }else if( fieldName== "filename" ){
                            fileName= valueStr;
                            //call echo("filename",filename)

                        }
                    }
                }
                //call echo(filename,templatepath)
                if( inStr(fileName, "[thistemplate]") > 0 ){
                    rowC= vbCrlf() + "��filename��" + replace(fileName, "[thistemplate]", templatepath) + vbCrlf() + rowC;
                }
                if( nIndex != 0 ){
                    rowC= rowC + "��parentid��" + parentIdArray[nIndex - 1] + vbCrlf();
                    rowC= rowC + "��columntype��" + columntypeArray[nIndex - 1] + vbCrlf();
                    rowC= rowC + "��flags��" + flagsArray[nIndex - 1] + vbCrlf();
                    if( columntypeArray[nIndex]== "" ){
                        columntypeArray[nIndex]= columntypeArray[nIndex - 1];
                    }
                    if( flagsArray[nIndex]== "" ){
                        flagsArray[nIndex]= flagsArray[nIndex - 1];
                    }

                }else{
                    rowC= rowC + "��parentid��-1" + vbCrlf();
                }
                rowC= rowC + "��sortrank��" + nCount + vbCrlf();
                nCount= nCount + 1;
                c= c + rowC + "-------------------------------" + vbCrlf();
            }
        }
    }
    //call die(createfile("1.txt",c))
    //��������
    if( c != "" ){
        nImportTXTData(c, tableName, "���");
    }
    return nCount;
}

//�µĽ�ȡ�ַ�20160216
string newGetStrCut(string content, string title){
    string s="";
    //��������Ϊ�˴�GitHub����ʱ����vbcrlfת�� chr(10)  20160409
    if( inStr(content, vbCrlf())== 0 ){
        content= replace(content, chr(10).ToString(), vbCrlf());
    }
    if( inStr(content, "��/" + title + "��") > 0 ){
        s= ADSql(phpTrim(getStrCut(content, "��" + title + "��", "��/" + title + "��", 0)));
    }else{
        s= ADSql(phpTrim(getStrCut(content, "��" + title + "��", vbCrlf(), 0)));
    }
    return s;
}

//����ת��
string contentTranscoding( string content){
    content= replace(replace(replace(replace(content, "<?", "&lt;?"), "?>", "?&gt;"), "<" + "%", "&lt;%"), "?>", "%&gt;");

    string[] splStr; int i=0; string s=""; string c=""; bool isTranscoding; bool isBR;
    isTranscoding= false;
    isBR= false;
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "[&htmlת��&]") > 0 ){
            isTranscoding= true;
        }
        if( inStr(s, "[&htmlת��end&]") > 0 ){
            isTranscoding= false;
        }
        if( inStr(s, "[&ȫ������&]") > 0 ){
            isBR= true;
        }
        if( inStr(s, "[&ȫ������end&]") > 0 ){
            isBR= false;
        }

        if( isTranscoding== true ){
            s= replace(replace(s, "[&htmlת��&]", ""), "<", "&lt;");
        }else{
            s= replace(s, "[&htmlת��end&]", "");
        }
        if( isBR== true ){
            s= replace(s, "[&ȫ������&]", "");
            if( right(aspTrim(s), 8) != "������/div>" ){
                s= s + "<br>";
            }
        }else{
            s= replace(s, "[&ȫ������end&]", "");
        }
        //��ǩ��ʽ������� 20160628
        if( inStr(s, "��article_lable��") > 0 ){
            s= replace(s, "��article_lable��", "");
            s= "<div class=\"article_lable\">" + s + "</div>";
        }else if( inStr(s, "��article_blockquote��") > 0 ){
            s= replace(s, "��article_blockquote��", "");
            s= "<div class=\"article_blockquote\">" + s + "</div>";
        }


        if( c != "" ){
            c= c + vbCrlf();
        }
        c= c + s;
    }
    c= replace(replace(c, "��b��", "<b>"), "��/b��", "</b>");
    c= replace(c, "������", "<");

    c= replace(replace(c, "��strong��", "<strong>"), "��/strong��", "</strong>");
    return c;
}


</script>
