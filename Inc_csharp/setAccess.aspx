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
//【#stop#】1     这种为不导入当前条数据

//调用callfile_setAccess文件函数
string callfile_setAccess(){
    switch ( cStr(Request["stype"]) ){
        case "backupDatabase" : backupDatabase() ;break;//备份数据库
        case "recoveryDatabase" : recoveryDatabase(); //恢复数据库

        break;
        default : eerr("setAccess页里没有动作", cStr(Request["stype"]));
        break;
    }
    return "";
}

//恢复数据库
string recoveryDatabase(){
    string backupDir=""; string backupFilePath="";
    string content=""; string s=""; string[] splStr; string tableName="";
    handlePower("恢复数据库");
    backupDir= adminDir + "/Data/BackUpDateBases/";
    backupFilePath= backupDir + "/" + cStr(Request["databaseName"]);
    if( checkFile(backupFilePath)== false ){
        eerr("数据库文件不存在", backupFilePath);
    }
    content= getFText(backupFilePath);
    splStr= aspSplit(content, "===============================" + vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        tableName= newGetStrCut(s, "table");
        if( tableName != "" ){
            connexecute("delete from " + db_PREFIX + tableName);
            echo(tableName, nImportTXTData(s, tableName, "添加"));
        }
    }
    echo("恢复数据库完成", "");
    return "";
}

//备份数据库
string backupDatabase(){
    bool isUnifyToFile; string tableNameList=""; string databaseTableNameList=""; string fieldConfig=""; string fieldName=""; string fieldType=""; string[] splField; string fieldValue=""; int nLen=0; bool isOK;
    string[] splStr; string[] splxx; string tableName=""; string s=""; string c=""; string backupDir=""; string backupFilePath="";
    handlePower("备份数据库");
    tableNameList= lCase(cStr(Request["tableNameList"])); //自定义备份数据表列表
    isUnifyToFile= IIF(cStr(Request["isUnifyToFile"])== "1", true, false); //统一放到一个文件里
    databaseTableNameList= lCase(db_PREFIX + "webcolumn" + vbCrlf() + getTableList()); //让db_PREFIX在最前面，因为文章类型要从这里读取
    nLen= len(db_PREFIX);

    //处理自定义表列表
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
            eerr("自定义备份表不正确 <a href=\"javascript:history.go(-1)\">点击返回</a>", tableNameList);
        }
        databaseTableNameList= c;
    }
    splStr= aspSplit(databaseTableNameList, vbCrlf());
    c= "";
    foreach(var eachtableName in splStr){
        tableName=eachtableName;
        tableName= aspTrim(tableName);
        isOK= true;
        //判断前缀是否一样
        if( nLen > 0 ){
            if( mid(tableName, 1, nLen) != db_PREFIX ){
                isOK= false;
            }
        }
        if( isOK== true ){
            fieldConfig= lCase(getFieldConfigList(tableName));
            echo(tableName, fieldConfig);
            rs = new OleDbCommand("select * from " + tableName, conn).ExecuteReader();

            c= c + "【table】" + mid(tableName, len(db_PREFIX) + 1,-1) + vbCrlf();
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
                        //后台菜单
                        if( tableName== db_PREFIX + "listmenu" && fieldName== "parentid" ){
                            fieldValue= getListMenuName(fieldValue);
                            //网站栏目
                        }else if( tableName== db_PREFIX + "webcolumn" && fieldName== "parentid" ){
                            fieldValue= getColumnName(fieldValue);
                        }
                        if( fieldValue != "" ){
                            if( inStr(fieldValue, vbCrlf()) > 0 ){
                                fieldValue= fieldValue + "【/" + fieldName + "】";
                            }
                            c= c + "【" + fieldName + "】" + fieldValue + vbCrlf();
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
    deleteFile(backupFilePath); //删除旧备份文件
    createFile(backupFilePath, c); //创建备份文件
    hr();
    echo("backupDir", backupDir);
    echo("backupFilePath", backupFilePath);
    eerr("操作完成", "<a href='?act=displayLayout&templateFile=layout_manageDatabases.html&lableTitle=数据库'>点击返回 备份恢复数据</a>");
    return "";
}

//重置数据库数据
void resetAccessData(){
    handlePower("恢复模板数据"); //管理权限处理
    conn=openConn();
    conn.Open();

    string[] splStr; int i=0; string s=""; string columnname=""; string title=""; int nCount=0; string webdataDir="";
    webdataDir= cStr(Request["webdataDir"]);
    if( webdataDir != "" ){
        if( checkFolder(webdataDir)== false ){
            eerr("网站数据目录不存在，恢复默认数据未成功", webdataDir);
        }
    }else{
        webdataDir= "/Data/WebData/";
    }

    //修改网站配置
    nImportTXTData(getFText(webdataDir + "/website.txt"), "website", "修改");
    batchImportDirTXTData(webdataDir, db_PREFIX + "WebColumn" + vbCrlf() + getTableList()); //加webcolumn是因为webcolumn必需新导入数据，因为后台文章类型要从里获得20160711

    echo("提示", "恢复数据完成");
    rw("<hr><a href='../" + EDITORTYPE + "web." + EDITORTYPE + "' target='_blank'>进入首页</a> | <a href=\"?\" target='_blank'>进入后台</a>");



    writeSystemLog("", "恢复默认数据" + db_PREFIX); //系统日志
}

//批量导入相应表信息
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
            //判断表 不重复操作
            if( inStr("|" + handleTableNameList + "|", "|" + tableName + "|")== 0 ){
                handleTableNameList= handleTableNameList + tableName + "|";

                folderPath= handlePath(webdataDir + "/" + tableName);
                if( checkFolder(folderPath)== true ){
                    connexecute("delete from " + db_PREFIX + tableName); //删除当前表全部数据
                    echo("tableName", tableName);
                    content= getDirAllFileList(folderPath, "txt");
                    splxx= aspSplit(content, vbCrlf());
                    foreach(var eachfilePath in splxx){
                        filePath=eachfilePath;
                        fileName= getFileName(filePath);
                        if( filePath != "" && inStr("_#", left(fileName, 1))== 0 ){
                            echo(tableName, filePath);
                            nImportTXTData(getFText(filePath), tableName, "添加");
                            doEvents();
                        }
                    }
                }
            }
        }
    }
    return "";
}

//导入数数
int nImportTXTData(string content, string tableName, string sType){
    string fieldConfigList=""; string[] splList; string listStr=""; string[] splStr; string[] splxx; string s=""; string sql=""; int nOK=0;
    string fieldName=""; string fieldType=""; string fieldValue=""; string addFieldList=""; string addValueList=""; string updateValueList="";
    string fieldStr="";
    tableName= aspTrim(lCase(tableName)); //表
    //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
    if( inStr(content, vbCrlf())== 0 ){
        content= replace(content, chr(10).ToString(), vbCrlf());
    }
    fieldConfigList= lCase(getFieldConfigList(db_PREFIX + tableName));
    splStr= aspSplit(fieldConfigList, ",");
    splList= aspSplit(content, vbCrlf() + "-------------------------------");
    nOK= 0;
    foreach(var eachlistStr in splList){
        listStr=eachlistStr;
        addFieldList= ""; //添加字段列表清空
        addValueList= ""; //添加字段列表值
        updateValueList= ""; //修改字段列表

        s= lCase(newGetStrCut(listStr, "#stop#"));
        if( s != "1" && s != "true" ){
            foreach(var eachfieldStr in splStr){
                fieldStr=eachfieldStr;
                if( fieldStr != "" ){
                    splxx= aspSplit(fieldStr, "|");
                    fieldName= splxx[0];
                    fieldType= splxx[1];
                    if( inStr(listStr, "【" + fieldName + "】") > 0 ){
                        listStr= listStr + vbCrlf(); //加个换行是为了让最后一个参数能添加进去 20160629
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
                            //加这个是为了让数子类型不要出现true 或 false 给  sqlserver用  20160803
                        }else if( fieldType== "yesno" || fieldType== "numb" ){
                            if( lCase(fieldValue)=="true" ){
                                fieldValue="1";
                            }else if( lCase(fieldValue)=="false" ){
                                fieldValue="0";
                            }

                        }
                        //call echo(tableName,fieldName)
                        //文章大类
                        if((tableName== "articledetail" || tableName== "webcolumn") && fieldName== "parentid" ){
                            //call echo(tableName,fieldName)
                            //call echo("fieldValue",fieldValue)
                            fieldValue= getColumnId(fieldValue);
                            //call echo("fieldValue",fieldValue)
                            //后台菜单
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
                            //默认数值类型为0
                        }else if( fieldValue== "" ){
                            fieldValue= "0";
                        }
                        //call echo(fieldName,fieldType & "("& replace(left(fieldValue,22),"<","&lt;") &")" ) :doevents
                        addValueList= addValueList + fieldValue; //添加值
                        updateValueList= updateValueList + fieldName + "=" + fieldValue; //修改值
                    }
                }
            }
            //字段列表不为空
            if( addFieldList != "" ){
                if( sType== "修改" ){
                    sql= "update " + db_PREFIX + "" + tableName + " set " + updateValueList;
                }else{
                    sql= "insert into " + db_PREFIX + "" + tableName + " (" + addFieldList + ") values(" + addValueList + ")";
                }
                //检测SQL
                if( checkSql(sql)== false ){
                    eerr("出错提示", "<hr>sql=" + sql + "<br>");
                }
                nOK= nOK + 1;
            }else{
                nOK= nBatchImportColumnList(splStr, listStr, nOK, tableName);

            }
        }

    }
    return nOK;
}
//批量导入栏目列表 20160716
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
            //为#号则不执行
        }else if( s== "【#sub#】" ){
            isColumn= true;
        }else if( isColumn== true ){
            columnName= s;
            if( inStr(columnName, "【|】") > 0 ){
                columnName= mid(columnName, 1, inStr(columnName, "【|】") - 1);
            }
            columnName= aspRTrim(columnName);
            nLen= len(columnName);
            columnName= aspLTrim(columnName);
            nLen= nLen - len(columnName);
            nIndex= cInt(nLen / 4);
            if( columnName != "" ){
                parentIdArray[nIndex]= columnName;
                rowC= rowC + "【columnname】" + columnName + vbCrlf();
                foreach(var eachfieldStr in splField){
                    fieldStr=eachfieldStr;
                    splxx= aspSplit(fieldStr + "|", "|");
                    fieldName= splxx[0];
                    if( fieldName != "" && fieldName != "columnname" && inStr(s, fieldName + "='") > 0 ){
                        valueStr= getStrCut(s, fieldName + "='", "'", 2);
                        rowC= rowC + "【" + fieldName + "】" + valueStr + vbCrlf();

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
                    rowC= vbCrlf() + "【filename】" + replace(fileName, "[thistemplate]", templatepath) + vbCrlf() + rowC;
                }
                if( nIndex != 0 ){
                    rowC= rowC + "【parentid】" + parentIdArray[nIndex - 1] + vbCrlf();
                    rowC= rowC + "【columntype】" + columntypeArray[nIndex - 1] + vbCrlf();
                    rowC= rowC + "【flags】" + flagsArray[nIndex - 1] + vbCrlf();
                    if( columntypeArray[nIndex]== "" ){
                        columntypeArray[nIndex]= columntypeArray[nIndex - 1];
                    }
                    if( flagsArray[nIndex]== "" ){
                        flagsArray[nIndex]= flagsArray[nIndex - 1];
                    }

                }else{
                    rowC= rowC + "【parentid】-1" + vbCrlf();
                }
                rowC= rowC + "【sortrank】" + nCount + vbCrlf();
                nCount= nCount + 1;
                c= c + rowC + "-------------------------------" + vbCrlf();
            }
        }
    }
    //call die(createfile("1.txt",c))
    //继续导入
    if( c != "" ){
        nImportTXTData(c, tableName, "添加");
    }
    return nCount;
}

//新的截取字符20160216
string newGetStrCut(string content, string title){
    string s="";
    //这样做是为了从GitHub下载时它把vbcrlf转成 chr(10)  20160409
    if( inStr(content, vbCrlf())== 0 ){
        content= replace(content, chr(10).ToString(), vbCrlf());
    }
    if( inStr(content, "【/" + title + "】") > 0 ){
        s= ADSql(phpTrim(getStrCut(content, "【" + title + "】", "【/" + title + "】", 0)));
    }else{
        s= ADSql(phpTrim(getStrCut(content, "【" + title + "】", vbCrlf(), 0)));
    }
    return s;
}

//内容转码
string contentTranscoding( string content){
    content= replace(replace(replace(replace(content, "<?", "&lt;?"), "?>", "?&gt;"), "<" + "%", "&lt;%"), "?>", "%&gt;");

    string[] splStr; int i=0; string s=""; string c=""; bool isTranscoding; bool isBR;
    isTranscoding= false;
    isBR= false;
    splStr= aspSplit(content, vbCrlf());
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "[&html转码&]") > 0 ){
            isTranscoding= true;
        }
        if( inStr(s, "[&html转码end&]") > 0 ){
            isTranscoding= false;
        }
        if( inStr(s, "[&全部换行&]") > 0 ){
            isBR= true;
        }
        if( inStr(s, "[&全部换行end&]") > 0 ){
            isBR= false;
        }

        if( isTranscoding== true ){
            s= replace(replace(s, "[&html转码&]", ""), "<", "&lt;");
        }else{
            s= replace(s, "[&html转码end&]", "");
        }
        if( isBR== true ){
            s= replace(s, "[&全部换行&]", "");
            if( right(aspTrim(s), 8) != "【《】/div>" ){
                s= s + "<br>";
            }
        }else{
            s= replace(s, "[&全部换行end&]", "");
        }
        //标签样式超简单添加 20160628
        if( inStr(s, "【article_lable】") > 0 ){
            s= replace(s, "【article_lable】", "");
            s= "<div class=\"article_lable\">" + s + "</div>";
        }else if( inStr(s, "【article_blockquote】") > 0 ){
            s= replace(s, "【article_blockquote】", "");
            s= "<div class=\"article_blockquote\">" + s + "</div>";
        }


        if( c != "" ){
            c= c + vbCrlf();
        }
        c= c + s;
    }
    c= replace(replace(c, "【b】", "<b>"), "【/b】", "</b>");
    c= replace(c, "【《】", "<");

    c= replace(replace(c, "【strong】", "<strong>"), "【/strong】", "</strong>");
    return c;
}


</script>
