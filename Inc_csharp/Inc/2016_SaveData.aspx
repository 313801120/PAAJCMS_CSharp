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
//保存数据   ?act=saveData
string saveData(string sType){

    if( cStr(Session["yzm"])=="" ){
        eerr("提示","验证码失效");
    }

    //if instr("|"& getFormFieldList() &"|","|yzm|") then
    if( cStr(Session["yzm"])!=cStr(Request.Form["yzm"]) ){
        eerr("提示","验证码错误");
    }
    Session["yzm"]="";			//清空验证码


    //保存文章评论
    if( sType== "articlecomment" ){
        autoSavePostData("", "tablecomment", "tablename||ArticleDetail,adddatetime|now,itemid||"+ cStr(Request["itemid"]) +",adddatetime,ip||"+ getIP());
        echo("提示", "评论提交成功，等待管理员审核");

    }else if( sType== "feedback" ){
        if( cStr(Request.Form["guestname"])=="" ){
            eerr("提示","姓名为空");
        }
        autoSavePostData("", "feedback", "isthrough|numb|0,adddatetime|now,ip||"+ getIP() +",columnid||" + cStr(Request.QueryString["columnid"]));
        echo("提示", "反馈提交成功，等待管理员审核");
    }else if( sType== "guestbook" ){
        if( cStr(Request.Form["guestname"])=="" ){
            eerr("提示","姓名为空");
        }
        autoSavePostData("", "guestbook", "isthrough|numb|0,adddatetime|now,ip||"+ getIP() +",columnid||" + cStr(Request.QueryString["columnid"]));
        echo("提示", "留言提交成功，等待管理员审核");

    }else if( sType== "articledetail" ){
        autoSavePostData("", "articledetail", "title|bodycontent,adddatetime|now,ip||"+ getIP());
        echo("提示", "文章提交成功");
    }
    Response.End();
    return "";
}
//自动保存POST数据到表
string autoSavePostData(string sID, string tableName, string fieldNameList){
    string sql="";
    sql=getPostSql(sID, tableName, fieldNameList);
    //检测SQL
    if( checkSql(sql)== false ){
        errorLog("出错提示：<hr>sql=" + sql + "<br>");
        return "";
    }
    //conn.execute(sql)			'checksql这一步就已经执行了不需要再执行了20160410
    return "";
}
//获得Post发送表单处理SQL语句 20160309
string getPostSql(string sID, string tableName, string fieldNameList){
    string valueStr=""; string editValueStr=""; string sql="";
    string[] splStr; string[] splxx; string s=""; string fieldList="";
    string fieldName=""; string defaultFieldValue=""; //字段名称
    string fieldSetType=""; //字段设置类型
    string fieldValue=""; //字段值

    string systemFieldList=""; //表字段列表
    systemFieldList= getHandleFieldList(db_PREFIX + tableName, "字段配置列表");

    string postFieldList=""; //post字段列表
    string[] splPost; string fieldContent=""; string fieldConfig="";
    postFieldList= getFormFieldList();
    //以后再把下面与上面这两种处理方法事成一种看看行不行
    splPost= aspSplit(postFieldList, "|");
    foreach(var eachfieldName in splPost){
        fieldName=eachfieldName;
        fieldContent= cStr(Request.Form[fieldName]);
        if( inStr(systemFieldList, "," + fieldName + "|") > 0 && inStr("," + fieldList + ",", "," + fieldName + ",")== 0 ){
            //为自定义的
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
            fieldName= splxx[0]; //字段名称
            fieldSetType= splxx[1]; //字段设置类型
            defaultFieldValue= splxx[2]; //默认字段值
            fieldValue= ADSqlRf(fieldName); //代替上面，因为它处理了'符号
            //call echo("fieldValue",fieldValue)
            //排序密码不处理
            if( fieldValue!="#NO******NO#" ){
                //md5加密
                if( fieldSetType== "md5" ){
                    fieldValue= myMD5(fieldValue);
                }

                if( fieldSetType== "yesno" ){
                    if( fieldValue== "" ){
                        fieldValue= defaultFieldValue;
                    }
                    //不为数字类型加单引号
                }else if( fieldSetType== "numb" ){
                    if( fieldValue== "" ){
                        fieldValue= defaultFieldValue;
                    }

                }else if( fieldName== "flags" ){
                    //PHP里用法
                    if( EDITORTYPE== "php" ){



                    }else{
                        fieldValue= "|" + arrayToString(aspSplit(fieldValue, ", "), "|");
                    }


                    fieldValue= "'" + fieldValue + "'";

                    //为时间
                }else if( fieldSetType== "time" || fieldSetType== "now" ){
                    if( fieldValue== "" ){
                        fieldValue= cStr(now());	//给.net用
                    }
                    fieldValue= "'" + fieldValue + "'";
                    //为时期
                }else if( fieldSetType== "date" ){
                    if( fieldValue== "" ){
                        fieldValue= aspDate();
                    }
                    fieldValue= "'" + fieldValue + "'";

                }else{
                    fieldValue= "'" + fieldValue + "'";
                }

                fieldValue=unescape(fieldValue);			//解码20160418

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
    //自定义字段是否需要写入默认值  有的
    splStr= aspSplit(fieldNameList, ",");
    foreach(var eachs in splStr){
        s=eachs;
        if( inStr(s, "|") > 0 ){
            splxx= aspSplit(s + "|||", "|");
            fieldName= splxx[0]; //字段名称
            fieldSetType= splxx[1]; //字段设置类型
            fieldValue= splxx[2]; //默认字段值

            if( inStr(systemFieldList, "," + fieldName + "|") > 0 && inStr("," + fieldList + ",", "," + fieldName + ",")== 0 ){

                if( fieldSetType== "date" && fieldValue=="" ){
                    fieldValue= aspDate();
                }else if( (fieldSetType== "time" || fieldSetType== "now") && fieldValue=="" ){
                    fieldValue=cStr(now());			//给.net用
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
