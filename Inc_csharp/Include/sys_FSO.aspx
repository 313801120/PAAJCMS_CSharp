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
/*
创建于：20160724
属性	功能和用途
Attributes	返回和文件相关的属性值，运用了FileAttributes枚举类型值
CreationTime	返回文件的创建时间
Exists	检查文件是否存在于给定的目录中
Extension	返回文件的扩展名
LastAccessTime	返回文件的上次访问时间
FullName	返回文件的绝对路径
LastWriteTime	返回文件的上次写操作时间
Name	返回给定文件的文件名
Delete（）	删除一个文件的方法，请务必谨慎地运用该方法
*/

 

//读文件
string getFText(string filePath){
	filePath=handlePath(filePath); 
	if(File.Exists(@filePath)==false){
		return "";
	} 
	//StreamReader fso = new StreamReader( @filePath ); 	
	StreamReader fso = new StreamReader(@filePath, Encoding.GetEncoding("gb2312"));
	string st = ""; 
	while ( !fso.EndOfStream ){ 
		if(st!=""){ 
			st+="\n";
		}
		st += fso.ReadLine();
	} 
	fso.Close();
	return st;

} 
//
string reaFile(string filePath, string codeset){
	return getFText(filePath);
}
//创建文件
bool createFile(string filePath, string content){
	filePath=handlePath(filePath);
	deleteFile(filePath);			//先删除文件，要不然会出错的
	try{
		FileStream aFile = new FileStream(@filePath, FileMode.OpenOrCreate);
		StreamWriter sw = new StreamWriter(aFile, Encoding.GetEncoding("gb2312"));
		sw.Write(content);
		sw.Close();
		aFile.Close();
		return true;
	}catch(IOException ex) { 
		return false;
	}  
}
//创建gbk文件，待改进
bool createFileGBK(string filePath, string content){
	return createFile(filePath,content);
}
//创建gbk文件，待改进
bool writeToFile(string filePath, string content, string setFileCode){
	return createFile(filePath,content);
}

//创建累加文件
bool createAddFile(string filePath, string content){
	string c=getFText(filePath)+content;
	return createFile(filePath,content);
}
//检测文件
bool checkFile(string filePath){
	filePath=handlePath(filePath);
	if(File.Exists(@filePath)){
		return true;
	}else{
		return false;
	} 
}
//移动文件
bool moveFile(string filePath,string toFilePath){
	filePath=handlePath(filePath);
	toFilePath=handlePath(toFilePath);
	if(checkFile(filePath)==true && checkFile(toFilePath)==false){	
　　　　File.Move(filePath,toFilePath);
		return true;
	}else{
		return false;
	}
}
//复制文件
bool copyFile(string filePath,string toFilePath){
	filePath=handlePath(filePath);
	toFilePath=handlePath(toFilePath);
	if(checkFile(filePath)==true && checkFile(toFilePath)==false){	
　　　　File.Copy(filePath,toFilePath);
		return true;
	}else{
		return false;
	}
}
//删除文件
bool deleteFile(string filePath){
	filePath=handlePath(filePath);
	if (File.Exists(filePath)) {
		FileInfo fi = new FileInfo(filePath);
		if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1){
			fi.Attributes = FileAttributes.Normal;
		}
		File.Delete(filePath);
		return true;
	}else{
		return false;
	}
}
//获得文件大小
int getFSize(string filePath){
	filePath=handlePath(filePath); 
	//判断当前路径所指向的是否为文件
	if (File.Exists(filePath) == true){ 
		//以获取其大小
		FileInfo fileInfo = new FileInfo(filePath);
		return cInt(fileInfo.Length); 
	}else{
		return 0;
	}
}

//获得文件创建时间
string getFileCreateTime(string filePath){
	filePath=handlePath(filePath); 
	if(File.Exists(@filePath)==false){
		return "";
	}
	FileInfo fi = new FileInfo(filePath);
	return  fi.CreationTime.ToString();
}
//获得文件修改时间
string getFileEditTime(string filePath){
	filePath=handlePath(filePath); 
	if(File.Exists(@filePath)==false){
		return "";
	}
	FileInfo fi = new FileInfo(filePath);
	return  fi.LastWriteTime.ToString();
}
//获得文件访问时间
string getFileVisitTime(string filePath){
	filePath=handlePath(filePath); 
	if(File.Exists(@filePath)==false){
		return "";
	}
	FileInfo fi = new FileInfo(filePath);
	return  fi.LastAccessTime.ToString();
}



//创建文件夹
bool createFolder(string folderPath){
	folderPath=handlePath(folderPath);
	if(!Directory.Exists(@folderPath)){
		Directory.CreateDirectory(folderPath);		
		return true;
	}else{
		return false;
	} 
} 
//连续创建目录 .net里不用里createFolder就可以连续创建目录
bool createDirFolder(string folderPath){
	return createFolder ( folderPath );
}
//检测文件夹
bool checkFolder(string folderPath){
	folderPath=handlePath(folderPath);
	if(Directory.Exists(@folderPath)){
		return true;
	}else{
		return false;
	} 
}
//删除文件夹
bool deleteFolder(string folderPath){
	folderPath=handlePath(folderPath);
	if(checkFolder(folderPath)==false){
		return false;
	}
	DirectoryInfo dir = new DirectoryInfo(folderPath); 
	if (dir.Exists){
		DirectoryInfo[] childs = dir.GetDirectories();
		foreach (DirectoryInfo child in childs){
			child.Delete(true);
		}
		dir.Delete(true);
	}
	return true;
}
//移动文件夹
bool moveFolder(string folderPath, string toFolderName){
	folderPath=handlePath(folderPath);
	toFolderName=handlePath(toFolderName);
	DirectoryInfo di = new DirectoryInfo(folderPath); 
	DirectoryInfo di2 = new DirectoryInfo(toFolderName); 
	if (!di.Exists){ 
		//源文件不存在
		return false ; 
	} 
	if (di2.Exists){ 
		//目标文件夹已存在
		return false;
	}else{
		di.MoveTo(toFolderName); 
		return true;
	}
}
//复制文件夹
bool copyFolder(string folderPath, string toFolderName){
	folderPath=handlePath(folderPath);
	toFolderName=handlePath(toFolderName);
 	if (!Directory.Exists(folderPath)){
        return false;
	} 
	//创建目录文件夹
    if (!Directory.Exists(toFolderName)){
        Directory.CreateDirectory(toFolderName);
    }else{
		return false;
	}

    string[] files = Directory.GetFiles(folderPath);
    foreach (string formFileName in files)
    {
        string fileName = Path.GetFileName(formFileName);
        string toFileName = Path.Combine(toFolderName, fileName);
        File.Copy(formFileName, toFileName);
    }
    string[] fromDirs = Directory.GetDirectories(folderPath);
    foreach (string fromDirName in fromDirs) {
        string dirName = Path.GetFileName(fromDirName);
        //string toDirName = Path.Combine(toFolderName, dirName); 
		copyFolder(folderPath +"/" + dirName, toFolderName+"/"+dirName); 
		
    }
	return true;
}
//获得文件夹修改时间
string getFolderEditTime(string folderPath){
	folderPath=handlePath(folderPath); 
	if(Directory.Exists(@folderPath)==false){
		return "";
	}
	DirectoryInfo dir = new DirectoryInfo(folderPath); 
	return  dir.LastWriteTime.ToString();
}
//获得文件夹访问时间
string getFolderVisitTime(string folderPath){
	folderPath=handlePath(folderPath); 
	if(Directory.Exists(@folderPath)==false){
		return "";
	}
	DirectoryInfo dir = new DirectoryInfo(folderPath); 
	return  dir.LastAccessTime.ToString();
}
 
//获得文件夹创建时间
string getFolderCreateTime(string folderPath){
	folderPath=handlePath(folderPath); 
	if(Directory.Exists(@folderPath)==false){
		return "";
	}
	DirectoryInfo dir = new DirectoryInfo(folderPath); 
	return  dir.CreationTime.ToString();
} 


// 《获得文件文件夹》
string getFileFolderList(string folderPath, string c, string action, string fileTypeList){
	folderPath=handlePath(folderPath); 
	string fileType;string s;string folderName;string tempS;
	if(action==""){
		action="|处理文件|处理文件夹|文件名称|文件夹名称|循环文件夹|";
	}
	//Response.Write("action="+action);
	if(fileTypeList==""){
		fileTypeList="|*|";
	}
	//判断文件是否存在
	if(Directory.Exists(@folderPath)==false){
		return c;
	}
	//文件
	if(inStr("|" + action + "|", "|处理文件|")>0){
		string[] files = Directory.GetFiles(folderPath);
		foreach (var filePath in files){
			string fileName= Path.GetFileName(filePath);
			fileType="";
			if(inStr(fileName,".")>0){
				fileType=lCase(mid(fileName,inStrRev(fileName,".")+1));
			}
			if(inStr("|" + action + "|", "|文件名称|")>0){
				s=fileName;
			}else if(inStr("|" + action + "|", "|文件类型|")>0){
				s=fileType;
			}else{ 
				s=filePath;
			}
			if(inStr("|" + fileTypeList + "|", "|"+ fileType +"|")>0 || inStr("|" + fileTypeList + "|", "|*|")>0){
				c=c+s+vbCrlf();
			} 
		}
	}
	//文件夹 
    string[] fromDirs = Directory.GetDirectories(folderPath);
    foreach (string fromDirName in fromDirs) {
        folderName = Path.GetFileName(fromDirName);
		//Response.Write("文件夹="+fromDirName+"<hr>");  
		
		if(inStr("|" + action + "|", "|处理文件夹|")>0){
			if(inStr("|" + action + "|", "|文件夹名称|")>0){
				s = folderName;
			}else{
				s=folderPath;
			}
			c=c+s+vbCrlf();
		}
		if(inStr("|" + action + "|", "|循环文件夹|")>0){
			tempS=getFileFolderList(fromDirName,"",action,fileTypeList);
			if(tempS!=""){
				c=c+tempS+vbCrlf();
			}
		}
		
		  
    }
	return c;
	 
}   

//获得当前目录下全部Jpg文件
string getDirJpgList(string folderPath){
	return getDirFileList(folderPath, "jpg");
}
//获得当前目录下全部Png文件
string getDirPngList(string folderPath){
	return getDirFileList(folderPath, "png");
}
//获得当前目录下全部Ini文件
string getDirIniList(string folderPath){
	return getDirFileList(folderPath, "ini");
}
//获得当前目录下全部Txt文件
string getDirTxtList(string folderPath){
	return getDirFileList(folderPath, "txt");
}
//获得当前目录下全部Js文件
string getDirJsList(string folderPath){
	return getDirFileList(folderPath, "js");
}
//获得当前目录下全部Css文件
string getDirCssList(string folderPath){
	return getDirFileList(folderPath, "css");
}
//获得当前目录下全部Html文件
string getDirHtmlList(string folderPath){
	return getDirFileList(folderPath, "html");
}
//获得当前目录下全部asp文件
string getDirAspList(string folderPath){
	return getDirFileList(folderPath, "asp");
}
//获得当前目录下全部Php文件
string getDirPhpList(string folderPath){
	return getDirFileList(folderPath, "php");
}

//获得当前目录下批量文件列表
string getDirFileList(string folderPath,string fileTypeList){
	if(fileTypeList==""){
		fileTypeList="*";
	} 
	return getFileFolderList(folderPath,"","|处理文件|","|"+fileTypeList+"|");
} 


//获得当前目录下全部Jpg文件名称
string getDirJpgNameList(string folderPath){
	return getDirFileNameList(folderPath, "jpg");
}
//获得当前目录下全部Png文件名称
string getDirPngNameList(string folderPath){
	return getDirFileNameList(folderPath, "png");
}
//获得当前目录下全部Ini文件名称
string getDirIniNameList(string folderPath){
	return getDirFileNameList(folderPath, "ini");
}
//获得当前目录下全部Txt文件名称
string getDirTxtNameList(string folderPath){
	return getDirFileNameList(folderPath, "txt");
}
//获得当前目录下全部Js文件名称
string getDirJsNameList(string folderPath){
	return getDirFileNameList(folderPath, "js");
}
//获得当前目录下全部Css文件名称
string getDirCssNameList(string folderPath){
	return getDirFileNameList(folderPath, "css");
}
//获得当前目录下全部Html文件名称
string getDirHtmlNameList(string folderPath){
	return getDirFileNameList(folderPath, "html");
}
//获得当前目录下全部asp文件名称
string getDirAspNameList(string folderPath){
	return getDirFileNameList(folderPath, "asp");
}
//获得当前目录下全部Php文件名称
string getDirPhpNameList(string folderPath){
	return getDirFileNameList(folderPath, "php");
}
//获得当前目录下批量文件名称列表
string getDirFileNameList(string folderPath,string fileTypeList){
	if(fileTypeList==""){
		fileTypeList="*";
	}
	//aspecho("fileTypeList",fileTypeList);
	return getFileFolderList(folderPath,"","|处理文件|文件名称|","|"+fileTypeList+"|");
}

//目录文件排序 待测试
string getDirFileSort(string folderPath){
	string[] splstr;string content="";string myFile="";string c="";int i=0;int nIndex=0;string[] arrayStr=new string[33];;
	content=getDirFileNameList(folderPath,"");
	splstr=aspSplit(content,vbCrlf());
	foreach(var eachmyFile in splstr){
		myFile=eachmyFile;
		if( inStr(myFile, "#")== 0 && inStr(myFile, "、") > 0 ){ //#代表不显示，并且要有、符号
			nIndex= cInt(replace(left(myFile, 2), "、", ""));
			arrayStr[nIndex]= myFile;
		}
		doEvents();
	}
	for( i= 0 ; i<= uBound(arrayStr); i++){
		if( arrayStr[i] != "" ){
			c= c + arrayStr[i] + vbCrlf();
		}
	}
	return c;
}


 //获得当前目录下批量文件列表
string getDirAllFileList(string folderPath,string fileTypeList){
	if(fileTypeList==""){
		fileTypeList="*";
	}
	//aspecho("fileTypeList",fileTypeList);
	return getFileFolderList(folderPath,"","|处理文件|循环文件夹|hidefolderlist|","|"+fileTypeList+"|");
} 
//获得当前目录下批量文件名称列表
string getDirAllFileNameList(string folderPath,string fileTypeList){
	if(fileTypeList==""){
		fileTypeList="*";
	}
	//aspecho("fileTypeList",fileTypeList);
	return getFileFolderList(folderPath,"","|处理文件|文件名称|循环文件夹|hidefolderlist|","|"+fileTypeList+"|");
}


//获得当前目录下文件夹
string getDirFolderList(string folderPath){ 
	return getFileFolderList(folderPath,"","|处理文件夹|","");
}
//获得当前目录下文件夹
string getDirFolderNameList(string folderPath){ 
	return getFileFolderList(folderPath,"","|处理文件夹|文件夹名称|","");
}
//获得当前目录下全部文件夹
string getDirAllFolderList(string folderPath){ 
	return getFileFolderList(folderPath,"","|处理文件夹|循环文件夹|","");
}
//获得当前目录下全部文件夹
string getDirAllFolderNameList(string folderPath){ 
	return getFileFolderList(folderPath,"","|处理文件夹|循环文件夹|文件夹名称|","");
}



/////////////////////////////////////////  


//处理成完成路径 (2013,9,27
string handlePath(string fFPath){ //Path前面不加ByVal 重定义，这样是为了让前面函数里可以使用这个路径完整调用
    fFPath= replace(fFPath, "/", "\\");
    fFPath= replace(fFPath, "\\\\", "\\");
    fFPath= replace(fFPath, "\\\\", "\\");
    bool isDir; //为目录
    isDir= false;
    if( right(fFPath, 1)== "\\" ){
        isDir= true;
    }
    if( inStr(fFPath, ":")== 0 ){
        if( left(fFPath, 1)== "\\" ){
            fFPath= mapPath("\\") + "\\" + fFPath;
        }else{
            fFPath= mapPath(".\\") + "\\" + fFPath;
        }
    }
    fFPath= replace(fFPath, "/", "\\");
    fFPath= replace(fFPath, "\\\\", "\\");
    fFPath= replace(fFPath, "\\\\", "\\");
    fFPath= fullPath(fFPath);
    if( isDir== true ){
        fFPath= fFPath + "\\";
    }
    string handlePath= fFPath;
    return handlePath;
}
//完整路径
string fullPath( string fFPath){
    string[] splStr; string s=""; string c="";
    c= "";
    fFPath= replace(fFPath, "/", "\\");
    splStr= aspSplit(fFPath, "\\");
    foreach(var eachs in splStr){
        s=eachs;
        s= aspTrim(s) ;
        s=eachs;
        if( s != "" && s != "." ){
            if( inStr(c, "\\") > 0 && s== ".." ){
                c= mid(c, 1, inStrRev(c, "\\") - 1);
            }else{
                if( c != "" && right(c, 1) != "\\" ){ c= c + "\\" ;}
                c= c + s;
            }
        }
    }
    string fullPath= c;
    return fullPath;
}

//打印文件大小
string printFileSize(Int64 fileSize){
    if (fileSize < 0){
        throw new ArgumentOutOfRangeException("fileSize");
    }else if (fileSize >= 1024 * 1024 * 1024){
        return string.Format("{0:########0.00} GB", ((Double)fileSize) / (1024 * 1024 * 1024));
    }else if (fileSize >= 1024 * 1024){
        return string.Format("{0:####0.00} MB", ((Double)fileSize) / (1024 * 1024));
    }else if (fileSize >= 1024){
        return string.Format("{0:####0.00} KB", ((Double)fileSize) / 1024);
    }else{
        return string.Format("{0} bytes", fileSize);
    }
}
string handleZipSize(Int64 fileSize){
	return printFileSize(fileSize);
}
/*
'0：E:\E盘\WEB网站\至前网站\filename.asp
'1：E:\E盘\WEB网站\至前网站\      可以这么调用 HandleFilePathArray(FilePath)(1)
'2：filename.asp
'3：filename
'4：asp
*/
//文件处理成数组20150124  数组  0原文件路径 1为文件路径   2为文件名称  3为去除文件类型文件名称   4为文件类型后缀名
string[] handleFilePathArray( string filePath){
    string fileDir=""; string fileName=""; string fileNoTypeName=""; string fileType="";
    filePath= replace(filePath,"/","\\");

    fileDir= mid(filePath, 1, inStrRev(filePath, "\\"));
    fileName= mid(filePath, inStrRev(filePath, "\\") + 1,-1);
    if( inStrRev(fileName, ".") > 0 ){
        fileNoTypeName= mid(fileName, 1, inStrRev(fileName, ".") - 1);
    }else{
        fileNoTypeName= "";
    }
    fileType= mid(fileName, inStrRev(fileName, ".") + 1,-1);
 
    return aspSplit(filePath + vbCrlf() + fileDir + vbCrlf() + fileName + vbCrlf() + fileNoTypeName + vbCrlf() + fileType, vbCrlf());
}
//获得文件属性
string getFileAttr( string filePath, string sType){ 
	string s="";
    if( sType== "0" ){
        s= handleFilePathArray(filePath)[0];
    }else if( sType== "1" ){
        s= handleFilePathArray(filePath)[1];
    }else if( sType== "name" || sType== "2" ){
        s= handleFilePathArray(filePath)[2];
    }else if( sType== "3" ){
        s= handleFilePathArray(filePath)[3];
    }else if( sType== "4" ){
        s= handleFilePathArray(filePath)[4];
    }
    return s;
}
string getStrFileName(string s){
	return getFileAttr(s,"2");
}
string getFileName(string s){
	return getFileAttr(s,"2");
}

//去除文件路径 返回路径中的文件名部分
string removeFileDir( string filePath){
    filePath= replace(filePath, "\\", "/");
    if( inStr(filePath, "/") > 0 ){
        filePath= mid(filePath, inStrRev(filePath, "/") + 1,-1);
    }
    return filePath;
}
</script>
