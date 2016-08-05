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
/*
�����ڣ�20160724
����	���ܺ���;
Attributes	���غ��ļ���ص�����ֵ��������FileAttributesö������ֵ
CreationTime	�����ļ��Ĵ���ʱ��
Exists	����ļ��Ƿ�����ڸ�����Ŀ¼��
Extension	�����ļ�����չ��
LastAccessTime	�����ļ����ϴη���ʱ��
FullName	�����ļ��ľ���·��
LastWriteTime	�����ļ����ϴ�д����ʱ��
Name	���ظ����ļ����ļ���
Delete����	ɾ��һ���ļ��ķ���������ؽ��������ø÷���
*/

 

//���ļ�
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
//�����ļ�
bool createFile(string filePath, string content){
	filePath=handlePath(filePath);
	deleteFile(filePath);			//��ɾ���ļ���Ҫ��Ȼ������
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
//����gbk�ļ������Ľ�
bool createFileGBK(string filePath, string content){
	return createFile(filePath,content);
}
//����gbk�ļ������Ľ�
bool writeToFile(string filePath, string content, string setFileCode){
	return createFile(filePath,content);
}

//�����ۼ��ļ�
bool createAddFile(string filePath, string content){
	string c=getFText(filePath)+content;
	return createFile(filePath,content);
}
//����ļ�
bool checkFile(string filePath){
	filePath=handlePath(filePath);
	if(File.Exists(@filePath)){
		return true;
	}else{
		return false;
	} 
}
//�ƶ��ļ�
bool moveFile(string filePath,string toFilePath){
	filePath=handlePath(filePath);
	toFilePath=handlePath(toFilePath);
	if(checkFile(filePath)==true && checkFile(toFilePath)==false){	
��������File.Move(filePath,toFilePath);
		return true;
	}else{
		return false;
	}
}
//�����ļ�
bool copyFile(string filePath,string toFilePath){
	filePath=handlePath(filePath);
	toFilePath=handlePath(toFilePath);
	if(checkFile(filePath)==true && checkFile(toFilePath)==false){	
��������File.Copy(filePath,toFilePath);
		return true;
	}else{
		return false;
	}
}
//ɾ���ļ�
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
//����ļ���С
int getFSize(string filePath){
	filePath=handlePath(filePath); 
	//�жϵ�ǰ·����ָ����Ƿ�Ϊ�ļ�
	if (File.Exists(filePath) == true){ 
		//�Ի�ȡ���С
		FileInfo fileInfo = new FileInfo(filePath);
		return cInt(fileInfo.Length); 
	}else{
		return 0;
	}
}

//����ļ�����ʱ��
string getFileCreateTime(string filePath){
	filePath=handlePath(filePath); 
	if(File.Exists(@filePath)==false){
		return "";
	}
	FileInfo fi = new FileInfo(filePath);
	return  fi.CreationTime.ToString();
}
//����ļ��޸�ʱ��
string getFileEditTime(string filePath){
	filePath=handlePath(filePath); 
	if(File.Exists(@filePath)==false){
		return "";
	}
	FileInfo fi = new FileInfo(filePath);
	return  fi.LastWriteTime.ToString();
}
//����ļ�����ʱ��
string getFileVisitTime(string filePath){
	filePath=handlePath(filePath); 
	if(File.Exists(@filePath)==false){
		return "";
	}
	FileInfo fi = new FileInfo(filePath);
	return  fi.LastAccessTime.ToString();
}



//�����ļ���
bool createFolder(string folderPath){
	folderPath=handlePath(folderPath);
	if(!Directory.Exists(@folderPath)){
		Directory.CreateDirectory(folderPath);		
		return true;
	}else{
		return false;
	} 
} 
//��������Ŀ¼ .net�ﲻ����createFolder�Ϳ�����������Ŀ¼
bool createDirFolder(string folderPath){
	return createFolder ( folderPath );
}
//����ļ���
bool checkFolder(string folderPath){
	folderPath=handlePath(folderPath);
	if(Directory.Exists(@folderPath)){
		return true;
	}else{
		return false;
	} 
}
//ɾ���ļ���
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
//�ƶ��ļ���
bool moveFolder(string folderPath, string toFolderName){
	folderPath=handlePath(folderPath);
	toFolderName=handlePath(toFolderName);
	DirectoryInfo di = new DirectoryInfo(folderPath); 
	DirectoryInfo di2 = new DirectoryInfo(toFolderName); 
	if (!di.Exists){ 
		//Դ�ļ�������
		return false ; 
	} 
	if (di2.Exists){ 
		//Ŀ���ļ����Ѵ���
		return false;
	}else{
		di.MoveTo(toFolderName); 
		return true;
	}
}
//�����ļ���
bool copyFolder(string folderPath, string toFolderName){
	folderPath=handlePath(folderPath);
	toFolderName=handlePath(toFolderName);
 	if (!Directory.Exists(folderPath)){
        return false;
	} 
	//����Ŀ¼�ļ���
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
//����ļ����޸�ʱ��
string getFolderEditTime(string folderPath){
	folderPath=handlePath(folderPath); 
	if(Directory.Exists(@folderPath)==false){
		return "";
	}
	DirectoryInfo dir = new DirectoryInfo(folderPath); 
	return  dir.LastWriteTime.ToString();
}
//����ļ��з���ʱ��
string getFolderVisitTime(string folderPath){
	folderPath=handlePath(folderPath); 
	if(Directory.Exists(@folderPath)==false){
		return "";
	}
	DirectoryInfo dir = new DirectoryInfo(folderPath); 
	return  dir.LastAccessTime.ToString();
}
 
//����ļ��д���ʱ��
string getFolderCreateTime(string folderPath){
	folderPath=handlePath(folderPath); 
	if(Directory.Exists(@folderPath)==false){
		return "";
	}
	DirectoryInfo dir = new DirectoryInfo(folderPath); 
	return  dir.CreationTime.ToString();
} 


// ������ļ��ļ��С�
string getFileFolderList(string folderPath, string c, string action, string fileTypeList){
	folderPath=handlePath(folderPath); 
	string fileType;string s;string folderName;string tempS;
	if(action==""){
		action="|�����ļ�|�����ļ���|�ļ�����|�ļ�������|ѭ���ļ���|";
	}
	//Response.Write("action="+action);
	if(fileTypeList==""){
		fileTypeList="|*|";
	}
	//�ж��ļ��Ƿ����
	if(Directory.Exists(@folderPath)==false){
		return c;
	}
	//�ļ�
	if(inStr("|" + action + "|", "|�����ļ�|")>0){
		string[] files = Directory.GetFiles(folderPath);
		foreach (var filePath in files){
			string fileName= Path.GetFileName(filePath);
			fileType="";
			if(inStr(fileName,".")>0){
				fileType=lCase(mid(fileName,inStrRev(fileName,".")+1));
			}
			if(inStr("|" + action + "|", "|�ļ�����|")>0){
				s=fileName;
			}else if(inStr("|" + action + "|", "|�ļ�����|")>0){
				s=fileType;
			}else{ 
				s=filePath;
			}
			if(inStr("|" + fileTypeList + "|", "|"+ fileType +"|")>0 || inStr("|" + fileTypeList + "|", "|*|")>0){
				c=c+s+vbCrlf();
			} 
		}
	}
	//�ļ��� 
    string[] fromDirs = Directory.GetDirectories(folderPath);
    foreach (string fromDirName in fromDirs) {
        folderName = Path.GetFileName(fromDirName);
		//Response.Write("�ļ���="+fromDirName+"<hr>");  
		
		if(inStr("|" + action + "|", "|�����ļ���|")>0){
			if(inStr("|" + action + "|", "|�ļ�������|")>0){
				s = folderName;
			}else{
				s=folderPath;
			}
			c=c+s+vbCrlf();
		}
		if(inStr("|" + action + "|", "|ѭ���ļ���|")>0){
			tempS=getFileFolderList(fromDirName,"",action,fileTypeList);
			if(tempS!=""){
				c=c+tempS+vbCrlf();
			}
		}
		
		  
    }
	return c;
	 
}   

//��õ�ǰĿ¼��ȫ��Jpg�ļ�
string getDirJpgList(string folderPath){
	return getDirFileList(folderPath, "jpg");
}
//��õ�ǰĿ¼��ȫ��Png�ļ�
string getDirPngList(string folderPath){
	return getDirFileList(folderPath, "png");
}
//��õ�ǰĿ¼��ȫ��Ini�ļ�
string getDirIniList(string folderPath){
	return getDirFileList(folderPath, "ini");
}
//��õ�ǰĿ¼��ȫ��Txt�ļ�
string getDirTxtList(string folderPath){
	return getDirFileList(folderPath, "txt");
}
//��õ�ǰĿ¼��ȫ��Js�ļ�
string getDirJsList(string folderPath){
	return getDirFileList(folderPath, "js");
}
//��õ�ǰĿ¼��ȫ��Css�ļ�
string getDirCssList(string folderPath){
	return getDirFileList(folderPath, "css");
}
//��õ�ǰĿ¼��ȫ��Html�ļ�
string getDirHtmlList(string folderPath){
	return getDirFileList(folderPath, "html");
}
//��õ�ǰĿ¼��ȫ��asp�ļ�
string getDirAspList(string folderPath){
	return getDirFileList(folderPath, "asp");
}
//��õ�ǰĿ¼��ȫ��Php�ļ�
string getDirPhpList(string folderPath){
	return getDirFileList(folderPath, "php");
}

//��õ�ǰĿ¼�������ļ��б�
string getDirFileList(string folderPath,string fileTypeList){
	if(fileTypeList==""){
		fileTypeList="*";
	} 
	return getFileFolderList(folderPath,"","|�����ļ�|","|"+fileTypeList+"|");
} 


//��õ�ǰĿ¼��ȫ��Jpg�ļ�����
string getDirJpgNameList(string folderPath){
	return getDirFileNameList(folderPath, "jpg");
}
//��õ�ǰĿ¼��ȫ��Png�ļ�����
string getDirPngNameList(string folderPath){
	return getDirFileNameList(folderPath, "png");
}
//��õ�ǰĿ¼��ȫ��Ini�ļ�����
string getDirIniNameList(string folderPath){
	return getDirFileNameList(folderPath, "ini");
}
//��õ�ǰĿ¼��ȫ��Txt�ļ�����
string getDirTxtNameList(string folderPath){
	return getDirFileNameList(folderPath, "txt");
}
//��õ�ǰĿ¼��ȫ��Js�ļ�����
string getDirJsNameList(string folderPath){
	return getDirFileNameList(folderPath, "js");
}
//��õ�ǰĿ¼��ȫ��Css�ļ�����
string getDirCssNameList(string folderPath){
	return getDirFileNameList(folderPath, "css");
}
//��õ�ǰĿ¼��ȫ��Html�ļ�����
string getDirHtmlNameList(string folderPath){
	return getDirFileNameList(folderPath, "html");
}
//��õ�ǰĿ¼��ȫ��asp�ļ�����
string getDirAspNameList(string folderPath){
	return getDirFileNameList(folderPath, "asp");
}
//��õ�ǰĿ¼��ȫ��Php�ļ�����
string getDirPhpNameList(string folderPath){
	return getDirFileNameList(folderPath, "php");
}
//��õ�ǰĿ¼�������ļ������б�
string getDirFileNameList(string folderPath,string fileTypeList){
	if(fileTypeList==""){
		fileTypeList="*";
	}
	//aspecho("fileTypeList",fileTypeList);
	return getFileFolderList(folderPath,"","|�����ļ�|�ļ�����|","|"+fileTypeList+"|");
}

//Ŀ¼�ļ����� ������
string getDirFileSort(string folderPath){
	string[] splstr;string content="";string myFile="";string c="";int i=0;int nIndex=0;string[] arrayStr=new string[33];;
	content=getDirFileNameList(folderPath,"");
	splstr=aspSplit(content,vbCrlf());
	foreach(var eachmyFile in splstr){
		myFile=eachmyFile;
		if( inStr(myFile, "#")== 0 && inStr(myFile, "��") > 0 ){ //#������ʾ������Ҫ�С�����
			nIndex= cInt(replace(left(myFile, 2), "��", ""));
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


 //��õ�ǰĿ¼�������ļ��б�
string getDirAllFileList(string folderPath,string fileTypeList){
	if(fileTypeList==""){
		fileTypeList="*";
	}
	//aspecho("fileTypeList",fileTypeList);
	return getFileFolderList(folderPath,"","|�����ļ�|ѭ���ļ���|hidefolderlist|","|"+fileTypeList+"|");
} 
//��õ�ǰĿ¼�������ļ������б�
string getDirAllFileNameList(string folderPath,string fileTypeList){
	if(fileTypeList==""){
		fileTypeList="*";
	}
	//aspecho("fileTypeList",fileTypeList);
	return getFileFolderList(folderPath,"","|�����ļ�|�ļ�����|ѭ���ļ���|hidefolderlist|","|"+fileTypeList+"|");
}


//��õ�ǰĿ¼���ļ���
string getDirFolderList(string folderPath){ 
	return getFileFolderList(folderPath,"","|�����ļ���|","");
}
//��õ�ǰĿ¼���ļ���
string getDirFolderNameList(string folderPath){ 
	return getFileFolderList(folderPath,"","|�����ļ���|�ļ�������|","");
}
//��õ�ǰĿ¼��ȫ���ļ���
string getDirAllFolderList(string folderPath){ 
	return getFileFolderList(folderPath,"","|�����ļ���|ѭ���ļ���|","");
}
//��õ�ǰĿ¼��ȫ���ļ���
string getDirAllFolderNameList(string folderPath){ 
	return getFileFolderList(folderPath,"","|�����ļ���|ѭ���ļ���|�ļ�������|","");
}



/////////////////////////////////////////  


//��������·�� (2013,9,27
string handlePath(string fFPath){ //Pathǰ�治��ByVal �ض��壬������Ϊ����ǰ�溯�������ʹ�����·����������
    fFPath= replace(fFPath, "/", "\\");
    fFPath= replace(fFPath, "\\\\", "\\");
    fFPath= replace(fFPath, "\\\\", "\\");
    bool isDir; //ΪĿ¼
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
//����·��
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

//��ӡ�ļ���С
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
'0��E:\E��\WEB��վ\��ǰ��վ\filename.asp
'1��E:\E��\WEB��վ\��ǰ��վ\      ������ô���� HandleFilePathArray(FilePath)(1)
'2��filename.asp
'3��filename
'4��asp
*/
//�ļ����������20150124  ����  0ԭ�ļ�·�� 1Ϊ�ļ�·��   2Ϊ�ļ�����  3Ϊȥ���ļ������ļ�����   4Ϊ�ļ����ͺ�׺��
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
//����ļ�����
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

//ȥ���ļ�·�� ����·���е��ļ�������
string removeFileDir( string filePath){
    filePath= replace(filePath, "\\", "/");
    if( inStr(filePath, "/") > 0 ){
        filePath= mid(filePath, inStrRev(filePath, "/") + 1,-1);
    }
    return filePath;
}
</script>
