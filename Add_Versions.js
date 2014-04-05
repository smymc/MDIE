/*
	AddVersion
	ファイル内の最初のVer表記を探してファイル名につける
*/
var strScriptPath;
var strOpenFile;
var objTextStream;
var strText;

//前に付加か後ろに付加を指定する
var isBefore = false;

var FSO = new ActiveXObject("Scripting.FileSystemObject");

function main(){
	var i;
	for(i=0;i<FolderView.Count;i++){
		if(FolderView.Items(i).Selected){
			submain(FolderView.Items(i).Path);
		}
	}
	return;
}
function submain(pth){
	var f,fn;
	var isFolder = true;
	var tmp;
	
	var addVer = AddVersion(pth);
	if ( addVer == "" ) {
		return;
	}
	
	f = (isFolder=FSO.FolderExists(pth))?FSO.GetFolder(pth):FSO.GetFile(pth);
	// （条件式）?（真の処理）:（偽の処理）
	fn = isBefore?(addVer + f.Name):(isFolder?(f.Name + addVer):
		(FSO.GetBaseName(f.Name) + addVer + "." + FSO.GetExtensionName(f.Name)));
	
	tmp = fn;
	
	//Delete 「～コピー 」
	tmp = tmp.replace(/コピー\s～\s/, "");
	tmp = tmp.replace(/ - copy \[\d\]/, "");
	// MDIE.echo( tmp + "\n" );
	fn = tmp;
	
	if ( FSO.FileExists(f.ParentFolder.Path + "\\" + fn)
	  || FSO.FolderExists(f.ParentFolder.Path + "\\" + fn)) {
		MDIE.echo( fn + "\n" + "同名ファイル・フォルダが存在します" );
		return;
	}
	f.Name = fn;
	return;
}

function AddVersion(pth) {
	strOpenFile   = pth;

	objTextStream = FSO.OpenTextFile(strOpenFile);
	
	var iCheck;
	var strMatch,strVersion;

	while (objTextStream.AtEndOfLine==false) {
		strText = objTextStream.ReadLine();

		//Trim Space
		strText = strText.replace(/^\s+/, "");
		strText = strText.replace(/\s+$/, "");

		strMatch = strText.match(/.*(Ver).*/i);
		
		if ( strMatch != undefined ) {
			strVersion = strMatch[0].replace(/Ver/, "_v");
			
			//Trim Mid Space
			strVersion = strVersion.replace(/\s/, "");
			break;
		}
	}

	objTextStream.Close();

	objTextStream = null;
	//FSO = null;
	if ( strMatch == undefined ) {
		strVersion = "";
	}
	return strVersion;
}
main();
