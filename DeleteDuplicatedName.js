/*
	Delete Duplicated Name
	ファイル名から「コピー」を削除する
	
	削除対象文字列：
		「コピー ～ 」
		「 - コピー」
		「 - copy [#] 」
		「(#)」
*/

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
	
	f = (isFolder=FSO.FolderExists(pth))?FSO.GetFolder(pth):FSO.GetFile(pth);
	fn = FSO.GetBaseName(f.Name) + "." + FSO.GetExtensionName(f.Name);
	
	tmp = fn;
	
	//Delete Duplicated Name
	tmp = tmp.replace(/コピー\s～\s/, "");
	tmp = tmp.replace(/\s-\sコピー/, "");
	tmp = tmp.replace(/ - copy \[\d*\]/, "");
	tmp = tmp.replace(/\(\d*\)/, "");
	fn = tmp;
	
	if ( FSO.FileExists(f.ParentFolder.Path + "\\" + fn)
	  || FSO.FolderExists(f.ParentFolder.Path + "\\" + fn)) {
		MDIE.echo( fn + "\n" + "同名ファイル・フォルダが存在します" );
		return;
	}
	f.Name = fn;
	return;
}

main();
