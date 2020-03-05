Dim fs, dirName, fileData
'获取文件夹下所有文件'
'fs       操作文件'
'dirName  获取到的文件夹'
'fileData 获取到的文件名并拼接成字符串'
Function getFileList()
	set fs = CreateObject("Scripting.FileSystemObject")
	set dirName = fs.GetFolder(getDirPath())      '获取文件夹'	
	for each file in dirName.Files
		fileData = fileData + (file.name + "*|*") '获取文件名并拼接成字符串'
	Next
	setData()
End Function

'打开文件并执行excel'
Function openFile(path)
	MsgBox(path)
End Function