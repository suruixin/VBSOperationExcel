Dim fs, dirName, fileData
'��ȡ�ļ����������ļ�'
'fs       �����ļ�'
'dirName  ��ȡ�����ļ���'
'fileData ��ȡ�����ļ�����ƴ�ӳ��ַ���'
Function getFileList()
	set fs = CreateObject("Scripting.FileSystemObject")
	set dirName = fs.GetFolder(getDirPath())      '��ȡ�ļ���'	
	for each file in dirName.Files
		fileData = fileData + (file.name + "*|*") '��ȡ�ļ�����ƴ�ӳ��ַ���'
	Next
	setData()
End Function

'���ļ���ִ��excel'
Function openFile(path)
	MsgBox(path)
End Function