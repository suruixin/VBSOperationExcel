Dim fs, dirName, fileData
'��ȡ�ļ����������ļ�'
'fs       �����ļ�'
'dirName  ��ȡ�����ļ���'
'fileData ��ȡ�����ļ�����ƴ�ӳ��ַ���'
Function getFileList()
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set dirName = fs.GetFolder(getDirPath())      '��ȡ�ļ���'	
	for each file in dirName.Files
		fileData = fileData + (file.name + "*|*") '��ȡ�ļ�����ƴ�ӳ��ַ���'
	Next
	setData()
	fileData = ""
	Set fs = nothing
	Set dirName = nothing
End Function

'���ļ���ִ��excel'
Function openFile(path)
	Dim oExcel, xlmodule, strCode, fE
	Set oExcel = CreateObject( "Excel.Application" )
	oExcel.Visible = True
	oExcel.DisplayAlerts = False
	Set fE = oExcel.WorkBooks.Open(path)
	Set xlmodule = fE.VBProject.VBComponents.Add(1)
	strCode = getMacros()

	xlmodule.CodeModule.AddFromString strCode
	fE.Application.Run "test"
	fE.Close
	oExcel.Quit
	Set oExcel= nothing
	Set fE= nothing
End Function