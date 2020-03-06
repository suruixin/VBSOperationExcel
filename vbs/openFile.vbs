Dim fs, dirName, fileData
'获取文件夹下所有文件'
'fs       操作文件'
'dirName  获取到的文件夹'
'fileData 获取到的文件名并拼接成字符串'
Function getFileList()
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set dirName = fs.GetFolder(getDirPath())      '获取文件夹'	
	for each file in dirName.Files
		fileData = fileData + (file.name + "*|*") '获取文件名并拼接成字符串'
	Next
	setData()
	fileData = ""
	Set fs = nothing
	Set dirName = nothing
End Function

'打开文件并执行excel'
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