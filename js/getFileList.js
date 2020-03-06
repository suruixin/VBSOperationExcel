/**
 * 输出html
 * @param fileList 可执行文件列表
 */
function node(fileList) {
	var node = "<table>" +
		"<theader>" +
			"<tr>" +
				"<th>文件名</th>" +
				"<th>文件路径</th>" +
				"<th>操作</th>" +
			"</tr>" +
		"</theader>" +
		"<tbody>";
	for (var i = 0; i < fileList.length; i++) {
		node += "<tr>" +
			"<td>" + fileList[i].fileName + "</td>" +
			"<td>" + fileList[i].filePath + "</td>" +
			"<td><button onClick='vbscript:openFile(\"" + fileList[i].filePath + "\")'>执行vb命令</button></td>" +
		"</tr>"
	}
	node += "</tbody></table>"
	return node;
}

/**
 * 输出文件路径
 */
function getDirPath () {
	if (!path.value) alert('文件夹为必填');
	return path.value;
}

/**
 * 处理获取到的文件列表
 */
function setData () {
	if (!fileData) { // 判断是否读取到文件
		alert("未读取到文件夹下有xls,xlsx,csv文件");
		return false
	}
	var arr = []; // 创建一个新数组用来存储有效文件
	var fileArr = fileData.split('*|*'); // 将获取到的文件进行处理
	for(var i = 0; i < fileArr.length - 1; i++) {
		var fileSuffix = fileArr[i].split('.')[fileArr[i].split('.') && fileArr[i].split('.').length - 1]; // 获取文件后缀
		if (fileArr[i] !== '' && (fileSuffix === 'xls' || fileSuffix === 'xlsx' || fileSuffix === 'csv')) {
			arr.push({
				fileName: fileArr[i],
				filePath: getDirPath() + '\\\\' + fileArr[i]
			})
		}
	};
	fileWrapper.innerHTML = node(arr);
}

/**
 * 获取宏命令
 */
function getMacros () {
	var dom = document.getElementById('macros');
	return dom.value
}