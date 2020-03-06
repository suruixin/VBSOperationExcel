/**
 * ���html
 * @param fileList ��ִ���ļ��б�
 */
function node(fileList) {
	var node = "<table>" +
		"<theader>" +
			"<tr>" +
				"<th>�ļ���</th>" +
				"<th>�ļ�·��</th>" +
				"<th>����</th>" +
			"</tr>" +
		"</theader>" +
		"<tbody>";
	for (var i = 0; i < fileList.length; i++) {
		node += "<tr>" +
			"<td>" + fileList[i].fileName + "</td>" +
			"<td>" + fileList[i].filePath + "</td>" +
			"<td><button onClick='vbscript:openFile(\"" + fileList[i].filePath + "\")'>ִ��vb����</button></td>" +
		"</tr>"
	}
	node += "</tbody></table>"
	return node;
}

/**
 * ����ļ�·��
 */
function getDirPath () {
	if (!path.value) alert('�ļ���Ϊ����');
	return path.value;
}

/**
 * �����ȡ�����ļ��б�
 */
function setData () {
	if (!fileData) { // �ж��Ƿ��ȡ���ļ�
		alert("δ��ȡ���ļ�������xls,xlsx,csv�ļ�");
		return false
	}
	var arr = []; // ����һ�������������洢��Ч�ļ�
	var fileArr = fileData.split('*|*'); // ����ȡ�����ļ����д���
	for(var i = 0; i < fileArr.length - 1; i++) {
		var fileSuffix = fileArr[i].split('.')[fileArr[i].split('.') && fileArr[i].split('.').length - 1]; // ��ȡ�ļ���׺
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
 * ��ȡ������
 */
function getMacros () {
	var dom = document.getElementById('macros');
	return dom.value
}