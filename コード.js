function getReceiptNum() {
	let ss = SpreadsheetApp.getActiveSpreadsheet()
	let sheet = ss.getSheetByName("事業所一覧")
	let data = sheet.getDataRange().getValues();
	let generatedNumbers = new Set(); // 生成済みの番号を追跡するためのSet

	for (var i = 0; i < data.length; i++) {
		let furigana = data[i][5]; // F列目のフリガナ

		// 住所に基づいて1桁目の番号を割り当て
		let firstDigit;
		let address = data[i][8]; // I列目の住所
		if (address.includes("相生市")) {
				firstDigit = 1;
		} else if (address.includes("赤穂市") || address.includes("赤穂郡")) {
				firstDigit = 2;
		} else if (address.includes("たつの市") || address.includes("太子町")) {
				firstDigit = 3;
		} else if (address.includes("姫路市")) {
				firstDigit = 4;
		} else if (address.includes("兵庫県")){
				firstDigit = 5;
		} else {
				firstDigit = 6;
		}
		// フリガナの先頭文字に基づいて2桁目の番号を割り当て
		let secondDigit;
		let firstChar = furigana.charAt(0);
		if (firstChar === '') {
			firstChar.trimStart();
		}
		let furiganaGroups = ['アイウエオ', 'カキクケコガギグゲゴ', 'サシスセソザジズゼゾ', 'タチツテトダヂヅデド', 'ナニヌネノ', 'ハヒフヘホバビブベボパピプペポ', 'マミムメモ', 'ヤユヨ', 'ラリルレロ','ワヲン'];
		secondDigit = furiganaGroups.findIndex(group => group.includes(firstChar)) + 1;

		let receiptNum = data[i][0]; // A列目の既存の番号

		if (!receiptNum) { // 既存の番号がない場合のみ新たに生成
			do {
					// 3桁目と4桁目をランダムに生成
					let thirdDigit = Math.floor(Math.random() * 10);
					let fourthDigit = Math.floor(Math.random() * 10);

					// 4桁の番号を生成
					receiptNum = firstDigit * 1000 + secondDigit * 100 + thirdDigit * 10 + fourthDigit;
			} while (generatedNumbers.has(receiptNum)); // すでに生成された番号であれば再度ランダムな数字を生成

			// 生成した番号をSetに追加
			generatedNumbers.add(receiptNum);

			// 番号をシートの1列目に入力
			sheet.getRange(i + 1, 1).setValue(receiptNum);
		}
	}
}

function overwritePdwList() {
  //スプレッドシートの情報を取得する
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName("インターンシップ受入れ先一覧")

  let lastRow = sheet.getLastRow()
  let colANumbers = sheet.getRange(2, 1, lastRow - 1).getValues();
  let colBNames = sheet.getRange(2, 2, lastRow - 1).getValues();

  let colPdwList = colANumbers.map((value, index) => {
		if (value[0] && colBNames[index][0]) {
			return value[0] + ' ' + colBNames[index][0];
		}
	}).filter(Boolean);

  // Googleフォームのプルダウン内の値を上書きする
  let form = FormApp.openById('1znvvf3TF8NNMPEf4MqUnFLOB8-FE0F8Zlv89giG1P7Q');
  let items = form.getItems(FormApp.ItemType.LIST);
  items.forEach(function(item){
    if(item.getTitle().match(/インターンシップ先を選.*$/)){
      let listItemQuestion = item.asListItem();
      let choices = [];
      colPdwList.forEach(function(name){
        if(name != ""){
					choices.push(listItemQuestion.createChoice(name));
        }
      });
      listItemQuestion.setChoices(choices);
    }
  });
}

function assignInternship() {
	// Googleスプレッドシートの取得
	let ss = SpreadsheetApp.getActiveSpreadsheet()

	// 受入れ先一覧と生徒一覧のシートを取得
	let internshipSheet = ss.getSheetByName('インターンシップ受入れ先一覧');
	let studentSheet = ss.getSheetByName('インターンシップ生徒一覧');

	// 受入れ先一覧と生徒一覧のデータを取得
	let internshipList = internshipSheet.getRange('A2:F200').getValues();
	internshipList = internshipList.filter(internship => internship[2]); // 通し番号はあるが受入れ人数が空の行を除外
	let studentList = studentSheet.getRange('A2:I401').getValues();
	studentList = studentList.map((student, index) => ({row: index + 2, data: student})); // 元の行番号を追跡
	studentList = studentList.filter(student => student.data[5]); // 生徒番号はあるが第1希望欄が空の行を除外

	// 受入れ先一覧を通し番号をキーとするマップに変換
	let internshipMap = new Map();
	internshipList.forEach(internship => {
		let department = [];
		let departments = internship[3].split('、'); // 文字列を分割
		departments.forEach(dep => {
			switch (dep) {
				case '機械科':
					department.push(1);
					break;
				case '電気科':
					department.push(2);
					break;
				case '商業科':
					department.push(3);
					break;
				case '全学科':
					department = [1, 2, 3];
					break;
			}
		});
	
		let sex;
		switch (internship[4]) {
			case '男子':
				sex = [1];
				break;
			case '女子':
				sex = [2];
				break;
			case '男女可':
				sex = [1, 2];
				break;
			default:
				sex = [];
		}
	
		internshipMap.set(internship[0], {
			name: internship[1],
			capacity: internship[2],
			department: department,
			sex: sex,
			students: []
		});
	});

	// 生徒一覧を成績（評定）順にソート
	studentList.sort((a, b) => b.data[8] - a.data[8]);

	// ソートされた生徒一覧をループ
	studentList.forEach((student) => {
		let assigned = false;
		for (let i = 5; i <= 7; i++) {
			let internship = internshipMap.get(student.data[i]);
			if (internship && 
				internship.students.length < internship.capacity && // 定員に達していないか確認
				internship.department.includes(student.data[3]) && // 学科が一致するか確認
				internship.sex.includes(student.data[2])) { // 性別が一致するか確認
				internship.students.push(student.data);
				studentSheet.getRange(student.row, 10).setValue(internship.name); // 生徒一覧のJ列に受入れ先を設定
				assigned = true;
				break;
			}
		}
		if (!assigned) {
			let range = studentSheet.getRange(student.row, 10);
			range.setValue('希望する受入れ先に割り当てることができませんでした。'); // J列にメッセージを設定
			range.setBackground('yellow'); // セルの背景を黄色に設定
		}
	});
}