// シート名が変更されていないかどうか確認してください
const SETTINGS_SHEET_NAME = '設定';
const STUDENT_SHEET_NAME = '生徒一覧';
const INTERNSHIP_SHEET_NAME = 'インターンシップ受入れ先一覧';
const INTERNSHIP_STU_SHEET_NAME = 'インターンシップ生徒一覧';
const SAVE_FOLDER_NAME = "インターンシップ";


function getReceiptNum() {
	let ss = SpreadsheetApp.getActiveSpreadsheet()
	let sheet = ss.getSheetByName("事業所一覧")
	let data = sheet.getDataRange().getValues();
	let generatedNumbers = new Set(); // 生成済みの番号を追跡するためのSet

	for (let i = 0; i < data.length; i++) {
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
  // スプレッドシートの取得
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName(INTERNSHIP_SHEET_NAME)

  let lastRow = sheet.getLastRow()
  let colANumbers = sheet.getRange(2, 1, lastRow - 1).getValues();
  let colBNames = sheet.getRange(2, 2, lastRow - 1).getValues();
	// フォームのプルダウンリストに番号を付与する
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
	// スプレッドシートの取得
	let ss = SpreadsheetApp.getActiveSpreadsheet()

	// 受入れ先一覧と生徒一覧のシートを取得
	let internshipSheet = ss.getSheetByName(INTERNSHIP_SHEET_NAME);
	let internshipStuSheet = ss.getSheetByName(INTERNSHIP_STU_SHEET_NAME);

  // インターンシップ生徒一覧の背景色をクリア
	internshipStuSheet.getRange('E2:I').setBackground('yellow');
  // インターンシップ生徒一覧の背景色をクリア 
	internshipStuSheet.getRange('J2:J').setBackground(null);
	// 受入れ先一覧H列以降の内容をクリア
	internshipSheet.getRange('I2:AC201').clearContent();
  // 受入れ先人数の内容の背景色をクリア
	internshipSheet.getRange('D2:D201').setBackground('yellow');

	// 受入れ先一覧と生徒一覧のデータを取得
	let internshipList = internshipSheet.getRange('A2:F200').getValues();
	internshipList = internshipList.filter(internship => internship[3]); // 通し番号はあるが受入れ人数が空の行を除外
	let internList = internshipStuSheet.getRange('A2:I401').getValues();
	internList = internList.map((student, index) => ({row: index + 2, data: student})); // 元の行番号を追跡
	internList = internList.filter(student => student.data[5]); // 生徒番号はあるが第1希望欄が空の行を除外

	// 受入れ先一覧を通し番号をキーとするマップに変換
	let internshipMap = new Map();
	internshipList.forEach(internship => {
		let department = [];
		let departments = internship[4].split('、'); // 文字列を分割
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
		switch (internship[5]) {
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
			capacity: internship[3],
			department: department,
			sex: sex,
			students: []
		});
	});

	// 生徒一覧を成績（評定）順にソート
	internList.sort((a, b) => b.data[8] - a.data[8]);

  let specialStudent = internList.find(student => student.data[0] === 2330);
	if (specialStudent) {
			let firstChoice = Number(specialStudent.data[5].toString().match(/^\d{1,3}/)[0]);
			let firstChoiceInternship = internshipList.find(internship => internship[0] === firstChoice);
			if (firstChoiceInternship) {
					internshipMap.get(firstChoice).students.push(specialStudent.data);
					internshipStuSheet.getRange(specialStudent.row, 10).setValue(firstChoiceInternship[1]);
					internshipStuSheet.getRange(specialStudent.row, 6).setBackground('aqua');
					internList = internList.filter(student => student.data[0] !== 2330);
			}
	}

	// 成績順にソートされた生徒の第１希望（i=5）に受入れ先を割り当て、割り当てられなかった生徒だけで第２、第３と同じように割り当てていく
	for (let i = 5; i <= 7; i++) {
		for (let j = 0; j < internList.length; j++) {
				let student = internList[j];
				// student.dataにassignedプロパティが既に存在しているか確認し、存在していればループをスキップ
				if (student.data.assigned) continue;
				let internshipId = Number(student.data[i].toString().match(/^\d{1,3}/)[0]);
				let internship = internshipMap.get(internshipId);
				if (internship && 
								internship.students.length < internship.capacity &&
								internship.department.includes(student.data[3]) &&
								internship.sex.includes(student.data[2])) {
								internship.students.push(student.data);
								internshipStuSheet.getRange(student.row, 10).setValue(internship.name);
								internshipStuSheet.getRange(student.row, i+1).setBackground('aqua');
								student.data.assigned = true;
								// この時点で次のstudentへ進むため、breakではなくcontinueを使用
								continue;
				}
		}
	};
	internList.forEach(student =>{
		if (!student.data.assigned) {
			let range = internshipStuSheet.getRange(student.row, 10);
			range.setValue('希望する受入れ先に割り当てることができませんでした。');
			range.setBackground('yellow');
			student.data.assigned = false;
		}
	});


  // 割り当てられた生徒をインターンシップ受入れ先一覧に書き込む
  let row = 2; // I列の2行目から開始
  internshipMap.forEach((value, key) => {
    const assignedStudents = value.students.map(student => student[1]); // 生徒一覧B列にある生徒の氏名を取得
    if (assignedStudents.length > 0) { // assignedStudentsが空でないことを確認
      internshipSheet.getRange(row, 9, 1, assignedStudents.length).setValues([assignedStudents]); // 受入れ先一覧I列に生徒の氏名を書き込む
    }
    if (assignedStudents.length < value.capacity) { 
      internshipSheet.getRange(row, 4).setBackground('aqua'); // 割り当て人数が受入れ先の人数に満たない場合水色に設定
    } else {
      internshipSheet.getRange(row, 4).setBackground('red');// 満員の場合赤色に設定
    }
    row++; // assignedStudentsが空でも次の行に移動
  });
}

function InsertInternshipAssignments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INTERNSHIP_STU_SHEET_NAME);
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const date = settingsSheet.getRange('B36').getValue(); // 日付を取得

  let count = 0;
  const doc = DocumentApp.create('インターンシップ受入れ先通知文');
  const body = doc.getBody();

  data.forEach((row) => {
    const studentNumber = row[0]; // A列
    const studentName = row[1]; // B列
    const internshipName = row[9]; // J列

    if (internshipName) {
      const paragraph = body.appendParagraph(`${studentNumber}　${studentName}さんのインターンシップ先は、${internshipName}に決定しました。参加承諾書１枚、確約書２枚を記入し、${date}までに提出してください。`);
      paragraph.setSpacingAfter(128); // パラグラフの後にスペースを追加

			// パラグラフ内のテキスト要素を取得し、インターンシップ名の部分のスタイルを変更
			const text = paragraph.editAsText();
			const startIndex = `${studentNumber}　${studentName}さんのインターンシップ先は、`.length;
			const endIndex = startIndex + internshipName.length;
			text.setBold(startIndex, endIndex - 1, true);
			text.setUnderline(startIndex, endIndex - 1, true);
			text.setFontSize(startIndex, endIndex - 1, 16);

      count++;
    }

    if (count % 4 === 0) {
      body.appendPageBreak();
    }
  });

  // 変更を保存
  doc.saveAndClose();
}

function InsertInternshipRequest() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
	const studentsSheet = ss.getSheetByName(STUDENT_SHEET_NAME);
	const internshipSheet = ss.getSheetByName(INTERNSHIP_SHEET_NAME);

	const date = settingsSheet.getRange("B50").getValue(); // 日付
	const schoolName = settingsSheet.getRange("B3").getValue(); // 学校名
	const sender = settingsSheet.getRange("B52").getValue(); // 送信者名
	const title = settingsSheet.getRange("B54").getValue(); // 題名
	const bodyRequest = settingsSheet.getRange("B56").getValue(); // 本文（お願いする場合）
	const bodyRefusal = settingsSheet.getRange("B58").getValue(); // 本文（参加者がいない場合)
	const enclosure = settingsSheet.getRange("B60").getValue(); // 同封物
	const enquete = settingsSheet.getRange("B62").getValue(); // アンケートのお願い文
	const enqueteQrName = settingsSheet.getRange("B64").getValue(); // アンケートQRコードのファイル名
	const footer = settingsSheet.getRange("B66").getValue(); // 最下段に表示する署名

	// 指定した名前のフォルダを取得または作成
	const folderIterator = DriveApp.getRootFolder().getFoldersByName(SAVE_FOLDER_NAME);
	let saveFolder;
	if (folderIterator.hasNext()) {
		saveFolder = folderIterator.next();
	} else {
		saveFolder = DriveApp.getRootFolder().createFolder(SAVE_FOLDER_NAME);
	}
  
	let internships = internshipSheet.getRange("A2:C200").getValues();
	internships = internships.filter(internship => internship[1]); // 通し番号はあるが受入れ先名が空の行を除外
	internships.forEach((internship, index) => {
	  const doc = DocumentApp.create(internship[1] + " - インターンシップ受入れお願い"); // 新しい文書を作成
	  const docBody = doc.getBody();
	  
	  docBody.appendParagraph("（公　印　省　略）" + "\n" + date).setAlignment(DocumentApp.HorizontalAlignment.RIGHT); // 日付
	  docBody.appendParagraph(internship[1] + "\n" + internship[2] + "　様").setAlignment(DocumentApp.HorizontalAlignment.LEFT); // 事業所名と担当者名
	  docBody.appendParagraph(schoolName).setAlignment(DocumentApp.HorizontalAlignment.RIGHT); // 学校名
	  docBody.appendParagraph(sender).setAlignment(DocumentApp.HorizontalAlignment.RIGHT); // 送信者名
	  
	  // 参加する生徒名の処理
		const studentNameInI = internshipSheet.getRange("I" + (index + 2)).getValue();// インターンシップ受入れ先I列に生徒名があるか確認
		if (studentNameInI !== "") {
			docBody.appendParagraph(title).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(16); // 題名
			docBody.appendParagraph(bodyRequest).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(16); // 本文
			docBody.appendParagraph("記").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(16);
			docBody.appendParagraph("参加生徒一覧").setAlignment(DocumentApp.HorizontalAlignment.L).setSpacingBefore(16).setSpacingAfter(16);
			
			// インターンシップ受入れ先一覧から生徒名を取得
			const internNamesRange = internshipSheet.getRange("I" + (index + 2) + ":AD" + (index + 2));
			const internNames = internNamesRange.getValues()[0].filter(name => name !== ""); // 空でない値をフィルタリング
			const studentNames = studentsSheet.getRange("B2:B" + studentsSheet.getLastRow()).getValues().flat();
			let internList = [];
			internNames.forEach(internName => {
				// 生徒名に基づいてstudentsSheetからインデックスを取得
				const studentIndex = studentNames.findIndex(name => name === internName);
				if (studentIndex !== -1) { // 生徒が見つかった場合
					// 学科情報を取得（D列に学科コードがあると仮定）
					const departmentCode = studentsSheet.getRange("D" + (studentIndex + 2)).getValue();
					let departmentName;
					// 学科コードに基づいて学科名を設定
					switch (departmentCode) {
					case 1:
						departmentName = "機械科";
						break;
					case 2:
						departmentName = "電気科";
						break;
					case 3:
						departmentName = "商業科";
						break;
					default:
						departmentName = ""; // 学科コードが不明の場合
					}
					// 学科名と生徒名を組み合わせてinternListに追加
					internList.push(`${departmentName}　${internName}`);
				}
			})
			// 参加生徒を表示
			internList.forEach(item => {
				docBody.appendParagraph(item).setIndentStart(36);
			});
			
			// 同封物を表示
			docBody.appendParagraph("同封物").setAlignment(DocumentApp.HorizontalAlignment.L).setSpacingBefore(16);
			docBody.appendParagraph(enclosure).setIndentStart(36);
			
			// アンケートのお願いを表示するテーブルを作成
			docBody.appendParagraph("アンケートのお願い")
				.setAlignment(DocumentApp.HorizontalAlignment.L)
				.setSpacingBefore(16)
			let table = docBody.appendTable();
			let row = table.appendTableRow();
			// テーブルの左セルにテキストを配置
			let cell1 = row.appendTableCell();
			cell1.appendParagraph(enquete)
			cell1.removeChild(cell1.getChild(0)); // 余分な最初の行を削除
			// Google ドライブから QR コードの画像を取得し、テーブルの次のセルに挿入
			let files = DriveApp.getFilesByName(enqueteQrName);
			if (files.hasNext()) {
				let imageFile = files.next();
				let image = imageFile.getBlob();
				let cell2 = row.appendTableCell("");
				cell2.appendImage(image).setHeight(75).setWidth(75);
				cell2.removeChild(cell2.getChild(0)); // 余分な最初の行を削除
			} else {
				console.log("指定されたQRコードの画像が見つかりません。");
			}
			// テーブルの境界線を非表示にする
			table.setBorderWidth(0);
		} else {
			// 参加者がいない場合の処理
			docBody.appendParagraph(title).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(32); // 題名
			docBody.appendParagraph(bodyRefusal).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(32); // 本文
		}
			// 最下段に表示する署名
			docBody.appendParagraph(footer).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingBefore(32);
		
			// 文書のURLをログに出力（確認用）
			console.log("Document created: " + doc.getUrl());

			// ファイルをフォルダ内に移動
			docFile = DriveApp.getFileById(doc.getId());
			saveFolder.addFile(docFile);
	});
}
