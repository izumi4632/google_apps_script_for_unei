/**
 * スプレッドシート表示の際に呼出し
 */
function onOpen() {
	
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	
	//スプレッドシートのメニューにカスタムメニュー「カレンダー連携 > 実行」を作成
	var subMenus = [];
	subMenus.push({
		name: "実行",
		functionName: "createSchedule"  //実行で呼び出す関数を指定
	});
	ss.addMenu("カレンダー連携", subMenus);
}


/**
 * 予定を作成する
 */
function createSchedule() {
	
	// 読み取り範囲（表の始まり行と終わり列）
	const topRow = 6;
	const lastCol = 20;
	
	// 0始まりで列を指定しておく
	const statusCellNum = 1;
	const dayCellNum = 2;
	const enddayCellNum = 4;
	const startCellNum = 6;
	const endCellNum = 7;
	const titleCellNum = 8;
	const locationCellNum = 9;
	const descriptionCellNum = 10;
	const guestsCellNum = 11;
	const recurrenceMethodCellNum = 12;
	const colorCellNum = 13;
	
	var recurrenceMethod = [
		"0-origin is not suitable for users",
		// 方法                NO  説明
		"addDailyRule()",     // 1   毎日繰り返す。
		"addMonthlyRule()",   // 2   毎月繰り返す。
		"addWeeklyRule()",    // 3   毎週繰り返す。
		"addYearlyRule()",    // 4   毎年繰り返す。
	];
	
	// シートを取得
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	
	// 連携するアカウント
	const gAccount = sheet.getRange("J3").getValues()[0][0];
	
	// googleカレンダーの取得
	var calender = CalendarApp.getCalendarById(gAccount);
	
	// 予定の最終行を取得
	var lastRow = sheet.getLastRow();
	
	//予定の一覧を取得
	var contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues();
	
	//順に予定を作成（今回は正しい値が来ることを想定）
	for (i = 0; i <= lastRow - topRow; i++) {
		
		//「済」っぽいのか、空の場合は飛ばす
		var status = contents[i][statusCellNum];
		if (
			status == "済" ||
			status == "済み" ||
			status == "OK" ||
			contents[i][dayCellNum] == ""
		) {
			continue;
		}
		
		// 値をセット 日時はフォーマットして保持
		var day = new Date(contents[i][dayCellNum]);
		var endday = contents[i][enddayCellNum];
		var startTime = contents[i][startCellNum];
		var endTime = contents[i][endCellNum];
		var title = contents[i][titleCellNum];
		var recurrenceMethodIdx = contents[i][recurrenceMethodCellNum];
		var color = contents[i][colorCellNum];
		// 場所と詳細とゲストをセット
		var options = {
			location: contents[i][locationCellNum],
			description: contents[i][descriptionCellNum],
			guests: contents[i][guestsCellNum]
		};
		
		try {
			// 繰り返し関数が未指定であれば単発のイベントを作成
			if (recurrenceMethodIdx == '') {
				// 開始終了時間が無ければ終日で設定
				if ((startTime == '' || endTime == '') && endday == '') {
					//予定を作成
					var event = calender.createAllDayEvent(
						title,
						new Date(day),
						options
					);
					//終了日のみ入力すれば終日連続
				} else if (startTime == '' || endTime == '') {
					var endday = endday.setDate(endday.getDate() + 1);
					//予定を作成
					var event = calender.createAllDayEvent(
						title,
						new Date(day),
						new Date(endday),
						options,
					);
					// 開始終了時間があれば範囲で設定
				} else {
					// 開始日時をフォーマット
					var startDate = new Date(day);
					startDate.setHours(startTime.getHours());
					startDate.setMinutes(startTime.getMinutes());
					// 終了日時をフォーマット
					var endDate = new Date(day);
					endDate.setHours(endTime.getHours());
					endDate.setMinutes(endTime.getMinutes());
					// 予定を作成
					var event = calender.createEvent(
						title,
						startDate,
						endDate,
						options
					);
				}
			// 繰り返し関数が指定されていればシリーズで作成
			} else {
				// 開始終了時間が無ければ終日で設定
				if ((startTime == '' || endTime == '') && endday == '') {
					// 予定を作成
					var evalstr = "CalendarApp.newRecurrence()." + recurrenceMethod[recurrenceMethodIdx]
					var event = calender.createAllDayEventSeries(
						title,
						new Date(day),
						// 繰り返し関数をevalして繰り返しを作成.最終日を指定
						eval(evalstr),
						options
					);
				//終了日のみ入力すれば終日連続
				} else if (startTime == '' || endTime == '') {
					// 予定を作成
					var evalstr = "CalendarApp.newRecurrence()." + recurrenceMethod[recurrenceMethodIdx]
					var event = calender.createAllDayEventSeries(
						title,
						new Date(day),
						// 繰り返し関数をevalして繰り返しを作成.最終日を指定
						eval(evalstr).until(new Date(endday)),
						options
					);
				// 開始終了時間があれば範囲で設定
				} else {
					// 開始日時をフォーマット
					var startDate = new Date(day);
					startDate.setHours(startTime.getHours());
					startDate.setMinutes(startTime.getMinutes());
					// 終了日時をフォーマット
					var endDate = new Date(day);
					endDate.setHours(endTime.getHours());
					endDate.setMinutes(endTime.getMinutes());
					// 予定を作成
					var evalstr = "CalendarApp.newRecurrence()." + recurrenceMethod[recurrenceMethodIdx]
					var event = calender.createEventSeries(
						title,
						startDate,
						endDate,
						// 繰り返しを作成
						eval(evalstr),
						options
					);
				}
			}
			// 作成したイベントに対しての関数の適用
			// ゲストの編集権限追加
			event.setGuestsCanModify(true);
			// 色の指定
			if (color != '') {
				event.setColor(color);
			}
			
			// 無事に予定が作成されたら「済」にする
			sheet.getRange(topRow + i, 2).setValue("済");
			
			// エラーの場合（今回はログ出力のみ）
		} catch (e) {
			Logger.log(e);
		}
		
	}
	// ブラウザへ完了通知
	Browser.msgBox("完了");
}

