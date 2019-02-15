//rangeの列のリスト
//0スタートなので、スプレッドシート上で利用するときは+1すること
var COL = {
	TITLE: 0,
	CAL: 1,
	ALLDAY: 2,
	START: 3,
	END: 4,
	LOOP: 5,
	LOOP_END: 6,
	DESCRIPTION: 7
};

var SS_ID = '1I6Joo8vVs3jcuCTdL0AtucB7edXD-PzbYja3ANLxpfU';

function onOpen(e) {
	Logger.log("onOpen");
	SpreadsheetApp.getUi().createAddonMenu()
		.addItem('起動', 'showSidebar')
		.addToUi();
}

function onInstall(e) {
	Logger.log("onInstall");
	onOpen(e);
}

function showSidebar() {
	Logger.log("showSidebar");
	var ui = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('ScheduleGenerator');
	SpreadsheetApp.getUi().showSidebar(ui);
}


/**
 * 予定入力用シートの追加
 *
 */
function addSheet() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	//現在のスプレッドシートにテンプレートをコピー
	var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('template').copyTo(ss);
	sheet.activate();
	updateCalendarList();
	return sheet.getName();
}

//カレンダー一覧を取得

/**
 * カレンダー一覧情報の更新
 *
 * @returns カレンダー一覧情報のMAP
 */
function updateCalendarList() {
	//自分のカレンダー一覧を取得
	var calendars = CalendarApp.getAllOwnedCalendars();
	var calendarList = {};
	for (var i in calendars) {
		//あとでカレンダー名からカレンダーobjを取得しやすいようにmapを作成
		calendarList[calendars[i].getName()] = calendars[i];
	}

	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheet = ss.getActiveSheet();
	var range = sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1);
	var calendarNames = Object.keys(calendarList);
	// カレンダー選択列に入力規則としてカレンダー名を指定
	var rule = SpreadsheetApp.newDataValidation().requireValueInList(calendarNames, true).setAllowInvalid(false).build();
	range.setDataValidation(rule);

	return calendarList;
}


/**
 * 予定の一括登録
 *
 */
function generateEvents() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheet = ss.getActiveSheet();

	var range = sheet.getRange(2, 2, sheet.getLastRow() - 1, sheet.getLastColumn() - 1);
	var values = range.getValues();

	var checkBox = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
	var results = checkBox.getValues();

	var result = 0;
	var calendarList = updateCalendarList();
	for (var i in values) {
		if (results[i][0] == false) {
			var value = values[i];
			var event = generateEvent(
				calendarList[value[COL.CAL]],
				value[COL.TITLE],
				value[COL.START],
				value[COL.END],
				value[COL.ALLDAY],
				value[COL.DESCRIPTION],
				value[COL.LOOP],
				value[COL.LOOP_END]
			);
			if (event != null) {
				results[i][0] = true;
				result.success++;
			} else {
				results[i][0] = false;
			}
		}
	}
	checkBox.setValues(results);
	return result;
}

function generateEvent(calendar, title, start, end, allDay, description, loop, loopEnd) {
	//タイトル・開始・終了は必須
	if (calendar == '' || title == '' || start == '' || end == '') {
		return null;
	}
	try {
		//var calendar = CalendarApp.getDefaultCalendar();
		//繰り返しの予定かどうかで分岐
		if (loop == '') {
			if (allDay == false) {
				//繰り返しなし・終日でない
				return calendar.createEvent(title, start, end, { 'description': description });
			} else if (allDay == true) {
				//繰り返しなし・終日
				return calendar.createAllDayEvent(title, start, end, { 'description': description });
			}
		} else if (loop != '') {
			//繰り返しルールの作成
			var recurrence = buildRecurrenceRule(loop, loopEnd);
			if (allDay == false) {
				//繰り返し・終日でない
				return calendar.createEventSeries(title, start, end, recurrence, { 'description': description });
			} else if (allDay == true) {
				//繰り返し・終日
				return calendar.createAllDayEventSeries(title, start, recurrence, { 'description': description });
			}
		}
	} catch (error) {
		return null;
	}
}

function buildRecurrenceRule(loop, end) {
	var recurrence = CalendarApp.newRecurrence();
	switch (loop) {
		case '毎日':
			recurrence = recurrence.addDailyRule();
			break;
		case '毎週':
			recurrence = recurrence.addWeeklyRule();
			break;
		case '毎月':
			recurrence = recurrence.addMonthlyRule();
			break;
		case '毎年':
			recurrence = recurrence.addYearlyRule();
			break;
	}
	recurrence.until(end);
	return recurrence;
}

function showModal(title, body, width, height) {
	var htmlOutput = HtmlService
		.createHtmlOutput(body)
		.setWidth(width)
		.setHeight(height);
	SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}
