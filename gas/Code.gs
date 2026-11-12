
/**
 * MISHKAH v5.9 - vMix Animation Fix Final (Atomic Transaction)
 * - تم توحيد عمليات الكتابة في GameControl لتتم في عملية واحدة فقط (setValues).
 * - هذا التغيير يضمن أن vMix يرى التغيير فوراً كحدث واحد، مما يفعل DataChange Animation.
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  // vMix Direct Call Support
  if (!e || !e.parameter || !e.parameter.action) {
    return handleUpdateScores(true); 
  }

  const action = e.parameter.action;
  
  ensureGameControlSheet(); 
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const statusSheet = ss.getSheetByName('Status');
  const questionsSheet = ss.getSheetByName('Questions');
  const controlSheet = ss.getSheetByName('GameControl');

  if (!statusSheet || !questionsSheet) {
     return ContentService.createTextOutput(JSON.stringify({success:false, error:"Required sheets missing"})).setMimeType(ContentService.MimeType.JSON);
  }

  // تهيئة الجدول القديم Status Sheet
  if (statusSheet.getLastRow() === 0) {
    statusSheet.appendRow(['QuestionID', 'VotingOpen', 'ShowAnswer', 'P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'TimerEnd', 'GameStarted', 'CorrectAnswer', 'ActiveSegment', 'LockID']);
    statusSheet.appendRow([1, 'FALSE', 'FALSE', 0, 0, 0, 0, 0, 0, 0, 0, '0', 'FALSE', '', '1', '']); 
    while (statusSheet.getLastRow() < 5) statusSheet.appendRow(['', '', '', 0, 0, 0, 0, 0, 0, 0, 0]);
  }

  const now = new Date().getTime();
  let result = { success: true, serverTime: now };

  try {
    // التعامل مع الأوامر الخاصة بـ GameControl بطريقة مجمعة (Atomic)
    if (['addSeg2', 'addSeg3', 'addSeg4', 'addSeg5', 'resetSegment', 'setView', 'syncSeg1', 'toggleSelect', 'commitSelection'].indexOf(action) !== -1) {
       result = handleGameControlAtomic(action, e.parameter, statusSheet, controlSheet);
    } else {
      // الأوامر الأخرى (Quiz Logic)
      switch(action) {
        case 'setActiveSegment':
          const segVal = Number(e.parameter.segment) || 1;
          
          // 1. تحديث ورقة Status (للمنطق البرمجي)
          statusSheet.getRange("O2").setValue(segVal);
          
          // 2. تحديث ورقة Results (لتغيير صور العناوين والهيدر)
          const resultsSheet = ss.getSheetByName('Results');
          if (resultsSheet) {
            var apCol = getColIndexByName(resultsSheet, "ActivePart");
            if (apCol > 0) {
              resultsSheet.getRange(2, apCol).setValue(segVal);
            } else {
              resultsSheet.getRange("K2").setValue(segVal);
            }
          }

          result.msg = "Active Segment Set to: " + segVal;
          break;

        case 'getStatus':
          if (statusSheet.getLastRow() < 5) while (statusSheet.getLastRow() < 5) statusSheet.appendRow(['', '', '', 0, 0, 0, 0, 0, 0, 0, 0]);
          let statusData = statusSheet.getRange(2, 1, 1, 13).getValues()[0];
          let qId = parseInt(statusData[0]) || 1;
          let isVotingOpen = String(statusData[1]).toUpperCase() === 'TRUE';
          let timerEnd = Number(statusData[11]) || 0;
          let gameStarted = String(statusData[12] || "").toUpperCase() === 'TRUE';
          let activeSegment = statusSheet.getRange("O2").getValue();
          if (activeSegment === "") activeSegment = 1;

          if (isVotingOpen && timerEnd > 0 && now >= timerEnd) {
            // Timer logic removed - voting stays open until manual toggle
          }
          
          const lastRow = Math.max(questionsSheet.getLastRow(), 1);
          const totalQuestions = Math.max(0, lastRow - 1);
          let questionObj = null;
          let correctAnswer = null;
          let questionRowIndex = qId + 1;
          if (questionRowIndex <= lastRow && questionRowIndex > 1) {
              let qsMaxCols = questionsSheet.getMaxColumns();
              if (qsMaxCols >= 1) {
                  let readCols = Math.min(qsMaxCols, 6);
                  let rowData = questionsSheet.getRange(questionRowIndex, 1, 1, readCols).getValues()[0];
                  while(rowData.length < 6) rowData.push(""); 
                  correctAnswer = String(rowData[5] || "").trim().toUpperCase();
                  questionObj = {
                    id: qId,
                    text: rowData[0], a: rowData[1], b: rowData[2], c: rowData[3], d: rowData[4], 
                    correct: correctAnswer
                  };
              }
          }
          if (correctAnswer) statusSheet.getRange("N2").setValue(correctAnswer);
          
          let grandTotals = [];
          // قراءة المجاميع من GameControl Row 8
          if (controlSheet && controlSheet.getLastRow() >= 8) {
            grandTotals = controlSheet.getRange("B8:I8").getValues()[0];
          }
          grandTotals = grandTotals.map(x => Number(x) || 0);

          // قراءة بيانات الاختيار المسبق (Blue/Green)
          let selectionData = { blue: [], green: [] };
          if (controlSheet && controlSheet.getLastRow() >= 22) {
            selectionData.blue = controlSheet.getRange("B21:I21").getValues()[0].map(x => Number(x) || 0);
            selectionData.green = controlSheet.getRange("B22:I22").getValues()[0].map(x => Number(x) || 0);
          }

          result.data = {
            status: {
              currentQuestionId: qId,
              totalQuestions: totalQuestions,
              isVotingOpen: isVotingOpen,
              playerAnswers: statusData.slice(3, 11).map(a => String(a || "").trim().toUpperCase()),
              timerEnd: timerEnd,
              correctAnswer: correctAnswer,
              gameStarted: gameStarted,
              activeSegment: Number(activeSegment),
              scores: statusSheet.getRange("D5:K5").getValues()[0], 
              grandTotals: grandTotals,
              selection: selectionData
            },
            question: questionObj
          };
          break;

        case 'submitAnswer':
          const pId_sub = parseInt(e.parameter.playerId);
          const ans = String(e.parameter.answer).trim().toUpperCase();
          const curStatus = statusSheet.getRange(2, 1, 1, 13).getValues()[0];
          if (String(curStatus[12]).toUpperCase() === 'TRUE' && String(curStatus[1]).toUpperCase() === 'TRUE') {
            const targetCell = statusSheet.getRange(2, 3 + pId_sub);
            if (targetCell.getValue() === "") targetCell.setValue(ans);
          }
          break;

        case 'toggleVoting':
          const vState = String(statusSheet.getRange(2, 2).getValue()).toUpperCase();
          if (vState !== 'TRUE') {
            statusSheet.getRange(2, 2).setValue('TRUE');
            statusSheet.getRange(2, 12).setValue(0); // No timer end set
          } else {
            statusSheet.getRange(2, 2).setValue('FALSE');
            statusSheet.getRange(2, 12).setValue(0);
          }
          break;
          
        case 'setGameStatus':
          const started = String(e.parameter.started).toUpperCase();
          statusSheet.getRange(2, 13).setValue(started);
          if (started === 'FALSE') {
            statusSheet.getRange(2, 2).setValue('FALSE');
            statusSheet.getRange(2, 12).setValue(0);
          }
          break;

        case 'nextQuestion':
          moveQ(statusSheet, 1);
          break;
        case 'prevQuestion':
          moveQ(statusSheet, -1);
          break;

        case 'updateScores':
          const scoresResult = handleUpdateScores(false);
          result.message = scoresResult.getContent();
          // بعد تحديث النقاط، نقوم بمزامنة Seg1 مع لوحة التحكم
          handleGameControlAtomic('syncSeg1', {}, statusSheet, controlSheet);
          break;
      }
    }
  } catch (err) {
    result.success = false;
    result.error = err.toString();
  }
  
  // CRITICAL: Force flush
  SpreadsheetApp.flush();
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================================
// CORE LOGIC: ATOMIC UPDATES FOR GAMECONTROL (vMix Fix)
// ==========================================================
function handleGameControlAtomic(action, params, statusSheet, controlSheet) {
  // توسيع النطاق ليشمل الصفوف 21 و 22
  var range = controlSheet.getRange("B2:I22");
  var data = range.getValues(); 
  
  // تنظيف البيانات
  for(var r=0; r<data.length; r++) {
    for(var c=0; c<8; c++) {
      data[r][c] = Number(data[r][c]) || 0;
    }
  }

  var msg = "OK";
  var pIdx = (parseInt(params.playerId) || 1) - 1; 

  switch(action) {
    case 'toggleSelect':
      data[19][pIdx] = (data[19][pIdx] === 1) ? 0 : 1;
      msg = "P" + (pIdx+1) + " Toggled in Blue";
      break;

    case 'commitSelection':
      for(var i=0; i<8; i++) {
        data[20][i] = data[19][i]; 
      }
      msg = "Committed Blue to Green";
      break;

    case 'syncSeg1':
      var s1 = statusSheet.getRange("D5:K5").getValues()[0];
      for(var i=0; i<8; i++) data[0][i] = Number(s1[i]) || 0;
      data = updateLiveRowInMemory(data, 'seg1');
      break;
      
    case 'addSeg2':
      data[1][pIdx] += 20;
      data = updateLiveRowInMemory(data, 'seg2');
      break;
      
    case 'addSeg3':
      data[2][pIdx] += 10;
      data[3][pIdx] += 1;
      if (data[3][pIdx] === 7) data[4][pIdx] = 5;
      data = updateLiveRowInMemory(data, 'seg3');
      break;
      
    case 'addSeg4':
      data[5][pIdx] += 25;
      data = updateLiveRowInMemory(data, 'seg4');
      break;
      
    case 'addSeg5':
      var pts = parseInt(params.points) || 0;
      data[8][pIdx] += pts;
      data = updateLiveRowInMemory(data, 'seg5');
      break;
      
    case 'resetSegment':
      var seg = params.segment;
      for(var i=0; i<8; i++) {
        if(seg=='2') data[1][i] = 0;
        if(seg=='3') { data[2][i]=0; data[3][i]=0; data[4][i]=0; }
        if(seg=='4') data[5][i] = 0;
        if(seg=='5') data[8][i] = 0;
      }
      break;
      
    case 'setView':
      data = updateLiveRowInMemory(data, params.view);
      break;
  }

  // إعادة حساب المجموع
  for(var c=0; c<8; c++) {
    data[6][c] = data[0][c] + data[1][c] + data[2][c] + data[4][c] + data[5][c] + data[8][c];
  }
  
  range.setValues(data);
  range.setNumberFormat("0");
  
  return { success: true, msg: msg };
}

function updateLiveRowInMemory(data, viewMode) {
  var sourceIdx = 0;
  if (viewMode == 'seg1') sourceIdx = 0;
  else if (viewMode == 'seg2') sourceIdx = 1;
  else if (viewMode == 'seg3') sourceIdx = 2; 
  else if (viewMode == 'seg4') sourceIdx = 5;
  else if (viewMode == 'total') sourceIdx = 6;
  else if (viewMode == 'seg5') sourceIdx = 8;
  
  for(var i=0; i<8; i++) {
    var val = data[sourceIdx][i];
    if (viewMode == 'seg3') {
       val += data[4][i]; 
    }
    data[10][i] = val; 
  }
  return data;
}

function moveQ(sheet, d) {
  const currentId = parseInt(sheet.getRange(2, 1).getValue()) || 1;
  const newId = Math.max(1, currentId + d);
  sheet.getRange(2, 1).setValue(newId);
  sheet.getRange(2, 2).setValue('TRUE'); // التصويت يفتح تلقائياً
  sheet.getRange(2, 12).setValue(0);
  sheet.getRange(2, 4, 1, 8).clearContent();
  sheet.getRange("N2").clearContent();
  sheet.getRange("P1").clearContent();
}

function handleUpdateScores(isTextOutput) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Status");
  var questionsSheet = ss.getSheetByName("Questions");

  if (!sheet || !questionsSheet) {
      const msg = "Error: Sheets not found";
      return isTextOutput ? ContentService.createTextOutput(msg) : { getContent: () => msg };
  }

  var currentQID = sheet.getRange("A2").getValue();
  var lastRow = questionsSheet.getLastRow();
  if (currentQID && (currentQID + 1) <= lastRow) {
      var realCorrectAns = questionsSheet.getRange(currentQID + 1, 6).getValue();
      sheet.getRange("N2").setValue(realCorrectAns);
  }
  
  var lockCellValue = String(sheet.getRange("P1").getValue());
  var lockParts = lockCellValue.split("|");
  var lockedQuestionID = lockParts[0];
  var creditedPlayersIndices = [];

  if (lockedQuestionID == currentQID && lockParts.length > 1 && lockParts[1] !== "") {
    creditedPlayersIndices = lockParts[1].split(",");
  } else if (lockedQuestionID != currentQID) {
    creditedPlayersIndices = [];
  }

  var correctAns = String(sheet.getRange("N2").getValue()).trim().toUpperCase();
  var playerAnswers = sheet.getRange("D2:K2").getValues()[0];
  
  if (sheet.getLastRow() < 5) sheet.appendRow(['Scores:', '', '', 0, 0, 0, 0, 0, 0, 0, 0]);
  var scoreRange = sheet.getRange("D5:K5");
  var currentScores = scoreRange.getValues()[0];
  var newScores = [];
  var changesMade = false;

  for (var i = 0; i < 8; i++) {
    var pAns = String(playerAnswers[i] || "").trim().toUpperCase();
    var currentScore = Number(currentScores[i]) || 0;
    var playerIndexStr = String(i);

    if (pAns === correctAns && correctAns !== "") {
      if (creditedPlayersIndices.indexOf(playerIndexStr) === -1) {
        newScores.push(currentScore + 10);
        creditedPlayersIndices.push(playerIndexStr); 
        changesMade = true;
      } else {
        newScores.push(currentScore);
      }
    } else {
      newScores.push(currentScore);
    }
  }

  var resultMsg = "No new points added";
  if (changesMade) {
    scoreRange.setValues([newScores]);
    scoreRange.setNumberFormat("0");
    sheet.getRange("P1").setValue(currentQID + "|" + creditedPlayersIndices.join(","));
    resultMsg = "Scores Updated (New Winners Added)";
  }
  return isTextOutput ? ContentService.createTextOutput(resultMsg) : { getContent: () => resultMsg };
}

function ensureGameControlSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('GameControl');
  if (!sheet) {
    sheet = ss.insertSheet('GameControl');
    sheet.getRange("A1").setValue("Segments / Players");
    for(var i=0; i<8; i++) sheet.getRange(1, i+2).setValue("P"+(i+1));
    sheet.getRange("A2").setValue("Segment 1 (اختر صح)");
    sheet.getRange("A3").setValue("Segment 2 (آيات بينات 20pts)");
    sheet.getRange("A4").setValue("Segment 3 (اختر بابك 10pts)");
    sheet.getRange("A5").setValue("Seg 3 Counter");
    sheet.getRange("A6").setValue("Seg 3 Bonus (+5)");
    sheet.getRange("A7").setValue("Segment 4 (التحدي البصري 25pts)");
    sheet.getRange("A8").setValue("TOTAL SCORE").setFontWeight("bold").setBackground("#ffff00");
    sheet.getRange("A10").setValue("Segment 5 (من أكون ؟)");
    sheet.getRange("B8:I8").setBackground("#e6f7ff").setFontWeight("bold");
    sheet.getRange("A12").setValue("LIVE SCREEN").setFontWeight("bold").setBackground("#00ff00");
    sheet.getRange("B2:I12").setValue(0).setNumberFormat("0");
  }
}

function getColIndexByName(sheet, name) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().toLowerCase() == name.toString().toLowerCase()) return i + 1;
  }
  return -1;
}
