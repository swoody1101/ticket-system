/**
 * 주어진 시트에서 Edit URL을 기준으로 기존 응답을 찾습니다.
 * @param {Sheet} sheet 탐색할 Google Sheets 시트 객체
 * @param {string} editResponseUrl 폼 응답의 수정 URL
 * @returns {object|null} existingCaseId와 existingRowIndex를 포함하는 객체 또는 null
 */
function findExistingCase(sheet, editResponseUrl) {
  if (!sheet) {
    return null;
  }

  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var editUrlColumnIndex = header.indexOf("Edit URL");
  var caseIdColumnIndex = header.indexOf("케이스 ID");
  var completionStatusIndex = header.indexOf("문제 해결 여부");
  var mainCategoryIndex = header.indexOf("문제 원인 유형 (대분류)");
  var subCategoryIndex = header.indexOf("문제 원인 유형 (소분류)");
  var issuePriorityIndex = header.indexOf("문제 중요도");
  var issueDateTimeIndex = header.indexOf("문제 발생 일시");

  console.log("findExistingCase 내부: editURLidx : " + editUrlColumnIndex);

  if (editUrlColumnIndex === -1) {
    return null;
  }

  var sheetData = sheet.getDataRange().getValues();
  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][editUrlColumnIndex] === editResponseUrl) {
      console.log(
        "findExistingCase 내부 editURL : " + sheetData[i][editUrlColumnIndex]
      );
      return {
        caseId: sheetData[i][caseIdColumnIndex],
        completionStatus: sheetData[i][completionStatusIndex],
        mainCategory: sheetData[i][mainCategoryIndex],
        subCategory: sheetData[i][subCategoryIndex],
        issuePriority: sheetData[i][issuePriorityIndex],
        issueDateTime: sheetData[i][issueDateTimeIndex],
        rowIndex: i + 1,
      };
    }
  }

  return null;
}

/**
 * 문제 중요도에 대한 폴더명을 반환합니다.
 * @param {string} problemPriority 중요도 카테고리
 * @returns {string} 중요도 폴더명
 */
function getPriorityCategoryFolderName(problemPriority) {
  switch (problemPriority) {
    case "Critical":
      return "01_Critical";
    case "High":
      return "02_High";
    case "Medium":
      return "03_Medium";
    case "Low":
      return "04_Low";
  }
}

/**
 * 문제 원인 유형 대분류에 대한 폴더명을 반환합니다.
 * @param {string} category 대분류 카테고리
 * @returns {string} 대분류 폴더명
 */
function getMainCategoryFolderName(category) {
  switch (category) {
    case "JLK":
      return "01_RB_JLK";
    case "Hospital":
      return "02_RB_Hospital";
    case "PACS":
      return "03_RB_PACS";
    case "Other":
      return "04_RB_Other";
    default:
      return "04_RB_Other";
  }
}

/**
 * 문제 원인 유형 소분류에 대한 폴더명을 반환합니다.
 * @param {string} subCategory 소분류 카테고리
 * @returns {string} 소분류 폴더명
 */
function getSubCategoryFolderName(subCategory) {
  switch (subCategory) {
    case "Software":
      return "01_Software";
    case "Hardware":
      return "02_Hardware";
    case "Infrastructure":
      return "03_Infra";
    case "Human":
      return "04_Human";
    default:
      return "04_Human";
  }
}

/**
 * 문제 원인 유형 대분류에 대한 접두사를 반환합니다.
 * @param {string} category 대분류 카테고리
 * @returns {string} 대분류 접두사
 */
function getMainCategoryPrefix(mainCategory) {
  switch (mainCategory) {
    case "JLK":
      return "JLK";
    case "Hospital":
      return "HS";
    case "PACS":
      return "PC";
    case "Other":
      return "OT";
    default:
      return "OT";
  }
}

/**
 * 문제 원인 유형 소분류에 대한 접두사를 반환합니다.
 * @param {string} subCategory 소분류 카테고리
 * @returns {string} 소분류 접두사
 */
function getSubCategoryPrefix(subCategory) {
  switch (subCategory) {
    case "Software":
      return "SW";
    case "Hardware":
      return "HW";
    case "Infrastructure":
      return "IF";
    case "Human":
      return "HM";
    default:
      return "HM";
  }
}

/**
 * 지정된 폴더 내에서 주어진 접두사를 가진 파일들의 최고 시퀀스 번호를 찾습니다.
 * @param {Folder} targetFolder 파일을 탐색할 폴더
 * @param {string} caseIdPrefix 시퀀스 번호를 찾을 파일명의 접두사
 * @returns {number} 찾은 최고 시퀀스 번호
 */
function findHighestSequenceNum(targetFolder, caseIdPrefix) {
  var highestSequenceNum = 0;
  var filesInFolder = targetFolder.getFilesByType("text/html");

  while (filesInFolder.hasNext()) {
    var file = filesInFolder.next();
    var fileName = file.getName();
    if (fileName.startsWith(caseIdPrefix) && fileName.endsWith(".html")) {
      var numPart = fileName.substring(
        fileName.lastIndexOf("-") + 1,
        fileName.lastIndexOf(".")
      );
      if (!isNaN(numPart)) {
        var currentNum = parseInt(numPart);
        if (currentNum > highestSequenceNum) {
          highestSequenceNum = currentNum;
        }
      }
    }
  }
  return highestSequenceNum;
}

/**
 * 문제 미해결 케이스를 지정된 폴더 내에서 새로운 케이스 ID를 생성합니다.
 * @param {Folder} targetFolder 케이스 ID를 생성할 폴더
 * @param {string} issueDatetime 메인 카테고리
 * @param {string} subCategory 서브 카테고리
 * @returns {string} 새로 생성된 케이스 ID
 */
function generateOpenCaseId(targetFolder, issueDatetime) {
  var dateObject = new Date(issueDatetime);
  var issueDate = Utilities.formatDate(dateObject, "Asia/Seoul", "yyMMdd");
  var caseIdPrefix = `TS-${issueDate}-`;
  var highestSequenceNum = findHighestSequenceNum(targetFolder, caseIdPrefix);
  var sequenceNum = highestSequenceNum + 1;
  return `${caseIdPrefix}${Utilities.formatString("%03d", sequenceNum)}`;
}

/**
 * 문제 해결 케이스를 지정된 폴더 내에서 새로운 케이스 ID를 생성합니다.
 * @param {Folder} targetFolder 케이스 ID를 생성할 폴더
 * @param {string} mainCategory 메인 카테고리
 * @param {string} subCategory 서브 카테고리
 * @returns {string} 새로 생성된 케이스 ID
 */
function generateClosedCaseId(targetFolder, mainCategory, subCategory) {
  var mainCategoryPrefix = getMainCategoryPrefix(mainCategory);
  var subCategoryPrefix = getSubCategoryPrefix(subCategory);

  var caseIdPrefix = `TS-${mainCategoryPrefix}-${subCategoryPrefix}-`;
  var highestSequenceNum = findHighestSequenceNum(targetFolder, caseIdPrefix);
  var sequenceNum = highestSequenceNum + 1;
  return `${caseIdPrefix}${Utilities.formatString("%03d", sequenceNum)}`;
}

/**
 * 웹 앱을 배포할 때 실행되는 함수입니다.
 * 이 함수는 ticketInfo.html 파일의 내용을 불러와 HTML 서비스 객체를 생성합니다.
 */
function doGet() {
  var htmlTemplate = HtmlService.createTemplateFromFile("ticketInfo");
  var htmlContent = htmlTemplate.evaluate();

  return htmlContent;
}
