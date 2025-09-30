const SPREADSHEET_ID = "1CkdsIKK0GS6MR6B0lKpW39_hyquuXCWIZJDglxt1Zac";
const OPEN_TICKET_ROOT_FOLDER_ID = "1HwXjWy30JMeeRbNi-BJOSUZn1ouD7Ty2";
const CLOSED_TICKET_ROOT_FOLDER_ID = "1HlgXSuFu0eyi6XY_RfL83Vj3G0TICPew";

function onFormSubmit(e) {
  var formResponses = e.response;
  var itemResponses = formResponses.getItemResponses();

  // "Trouble Shooting Casebook" 스프레드시트의 실제 ID를 사용
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var openSheet = spreadsheet.getSheetByName("미해결 티켓 시트");
  var closedSheet = spreadsheet.getSheetByName("해결 티켓 시트");

  // 폼 응답에서 필요한 정보를 추출
  var responseData = {};
  itemResponses.forEach(function (itemResponse) {
    var title = itemResponse.getItem().getTitle();
    var response = itemResponse.getResponse();
    responseData[title] = response;
  });
  var 문서제목 = responseData["문서 제목"] || "";
  var 문제문의자 = responseData["문제 문의자"] || "";
  var 문제해결담당자 = responseData["문제 해결 담당자"] || "";
  var 문제발생일시 = responseData["문제 발생 일시"] || "";
  var 문제발생환경 = responseData["문제 발생 환경"] || "";
  var 문제현상 = responseData["문제 현상"] || "";
  var 문제해결여부 = responseData["문제 해결 여부"] || "";
  var 문제중요도 = responseData["문제 중요도"] || "";
  var 문제원인대분류 = responseData["문제 원인 유형 (대분류)"] || "";
  var 문제원인소분류 = responseData["문제 원인 유형 (소분류)"] || "";
  var 문제상세원인 = responseData["문제 상세 원인"] || "";
  var 시도한조치내역 = responseData["시도한 조치 내역"] || "";
  var 문제해결방안 = responseData["문제 해결 방안"] || "";
  var 문제발생제품명및버전 =
    responseData["(선택) 문제 발생 제품명 및 버전"] || "";
  var 재발방지대책 = responseData["(선택) 재발 방지 대책"] || "";
  var 참고자료 = responseData["(선택) 참고 자료(링크)"] || "";
  var editResponseUrl = formResponses.getEditResponseUrl();

  // 기존 응답 여부 및 티켓 종료 여부 확인
  var existingCaseId = null;
  var existingCompletionStatus = null;
  var existingMainCategory = null;
  var existingSubCategory = null;
  var existingIssuePriority = null;
  var existingIssueDateTime = null;
  var existingRowIndex = -1;
  var existingSheet = null;

  var existingSheetResult = findExistingCase(openSheet, editResponseUrl);
  console.log(existingSheetResult);
  if (existingSheetResult) {
    existingCaseId = existingSheetResult.caseId;
    existingCompletionStatus = existingSheetResult.completionStatus;
    existingMainCategory = existingSheetResult.mainCategory;
    existingSubCategory = existingSheetResult.subCategory;
    existingIssuePriority = existingSheetResult.issuePriority;
    existingIssueDateTime = existingSheetResult.issueDateTime;
    existingRowIndex = existingSheetResult.rowIndex;
    existingSheet = openSheet;
  } else {
    existingSheetResult = findExistingCase(closedSheet, editResponseUrl);
    if (existingSheetResult) {
      existingCaseId = existingSheetResult.caseId;
      existingCompletionStatus = existingSheetResult.completionStatus;
      existingMainCategory = existingSheetResult.mainCategory;
      existingSubCategory = existingSheetResult.subCategory;
      existingIssuePriority = existingSheetResult.issuePriority;
      existingIssueDateTime = existingSheetResult.issueDateTime;
      existingRowIndex = existingSheetResult.rowIndex;
      existingSheet = closedSheet;
    }
  }

  console.log("existingSheetResult: " + existingSheetResult);

  // 결과 저장 폴더 정보
  var openTicketRootFolder = DriveApp.getFolderById(OPEN_TICKET_ROOT_FOLDER_ID);
  var closedTicketRootFolder = DriveApp.getFolderById(
    CLOSED_TICKET_ROOT_FOLDER_ID
  );
  var openTicketCategoryFolderName = getPriorityCategoryFolderName(문제중요도);
  var closedTicketMainCategoryFolderName =
    getMainCategoryFolderName(문제원인대분류);
  var closedTicketSubCategoryFolderName =
    getSubCategoryFolderName(문제원인소분류);

  // 티켓 해결 여부에 따라 구글 시트 및 HTML 저장 폴더 경로 설정
  var targetsheet = openSheet;
  var targetFolder = openTicketRootFolder;
  if (문제해결여부 === "미해결") {
    // Open Ticket
    var openTicketFolder = openTicketRootFolder.getFoldersByName(
      openTicketCategoryFolderName
    );
    if (openTicketFolder.hasNext()) {
      targetFolder = openTicketFolder.next();
    } else {
      targetFolder = openTicketRootFolder.createFolder(
        openTicketCategoryFolderName
      );
    }
    targetsheet = openSheet;
  } else {
    // Closed Ticket
    var closedTicketMainCategoryFolder =
      closedTicketRootFolder.getFoldersByName(
        closedTicketMainCategoryFolderName
      );
    if (closedTicketMainCategoryFolder.hasNext()) {
      targetFolder = closedTicketMainCategoryFolder.next();
    } else {
      targetFolder = closedTicketRootFolder.createFolder(
        closedTicketMainCategoryFolderName
      );
    }
    var closedTicketSubCategoryFolder = targetFolder.getFoldersByName(
      closedTicketSubCategoryFolderName
    );
    if (closedTicketSubCategoryFolder.hasNext()) {
      targetFolder = closedTicketSubCategoryFolder.next();
    } else {
      targetFolder = targetFolder.createFolder(
        closedTicketSubCategoryFolderName
      );
    }
    targetsheet = closedSheet;
  }

  // 케이스 ID 생성 및 수정
  var caseId = existingCaseId;
  // 기존 Ticket인 경우
  if (existingCaseId && existingCaseId.startsWith("TS-")) {
    // Open Ticket
    if (existingCompletionStatus === "미해결") {
      // 미해결 -> 해결
      if (문제해결여부 === "해결") {
        caseId = generateClosedCaseId(
          targetFolder,
          문제원인대분류,
          문제원인소분류
        );
      } else {
        // 미해결 -> 미해결
        // 카테고리(문제 중요도, 문제 발생 일시) 변경
        var existingIssueDateObject = new Date(existingIssueDateTime);
        var inputIssuDateOject = new Date(문제발생일시);
        var existingIssueDate = Utilities.formatDate(
          existingIssueDateObject,
          "Asia/Seoul",
          "yyMMdd"
        );
        var inputIssuDate = Utilities.formatDate(
          inputIssuDateOject,
          "Asia/Seoul",
          "yyMMdd"
        );
        if (
          existingIssueDate !== inputIssuDate ||
          existingIssuePriority !== 문제중요도
        ) {
          caseId = generateOpenCaseId(targetFolder, 문제발생일시);
        }
      }
      // Closed Ticket
    } else {
      // 해결 -> 미해결
      if (문제해결여부 === "미해결") {
        caseId = generateOpenCaseId(targetFolder, 문제발생일시);
      } else {
        // 해결 -> 해결
        // 카테고리(문제 원인 대분류, 문제 원인 소분류) 변경
        if (
          existingMainCategory != 문제원인대분류 ||
          existingSubCategory != 문제원인소분류
        ) {
          caseId = generateClosedCaseId(
            targetFolder,
            문제원인대분류,
            문제원인소분류
          );
        }
      }
    }
    // 신규 Ticket인 경우
  } else {
    console.log("신규 티켓!!!");
    console.log("문제 발생 일시: ");
    // Open Ticket
    if (문제해결여부 === "미해결") {
      console.log(문제발생일시);
      caseId = generateOpenCaseId(targetFolder, 문제발생일시);
      // Closed Ticket
    } else {
      caseId = generateClosedCaseId(
        targetFolder,
        문제원인대분류,
        문제원인소분류
      );
    }
  }

  // HTML 탬플릿 로드
  var htmlTemplate = HtmlService.createTemplateFromFile("ticketInfo");
  htmlTemplate.caseId = caseId;
  htmlTemplate.문서제목 = 문서제목;
  htmlTemplate.문제문의자 = 문제문의자;
  htmlTemplate.문제해결담당자 = 문제해결담당자;
  htmlTemplate.문제발생일시 = 문제발생일시;
  htmlTemplate.문제발생환경 = 문제발생환경;
  htmlTemplate.문제현상 = 문제현상;
  htmlTemplate.문제해결여부 = 문제해결여부;
  htmlTemplate.문제중요도 = 문제중요도;
  htmlTemplate.문제원인대분류 = 문제원인대분류;
  htmlTemplate.문제원인소분류 = 문제원인소분류;
  htmlTemplate.문제상세원인 = 문제상세원인;
  htmlTemplate.시도한조치내역 = 시도한조치내역;
  htmlTemplate.문제해결방안 = 문제해결방안;
  htmlTemplate.문제발생제품명및버전 = 문제발생제품명및버전;
  htmlTemplate.재발방지대책 = 재발방지대책;
  htmlTemplate.참고자료 = 참고자료;
  var htmlContent = htmlTemplate.evaluate().getContent();

  // 기존 파일 삭제
  console.log("삭제 전 caseId : " + existingCaseId);
  if (existingCaseId && existingCaseId.startsWith("TS-")) {
    var existingFiles = DriveApp.getFilesByName(`${existingCaseId}.html`);
    while (existingFiles.hasNext()) {
      var file = existingFiles.next();
      file.setTrashed(true); // 파일을 휴지통으로 이동
    }
  }

  // HTML 파일 생성
  var newFileName = `${caseId}.html`;
  targetFolder.createFile(newFileName, htmlContent, "text/html");

  // 자동으로 작성된 내용 업데이트
  var rowData = [
    caseId, // 맨 앞 컬럼에 케이스 ID 추가
    formResponses.getTimestamp(), // 티켓 작성 일시
    formResponses.getRespondentEmail(), // 이메일 주소 (응답자가 로그인한 경우)
    문서제목,
    문제문의자,
    문제해결담당자,
    문제발생일시,
    문제발생환경,
    문제현상,
    문제해결여부,
    문제중요도,
    문제원인대분류,
    문제원인소분류,
    문제상세원인,
    시도한조치내역,
    문제해결방안,
    문제발생제품명및버전,
    재발방지대책,
    참고자료,
    editResponseUrl,
  ];

  if (existingRowIndex !== -1) {
    if (existingSheet === targetsheet) {
      targetsheet
        .getRange(existingRowIndex, 1, 1, rowData.length)
        .setValues([rowData]);
    } else {
      existingSheet.deleteRow(existingRowIndex);
      targetsheet.appendRow(rowData);
    }
  } else {
    targetsheet.appendRow(rowData);
  }
}
