function exportTableToCSV() {
  var csv = [];
  // CSVのヘッダーを設定
  csv.push(["科目群", "科目区分1", "科目区分2", "科目名", "取得年度", "学期", "単位", "成績", "GP"].join(","));

  // 成績表の行を選択
  var rows = document.querySelectorAll("tr.operationboxf");
  let currentGroup = "";    // 科目群
  let currentCategory1 = ""; // 科目区分1
  let currentCategory2 = ""; // 科目区分2

  // 学籍番号と名前を取得
  var studentInfoText = document.querySelector("td[colspan='2']").innerText.trim();
  var studentInfoParts = studentInfoText.match(/(\S+)\s+(\S+)\s+(\S+)さんの成績です。/);
  var studentId = studentInfoParts[1];
  var studentName = (studentInfoParts[2] + studentInfoParts[3]).replace(/\s+/g, '');
  var filename = `${studentId}_${studentName}_Grades.csv`;

  // テーブルの行をループしてデータを取得
  rows.forEach(row => {
    var cols = row.querySelectorAll("td");

    // 行にデータが少ない場合は、「科目種別」の行として識別
    if (cols.length < 6 || (cols[1].innerText.trim() === "" && cols[2].innerText.trim() === "")) {
      const subjectName = cols[0].innerText.trim();

      // 科目群の設定
      if (subjectName.startsWith("◎") && subjectName.endsWith("◎")) {
        currentGroup = subjectName.replace(/◎/g, "").trim();
        currentCategory1 = "";  // 新しい科目群が始まったら区分をリセット
        currentCategory2 = "";
      }
      // 科目区分1の設定
      else if (subjectName.startsWith("【") && subjectName.endsWith("】")) {
        currentCategory1 = subjectName.replace(/【|】/g, "").trim();
        currentCategory2 = "";  // 新しい区分1が始まったら区分2をリセット
      }
      // 科目区分2の設定
      else if (subjectName.startsWith("《") && subjectName.endsWith("》")) {
        currentCategory2 = subjectName.replace(/《|》/g, "").trim();
      }
      return;  // 次の行へ
    }

    // 各科目の行データを取得し、「科目群」「科目区分1」「科目区分2」を追加
    var rowData = [
      currentGroup,
      currentCategory1,
      currentCategory2,
      cols[0].innerText.trim(),
      cols[1].innerText.trim(),
      cols[2].innerText.trim(),
      cols[3].innerText.trim(),
      cols[4].innerText.trim(),
      cols[5].innerText.trim()
    ];
    csv.push(rowData.join(","));
  });

  // CSVファイルを生成
  // Excelで開いたときに文字化けしないようにUTF-8 BOMを付与
  var csvFile = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csv.join("\n")], {type: "text/csv"});
  var downloadLink = document.createElement("a");
  downloadLink.download = filename;
  downloadLink.href = window.URL.createObjectURL(csvFile);
  downloadLink.style.display = "none";
  document.body.appendChild(downloadLink);
  downloadLink.click();
  document.body.removeChild(downloadLink);
}

// メッセージを受信してCSVエクスポート関数を実行
chrome.runtime.onMessage.addListener(function(request, sender, sendResponse) {
  if (request.action === "exportCSV") {
    exportTableToCSV();
    sendResponse({result: "success"});
  }
});