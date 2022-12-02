Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("btn_save").onclick = save;
  }
});

export async function save() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem("Settings");
      if (!sheet){
        let sheets = context.workbook.worksheets;
        sheet = sheets.add("Settings");
      }
      
      let headers = [
        ["Jira Host", "Jira User", "Jira Token"]
      ];
      let headerRange = sheet.getRange("A1:A3");
      headerRange.values = headers;
      headerRange.format.fill.color = "#4472C4";
      headerRange.format.font.color = "white";

      let jiraHostRange = sheet.getRange("B1");
      jiraHostRange.format.protection.locked = true;
      jiraHostRange.value = document.getElementById("txt_host").value;

      let jiraUserRange = sheet.getRange("B2");
      jiraUserRange.format.protection.locked = true;
      jiraUserRange.value = document.getElementById("txt_user").value;

      let jiraTokenRange = sheet.getRange("B3");
      jiraTokenRange.format.protection.locked = true;
      jiraHostRange.format.fill.color = "white";
      jiraHostRange.format.font.color = "white";
      jiraTokenRange.value = document.getElementById("txt_token").value;

      sheet.visibility = Excel.SheetVisibility.hidden;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
