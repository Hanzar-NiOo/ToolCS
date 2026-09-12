function exportSheetAsXlsx(spreadsheetId)
{
  var url = "https://www.googleapis.com/drive/v3/files/" + spreadsheetId
    + "/export?mimeType=" + encodeURIComponent(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token }
  });
  return response.getBlob();
}

function extractDayMonth(isoDateString) {
  var date = new Date(isoDateString + "T00:00:00Z");
  return Utilities.formatDate(date, "UTC", "ddMMM");
}

function formatNumber(value)
{
  const num = Number(value);
  if (isNaN(num))
  {
    return 0;
  }
  return num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatCount(value)
{
  const num = Number(value);
  if (isNaN(num))
  {
    return 0;
  }
  var count = num.toLocaleString('en-US', {
    maximumFractionDigits: 0
  });
  return count;
}
