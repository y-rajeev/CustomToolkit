function pullDataFromERPNext() {
  var loginUrl = PropertiesService.getScriptProperties().getProperties('loginUrl');
  var apiUrl = PropertiesService.getScriptProperties().getProperties('apiUrl');
  var usr = PropertiesService.getScriptProperties().getProperty('usr');
  var pwd = PropertiesService.getScriptProperties().getProperty('pwd');

  var loginResponse = UrlFetchApp.fetch(loginUrl, {
    method: 'post',
    payload: {
      usr: usr,
      pwd: pwd
    },
    followRedirects: false
  });

  var cookies = loginResponse.getAllHeaders()['Set-Cookie'];
  var sessionCookie = cookies[0].split(';')[0];  // e.g., sid=xyz

  var apiResponse = UrlFetchApp.fetch(apiUrl, {
    method: 'get',
    headers: {
      'Cookie': sessionCookie
    }
  });

  var result = JSON.parse(apiResponse.getContentText()).message;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FNSKU Mapping");
  sheet.getRange('B:H').clearContent();

  if (result.length > 0) {
    var headers = Object.keys(result[0]);
    var data = result.map(row => headers.map(h => row[h]));

    // Write headers to row 1 starting from column B (i.e., column 2)
    sheet.getRange(1, 2, 1, headers.length).setValues([headers]);

    // Write data starting from row 2, column B
    sheet.getRange(2, 2, data.length, headers.length).setValues(data);
  }
}
