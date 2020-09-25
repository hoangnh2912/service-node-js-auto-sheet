const fs = require("fs");
const readline = require("readline");
const { google } = require("googleapis");
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
const TOKEN_PATH = "token.json";
fs.readFile("./sheet.json", (err, content) => {
  if (err) return console.log("Error loading client secret file:", err);
  authorize(JSON.parse(content), listMajors);
});

function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
  });
  console.log("Authorize this app by visiting this url:", authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question("Enter the code from that page here: ", (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err)
        return console.error(
          "Error while trying to retrieve access token",
          err
        );
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log("Token stored to", TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

function listMajors(auth) {
  const spreadsheetId = "1YnoFsRn84XWRBcdXT8E6H_gYNIhwhJzqgJ-VSTPmZEI"; // công ty
  // const spreadsheetId = "1SmS_QaPSTugN0mSEC8iv3JsoWSrviYtRS4POeneu1N0"; // clone
  const sheets = google.sheets({ version: "v4", auth });
  let values = [
    ["Carrect", "Thêm tính năng cho phase 2"],
    ["GOTECHT", "Làm APP GOTECH TV"],
  ];
  const resource = {
    values,
  };

  const name = "Hải Hoàng";
  var month = new Date().getMonth() + 1;
  var month = month > 9 ? month : "0" + month;
  const year = new Date().getFullYear();

  for (let index = new Date().getDate(); index <= 31; index++) {
    const date = `${index > 9 ? index : "0" + index}/${month}/${year}`;
    sheets.spreadsheets.values.get(
      {
        spreadsheetId,
        range: `${date}!C1:C99`,
      },
      (err, res) => {
        if (err) return console.log("Date not exist");
        const rows = res.data.values;
        if (rows) {
          rows.map((elem, index) => {
            if (JSON.stringify(elem).includes(name)) {
              sheets.spreadsheets.values.update(
                {
                  spreadsheetId,
                  range: `${date}!D${index + 1}:E${index + 2}`,
                  resource,
                  valueInputOption: "USER_ENTERED",
                },
                (err, result) => {
                  if (err) {
                    console.log(err);
                  } else {
                    console.log(date + " cells updated.");
                  }
                }
              );
            }
          });
        } else {
          console.log("No data found.");
        }
      }
    );
  }
}
