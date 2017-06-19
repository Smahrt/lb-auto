

var fs = require('fs');
var readline = require('readline');
var google = require('googleapis');
var googleAuth = require('google-auth-library');

// If modifying these scopes, delete your previously saved credentials
// at ~/.credentials/sheets.googleapis.com-nodejs-quickstart.json
var SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
var TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH ||
    process.env.USERPROFILE) + '/.credentials/';
var TOKEN_PATH = TOKEN_DIR + 'sheets.googleapis.com-nodejs-quickstart.json';

// Load client secrets from a local file.
fs.readFile('client_secret.json', function processClientSecrets(err, content) {
  if (err) {
    console.log('Error loading client secret file: ' + err);
    return;
  }
  // Authorize a client with the loaded credentials, then call the
  // Google Sheets API.
  authorize(JSON.parse(content), initApp);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  var clientSecret = credentials.installed.client_secret;
  var clientId = credentials.installed.client_id;
  var redirectUrl = credentials.installed.redirect_uris[0];
  var auth = new googleAuth();
  var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, function(err, token) {
    if (err) {
      getNewToken(oauth2Client, callback);
    } else {
      oauth2Client.credentials = JSON.parse(token);
      callback(oauth2Client);
    }
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
  var authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES
  });
  console.log('Authorize this app by visiting this url: ', authUrl);
  var rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  rl.question('Enter the code from that page here: ', function(code) {
    rl.close();
    oauth2Client.getToken(code, function(err, token) {
      if (err) {
        console.log('Error while trying to retrieve access token', err);
        return;
      }
      oauth2Client.credentials = token;
      storeToken(token);
      callback(oauth2Client);
    });
  });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
  try {
    fs.mkdirSync(TOKEN_DIR);
  } catch (err) {
    if (err.code != 'EEXIST') {
      throw err;
    }
  }
  fs.writeFile(TOKEN_PATH, JSON.stringify(token));
  console.log('Token stored to ' + TOKEN_PATH);
}

/** APP RUNS HERE **/
function initApp(auth) {
    var current_sheet = '1L6cJIgR-UhriPOn5LQlU-nyug0OK2WxNyQ4LVHG0dds';
    var sheet_name = getPeriod();
    var authe = auth;
    
    getSheetValues(authe, current_sheet, '!B3:AC', sheet_name).then(
        data => {
            var rows = data;
            
            createSheet(authe, sheet_name).then(id => {
                updateSheet(authe, id, rows, 'A1:AB');
            },
                                               err => {
                console.log(err);
            });
    },
        err => {
            console.log(err);
        });
        
}

function updateSheet(auth, id, values, range) {
    
  var sheets = google.sheets('v4');
        sheets.spreadsheets.values.update({
            auth: auth,
            spreadsheetId: id, //'1-vEWtF9i08jx6RuH14VXkWqhJAaI4DblqMgq1MuJQ6I'
            range: range,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: values
            }
        }, function(err, response) {
            if (err) {
              console.log('The API returned an error while updating: ' +err);
              return;
            }

            // TODO: Change code below to process the `response` object:
            console.log(JSON.stringify(response, null, 2));
          });
}

function createSheet(auth, period) {
    
  var sheets = google.sheets('v4');
    
    return new Promise((resolve, reject) => {
        
        sheets.spreadsheets.create({
            auth: auth,
            resource: {
                properties: {
                    title: "Mainone Link Budget Report for "+period
                },
                sheets: [
                    {
                        properties: {
                            title: period
                        }
                    }
                ]
            }
        }, function (err, response) {
            if (err) {
                console.log('The API returned an error while creating: ' +err);
                reject(err);
            }
            var newid = response.spreadsheetId;

            console.log(JSON.stringify(response, null, 2));

            resolve(newid);
        });
    });
}

function getSheetValues(auth, id, range, period) {
    
    var sheets = google.sheets('v4');
    var extracted = [];
    
    return new Promise((resolve, reject) => {
            sheets.spreadsheets.values.get({
        auth: auth,
        spreadsheetId: id,
        range: 'May 2017!B3:AC', // "'"+period+range+"'"
      }, function(err, response) {
        if (err) {
          reject(new Error('The API returned an error while getting: ' + err));
        }
        var rows = response.values;
        if (rows.length == 0) {
            reject(new Error('No data found.'))
        } else {
            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                // Print columns A and E, which correspond to indices 0 and 4.
                console.log('%s, %s', row[0], row[4]);
                extracted.push([row[0], row[4], row[14], row[17], row[16], row[19], row[20], row[21], row[22], row[23], row[27]]);
            }
            resolve(extracted);
        }
      });
    });
  
}

function getPeriod() {
    let month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    var d = new Date();
    let period = month[d.getMonth()]+" "+d.getFullYear();
    
    return period;
}