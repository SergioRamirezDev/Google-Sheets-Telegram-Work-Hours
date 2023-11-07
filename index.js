require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const token = process.env.TELEGRAMTOKEN;
const moment = require('moment');
const fs = require('fs').promises;
const path = require('path');
const proc = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = path.join(proc.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(proc.cwd(), 'credentials.json');
(async function () {
    var google_c = await fs.readFile(CREDENTIALS_PATH);
    console.log("CREDENTIALS_PATH", JSON.parse(google_c))
    var log = await fs.readFile('log.json');
    log = JSON.parse(log);
    const bot = new TelegramBot(token, {
        polling: true
    });

    async function loadSavedCredentialsIfExist() {
        try {
            console.log("loadSavedCredentialsIfExist", TOKEN_PATH)
            if (!fs.existsSync(TOKEN_PATH)) {
                return null;
            }
            const content = await fs.readFile(TOKEN_PATH);
            const credentials = JSON.parse(content);
            console.log("credentials", credentials)
            return google.auth.fromJSON(credentials);
        } catch (err) {
            return null;
        }
    }

    async function saveCredentials(client) {
        const content = await fs.readFile(CREDENTIALS_PATH);
        const keys = JSON.parse(content);
        const key = keys.installed || keys.web;
        const payload = JSON.stringify({
            type: 'authorized_user',
            client_id: key.client_id,
            client_secret: key.client_secret,
            refresh_token: client.credentials.refresh_token,
        });
        await fs.writeFile(TOKEN_PATH, payload);
    }

    async function authorize() {
        let client = await loadSavedCredentialsIfExist();
        console.log("client", client);
        if (client) {
            return client;
        }
        console.log("authenticate")
        client = await authenticate({
            scopes: SCOPES,
            keyfilePath: CREDENTIALS_PATH,
        }).catch(err => console.log(err));
        console.log('Tokens:', localAuth.credentials);
        if (client.credentials) {
            await saveCredentials(client);
        }
        return client;
    }

    async function createSpreadsheet(auth, msg, bot) {
        const gogoleauth = google.sheets({ version: 'v4', auth });
        const date = moment().format("DD/MM/YY");
        const week = moment().week();
        const workhoursheetid = process.env.WORKSHEETID;
        const themeid = process.env.TEMPLATEID;
        try {
            var spreadsheet = await gogoleauth.spreadsheets;
            var sheets = await spreadsheet.get({
                spreadsheetId: workhoursheetid
            });

            console.log(sheets)

            var sheetIndex = await sheets.data.sheets.findIndex(sh => sh.properties.sheetId == week);

            if (sheetIndex == -1) {

                var copytheme = await gogoleauth.spreadsheets.sheets.copyTo({
                    spreadsheetId: themeid,
                    sheetId: 0,
                    requestBody: {
                        destinationSpreadsheetId: workhoursheetid
                    }
                });

                await spreadsheet.batchUpdate({
                    spreadsheetId: workhoursheetid,
                    requestBody: {
                        requests: [
                            {
                                duplicateSheet: {
                                    newSheetId: week,
                                    newSheetName: `Week ${week}`,
                                    sourceSheetId: copytheme.data.sheetId
                                }
                            }, {
                                deleteSheet: {
                                    sheetId: copytheme.data.sheetId
                                }
                            }
                        ]
                    }
                });

                await spreadsheet.values.batchUpdate({
                    spreadsheetId: workhoursheetid,
                    valueInputOption: "USER_ENTERED",
                    requestBody: {
                        data: [{
                            range: `Week ${week}!K1`,
                            values: [[moment().day(1).format("DD/MM/YY")]],
                        }, {
                            range: `Week ${week}!T1`,
                            values: [[moment().day(2).format("DD/MM/YY")]],
                        }, {
                            range: `Week ${week}!AC1`,
                            values: [[moment().day(3).format("DD/MM/YY")]],
                        }, {
                            range: `Week ${week}!AL1`,
                            values: [[moment().day(4).format("DD/MM/YY")]],
                        }, {
                            range: `Week ${week}!AU1`,
                            values: [[moment().day(5).format("DD/MM/YY")]],
                        }, {
                            range: `Week ${week}!BD1`,
                            values: [[moment().day(6).format("DD/MM/YY")]],
                        }, {
                            range: `Week ${week}!BM1`,
                            values: [[moment().day(7).format("DD/MM/YY")]],
                        }]
                    }
                });

                sheets = await spreadsheet.get({
                    spreadsheetId: workhoursheetid
                });

                sheetIndex = await sheets.data.sheets.findIndex(sh => sh.properties.sheetId == week);
            }

            const actualsheet = sheets.data.sheets[sheetIndex];

            var worksheets = await spreadsheet.values.get({
                spreadsheetId: workhoursheetid,
                range: `${actualsheet.properties.title}!A4:A30`
            });

            var rows = worksheets.data.values;
            var startAddUsers = 4;
            var userExists = -1;
            if (rows && rows.length > 0) {
                startAddUsers = startAddUsers + rows.length;
                userExists = await rows.findIndex(row => {
                    return row.indexOf(`${msg.from.id}`) != -1
                });
            }

            if (userExists == -1) {
                await spreadsheet.values.update({
                    spreadsheetId: workhoursheetid,
                    valueInputOption: "USER_ENTERED",
                    range: `${actualsheet.properties.title}!A${startAddUsers}:B${startAddUsers}`,
                    requestBody: {
                        values: [[msg.from.id, `${msg.from.first_name} ${msg.from.last_name}`]]
                    }
                });
            }

            worksheets = await spreadsheet.values.get({
                spreadsheetId: workhoursheetid,
                range: `${actualsheet.properties.title}!A4:A30`
            });

            rows = worksheets.data.values;

            userExists = await rows.findIndex(row => {
                return row.indexOf(`${msg.from.id}`) != -1
            });

            var userRow = userExists + 4;
            var workdays = [["C", "D", "E", "F", "G", "H"], ["L", "M", "N", "O", "P", "Q"], ["U", "V", "W", "X", "Y", "Z"], ["AD", "AE", "AF", "AG", "AH", "AI"], ["AM", "AN", "AO", "AP", "AQ", "AR"], ["AV", "AW", "AX", "AY", "AZ", "BA"], ["BE", "BF", "BG", "BH", "BI", "BJ"]]
            const dayoftheweek = moment().day() - 1;
            const commands = ["/entrada", "/salida", "/comidasalida", "/comidaregreso", "/descansosalida", "/descansoregreso"];
            const commandDescriptions = ["entrada", "salida", "comidasalida", "comidaregreso", "descansosalida", "descansoregreso"];
            var indexOfCommand = await commands.findIndex(text => text == msg.text);
            if (indexOfCommand != -1) {
                worksheets = await spreadsheet.values.get({
                    spreadsheetId: workhoursheetid,
                    range: `${actualsheet.properties.title}!${workdays[dayoftheweek][indexOfCommand]}${userRow}`
                });

                rows = worksheets.data.values;

                if (!rows || (rows.length > 0 && rows[0][0] == "")) {
                    var actualdate = moment(msg.date * 1000);
                    var dt = moment(actualdate).format("MM/DD/YY HH:mm:ss");
                    var textdate = moment(actualdate).format('h:mm a');

                    console.log(`${userExists} - ${`${msg.from.first_name} ${msg.from.last_name}`} | ${actualsheet.properties.title}!${workdays[dayoftheweek][indexOfCommand]}${userRow} | ${dt}`)

                    await spreadsheet.values.update({
                        spreadsheetId: workhoursheetid,
                        valueInputOption: "USER_ENTERED",
                        range: `${actualsheet.properties.title}!${workdays[dayoftheweek][indexOfCommand]}${userRow}`,
                        requestBody: {
                            values: [[dt]]
                        }
                    });
                    await bot.sendMessage(msg.chat.id, `Hora de ${commandDescriptions[indexOfCommand]} actualizada a las ${textdate}.`);
                } else {
                    await bot.sendMessage(msg.chat.id, `Ya se ha usado esta opción anteriormente, por favor contacta a un administrador si necesitas hacer alguna modificación en el horario.`);
                }
            }

        } catch (error) {
            console.log(error.message)
        }
    }

    bot.onText(/\/echo (.+)/, (msg, match) => {
        const chatId = msg.chat.id;
        const resp = match[1];
        bot.sendMessage(chatId, resp);
    });

    bot.on('message', async (msg) => {
        console.log(msg)
        try {
            log.push(msg);
            fs.writeFile("log.json", JSON.stringify(log));
        } catch (error) {
            console.log(error.message)
        }

        switch (msg.text) {
            case "/start":
                await bot.sendMessage(msg.chat.id, `Bienvenid@ ${msg.from.first_name} ${msg.from.last_name}.`);
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/entrada":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/salida":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/comidasalida":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/comidaregreso":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/descansosalida":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/descansoregreso":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            default:
                await bot.sendMessage(msg.chat.id, `Porfavor selecciona desde el menu.`);
                await bot.sendPhoto(msg.chat.id, "./menu.jpg")
                break;
        }
    });

}());