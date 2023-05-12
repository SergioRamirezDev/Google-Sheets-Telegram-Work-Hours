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
    const bot = new TelegramBot(token, { polling: true });

    async function loadSavedCredentialsIfExist() {
        try {
            const content = await fs.readFile(TOKEN_PATH);
            const credentials = JSON.parse(content);
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
        if (client) {
            return client;
        }
        client = await authenticate({
            scopes: SCOPES,
            keyfilePath: CREDENTIALS_PATH,
        });
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
                                    newSheetName: date,
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
            var workdays = [["C", "D", "E", "F"], ["H", "I", "J", "K"], ["M", "N", "O", "P"], ["R", "S", "T", "U"], ["W", "X", "Y", "Z"], ["AB", "AC", "AD", "AE"], ["AG", "AH", "AI", "AJ"]]
            const dayoftheweek = moment().day() - 1;
            const commands = ["/entrada", "/salida", "/horadecomida", "/descanso"];
            const commandDescriptions = ["entrada", "salida", "comida", "descanso"];
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

    /*
        Commands
        entrada - Entrada
        salida - Salida
        horadecomida - Hora de comida
        descanso - Descanso
    */

    bot.onText(/\/echo (.+)/, (msg, match) => {
        const chatId = msg.chat.id;
        const resp = match[1];
        bot.sendMessage(chatId, resp);
    });

    bot.on('message', async (msg) => {
        try {
            var log = await fs.readFile('log.json');
            log = JSON.parse(log);
            log.push(msg);
            fs.writeFile("log.json", JSON.stringify(log));
        } catch (error) {
            console.log(error.message)
        }

        switch (msg.text) {
            case "/start":
                var st = await bot.sendMessage(msg.chat.id, `Bienvenid@ ${msg.from.first_name} ${msg.from.last_name}.`);
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/entrada":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/salida":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/horadecomida":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            case "/descanso":
                authorize().then((auth) => createSpreadsheet(auth, msg, bot)).catch(console.error);
                break;
            default:
                var st = await bot.sendMessage(msg.chat.id, `Porfavor selecciona desde el menu.`);
                var ph = await bot.sendPhoto(msg.chat.id, "./menu.jpg")
                /*setTimeout(() => {
                    bot.deleteMessage(msg.chat.id, st.message_id);
                    bot.deleteMessage(msg.chat.id, ph.message_id)
                }, 5000);
                bot.deleteMessage(msg.chat.id, msg.message_id)*/
                break;
        }
    });

}());