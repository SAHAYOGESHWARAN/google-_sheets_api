const express = require('express');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const dotenv = require('dotenv');

dotenv.config();

const app = express();
const PORT = 3000;

// OAuth2 client setup
const credentialsPath = path.join(__dirname, 'credentials.json');
let credentials;

try {
    credentials = JSON.parse(fs.readFileSync(credentialsPath));
} catch (error) {
    console.error('Error reading credentials file:', error);
    process.exit(1);
}

const { installed } = credentials;
if (!installed) {
    console.error('Invalid credentials format: Missing "installed" object');
    process.exit(1);
}

const { client_secret, client_id, redirect_uris } = installed;
if (!client_secret || !client_id || !redirect_uris) {
    console.error('Invalid credentials format: Missing required fields');
    process.exit(1);
}

const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
);

// Session middleware
app.use(session({
    secret: 'YOUR_SESSION_SECRET',
    resave: false,
    saveUninitialized: true
}));

app.use(express.json()); // To parse JSON request bodies

// Routes
app.get('/', (req, res) => {
    res.send('<a href="/auth">Authorize with Google</a>');
});

app.get('/auth', (req, res) => {
    const url = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: ['https://www.googleapis.com/auth/spreadsheets']
    });
    res.redirect(url);
});

app.get('/oauth2callback', async (req, res) => {
    const code = req.query.code;
    try {
        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        req.session.tokens = tokens;
        res.send('Authentication successful! You can now access Google Sheets.');
    } catch (error) {
        console.error('Error exchanging code for tokens:', error.response ? error.response.data : error.message);
        res.status(500).send('Error exchanging code for tokens');
    }
});

app.get('/sheets', async (req, res) => {
    if (!req.session.tokens) {
        return res.redirect('/');
    }

    oAuth2Client.setCredentials(req.session.tokens);

    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: '12KLryfOauMIlzQXEv31gyD9QUymV9hrwibumotqs_Jc',
            range: 'Sheet1!A1:D10',
        });

        console.log('Data fetched from Google Sheets:', response.data.values); // Debugging step
        res.json(response.data.values);
    } catch (error) {
        console.error('Error fetching data from Google Sheets:', error.response ? error.response.data : error.message);
        res.status(500).send('Error fetching data from Google Sheets');
    }
});

app.post('/sheets/add', async (req, res) => {
    if (!req.session.tokens) {
        return res.redirect('/');
    }

    oAuth2Client.setCredentials(req.session.tokens);

    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const { range, values } = req.body;

    console.log('Adding data to sheet:', { range, values }); // Debugging step

    try {
        const response = await sheets.spreadsheets.values.append({
            spreadsheetId: '12KLryfOauMIlzQXEv31gyD9QUymV9hrwibumotqs_Jc',
            range: range,
            valueInputOption: 'RAW', // 'RAW' or 'USER_ENTERED'
            requestBody: {
                values: values,
            },
        });

        console.log('Data added successfully:', response.data); // Debugging step
        res.json(response.data);
    } catch (error) {
        console.error('Error adding data to Google Sheets:', error.response ? error.response.data : error.message);
        res.status(500).send('Error adding data to Google Sheets');
    }
});


app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
