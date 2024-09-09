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
const oAuth2Client = new google.auth.OAuth2(
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET,
    process.env.REDIRECT_URI
);

// Load credentials from file
const credentialsPath = path.join(__dirname, 'credentials.json');
const { client_secret, client_id, redirect_uris } = JSON.parse(fs.readFileSync(credentialsPath)).installed;
oAuth2Client.setCredentials({
    client_id,
    client_secret,
    redirect_uri: redirect_uris[0],
});

// Session middleware
app.use(session({
    secret: 'YOUR_SESSION_SECRET',
    resave: false,
    saveUninitialized: true
}));

// Routes
app.get('/', (req, res) => {
    res.send('<a href="/auth">Authorize with Google</a>');
});

app.get('/auth', (req, res) => {
    const url = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: ['https://www.googleapis.com/auth/spreadsheets.readonly']
    });
    res.redirect(url);
});

app.get('/oauth2callback', async (req, res) => {
    const code = req.query.code;
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    req.session.tokens = tokens;
    res.send('Authentication successful! You can now access Google Sheets.');
});

app.get('/sheets', async (req, res) => {
    if (!req.session.tokens) {
        return res.redirect('/');
    }

    oAuth2Client.setCredentials(req.session.tokens);

    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: 'YOUR_SPREADSHEET_ID',
            range: 'Sheet1!A1:D10',
        });

        res.json(response.data.values);
    } catch (error) {
        res.status(500).send('Error fetching data from Google Sheets');
    }
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
