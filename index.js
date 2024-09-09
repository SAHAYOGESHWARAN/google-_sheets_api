const express = require('express');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const dotenv = require('dotenv');

dotenv.config();

const app = express();
const PORT = 3000;

// OAuth2 client setup (Read credentials from a file)
const credentialsPath = path.join(__dirname, 'credentials.json');
let credentials;

try {
    credentials = JSON.parse(fs.readFileSync(credentialsPath));
} catch (error) {
    console.error('Error reading credentials file:', error);
    process.exit(1);
}

const { client_secret, client_id, redirect_uris } = credentials.installed;
const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
);

// Session middleware for storing tokens
app.use(session({
    secret: 'YOUR_SESSION_SECRET',
    resave: false,
    saveUninitialized: true
}));

app.use(express.json()); // To parse JSON request bodies

// Step 1: Display the Google OAuth authorization link
app.get('/', (req, res) => {
    res.send('<a href="/auth">Authorize with Google</a>');
});

// Step 2: Redirect to Google OAuth page to get permission
app.get('/auth', (req, res) => {
    const url = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: ['https://www.googleapis.com/auth/spreadsheets']
    });
    res.redirect(url);
});

// Step 3: Handle OAuth2 callback, exchange the code for tokens, and store them in session
app.get('/oauth2callback', async (req, res) => {
    const code = req.query.code;
    try {
        const { tokens } = await oAuth2Client.getToken(code);  // Get OAuth tokens
        oAuth2Client.setCredentials(tokens);
        req.session.tokens = tokens;  // Store tokens in session
        res.send('Authentication successful! You can now access Google Sheets.');
    } catch (error) {
        console.error('Error exchanging code for tokens:', error);
        res.status(500).send('Error exchanging code for tokens');
    }
});

// Step 4: Read data from Google Sheets
app.get('/sheets', async (req, res) => {
    if (!req.session.tokens) {
        return res.redirect('/auth');  // Redirect to auth if not authenticated
    }

    oAuth2Client.setCredentials(req.session.tokens);  // Set credentials from session

    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const spreadsheetId = 'YOUR_SPREADSHEET_ID';  // Replace with your spreadsheet ID

    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: 'Sheet1!A1:D10',  // Specify the range in your Google Sheet
        });

        res.json(response.data.values);  // Return the fetched data
    } catch (error) {
        console.error('Error fetching data from Google Sheets:', error);
        res.status(500).send('Error fetching data from Google Sheets');
    }
});

// Step 5: Add data to Google Sheets
app.post('/sheets/add', async (req, res) => {
    if (!req.session.tokens) {
        return res.redirect('/auth');  // Redirect to auth if not authenticated
    }

    oAuth2Client.setCredentials(req.session.tokens);  // Set credentials from session

    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const { range, values } = req.body;  // Get the range and values from the request body
    const spreadsheetId = 'YOUR_SPREADSHEET_ID';  // Replace with your spreadsheet ID

    try {
        const response = await sheets.spreadsheets.values.append({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',  // 'RAW' or 'USER_ENTERED'
            requestBody: { values },  // The data to add
        });

        console.log('Data added successfully:', response.data);  // Log response for debugging
        res.json(response.data);
    } catch (error) {
        console.error('Error adding data to Google Sheets:', error);
        res.status(500).send('Error adding data to Google Sheets');
    }
});

// Server setup
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
