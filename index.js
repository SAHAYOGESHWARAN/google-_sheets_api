const express = require('express');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const dotenv = require('dotenv');

// Load environment variables
dotenv.config();

const app = express();
const PORT = 3000;

// OAuth2 client setup
const credentialsPath = path.join(__dirname, 'credentials.json');
let credentials;

if (!fs.existsSync(credentialsPath)) {
    console.error('Credentials file not found at:', credentialsPath);
    process.exit(1);
}

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

// Session middleware
app.use(session({
    secret: 'your-session-secret',
    resave: false,
    saveUninitialized: false,
    cookie: { secure: false }
}));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Save tokens to a file
const TOKEN_PATH = path.join(__dirname, 'token.json');

// Function to load saved tokens from a file
function loadSavedTokens() {
    if (fs.existsSync(TOKEN_PATH)) {
        const tokens = JSON.parse(fs.readFileSync(TOKEN_PATH));
        oAuth2Client.setCredentials(tokens);
        console.log('Tokens loaded from file');
    }
}

// Function to save tokens to a file
function saveTokens(tokens) {
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
    console.log('Tokens saved to file');
}

// Load saved tokens when the server starts
loadSavedTokens();

// Routes
app.get('/', (req, res) => {
    res.send('<a href="/auth">Authorize with Google</a>');
});

// Google OAuth authorization route
app.get('/auth', (req, res) => {
    const url = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: ['https://www.googleapis.com/auth/spreadsheets']
    });
    res.redirect(url);
});

// OAuth2 callback route
app.get('/oauth2callback', async (req, res) => {
    const code = req.query.code;
    if (!code) {
        return res.status(400).send('No code found in query parameters');
    }

    try {
        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        req.session.tokens = tokens;
        saveTokens(tokens);  // Save tokens to a file for reuse
        res.redirect('/enter-details');  // Redirect to the form page
    } catch (error) {
        console.error('Error exchanging code for tokens:', error.response ? error.response.data : error.message);
        res.status(500).send('Error exchanging code for tokens');
    }
});

// Route to show the form for entering details
app.get('/enter-details', (req, res) => {
    if (!oAuth2Client.credentials) {
        return res.redirect('/');
    }

    // Simple HTML form to enter data
    res.send(`
        <form action="/sheets/add" method="post">
            <label for="name">Name:</label>
            <input type="text" name="name" required><br><br>
            <label for="email">Email:</label>
            <input type="email" name="email" required><br><br>
            <button type="submit">Submit</button>
        </form>
    `);
});

// Add data to Google Sheets (handles form submission)
app.post('/sheets/add', async (req, res) => {
    if (!oAuth2Client.credentials) {
        return res.redirect('/');
    }

    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const { name, email } = req.body;
    const spreadsheetId = process.env.SPREADSHEET_ID;

    if (!name || !email) {
        return res.status(400).send('Name and email are required');
    }

    // Prepare the data to be added
    const values = [[name, email]];
    const range = 'Sheet1!A1:B1';  // You can adjust this range

    try {
        const response = await sheets.spreadsheets.values.append({
            spreadsheetId: spreadsheetId,
            range: range,
            valueInputOption: 'RAW',
            requestBody: {
                values: values,
            },
        });

        console.log('Data added successfully:', response.data);
        res.send('Data added successfully!');
    } catch (error) {
        console.error('Error adding data to Google Sheets:', error.response ? error.response.data : error.message);
        res.status(500).send(`Error adding data to Google Sheets: ${error.response ? error.response.data : error.message}`);
    }
});

// Start server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
