const express = require('express');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const dotenv = require('dotenv');
const mongoose = require('mongoose');

// Load environment variables
dotenv.config();

const app = express();
const PORT = 5000;

// MongoDB connection setup
mongoose.connect(process.env.MONGODB_URI, { useNewUrlParser: true, useUnifiedTopology: true })
    .then(() => console.log('Connected to MongoDB'))
    .catch(err => console.error('Error connecting to MongoDB:', err));

// Define a Mongoose schema for user data
const userSchema = new mongoose.Schema({
    name: { type: String, required: true },
    email: { type: String, required: true, unique: true }
});

const User = mongoose.model('User', userSchema);

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
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

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

function loadSavedTokens() {
    if (fs.existsSync(TOKEN_PATH)) {
        const tokens = JSON.parse(fs.readFileSync(TOKEN_PATH));
        oAuth2Client.setCredentials(tokens);
        console.log('Tokens loaded from file');
    }
}

function saveTokens(tokens) {
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
    console.log('Tokens saved to file');
}

loadSavedTokens();

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
    if (!code) {
        return res.status(400).send('No code found in query parameters');
    }

    try {
        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        req.session.tokens = tokens;
        saveTokens(tokens);
        res.redirect('/enter-details');
    } catch (error) {
        console.error('Error exchanging code for tokens:', error.response ? error.response.data : error.message);
        res.status(500).send('Error exchanging code for tokens');
    }
});

// Route to show form for entering details
app.get('/enter-details', (req, res) => {
    if (!oAuth2Client.credentials) {
        return res.redirect('/');
    }

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

// Add data to Google Sheets and MongoDB (handles form submission)
app.post('/sheets/add', async (req, res) => {
    if (!oAuth2Client.credentials) {
        return res.redirect('/');
    }

    const { name, email } = req.body;

    // Check for duplicates in MongoDB
    const existingUser = await User.findOne({ email });
    if (existingUser) {
        return res.status(400).send('This email already exists. Duplicate entries are not allowed.');
    }

    // If no duplicates, save to MongoDB
    const newUser = new User({ name, email });
    await newUser.save();

    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const spreadsheetId = process.env.SPREADSHEET_ID;
    const values = [[name, email]];
    const range = 'Sheet1!A1:B1';

    try {
        const response = await sheets.spreadsheets.values.append({
            spreadsheetId: spreadsheetId,
            range: range,
            valueInputOption: 'RAW',
            requestBody: { values }
        });

        console.log('Data added successfully to Google Sheets:', response.data);
        res.send('Data added successfully to both MongoDB and Google Sheets!');
    } catch (error) {
        console.error('Error adding data to Google Sheets:', error.response ? error.response.data : error.message);
        res.status(500).send(`Error adding data to Google Sheets: ${error.response ? error.response.data : error.message}`);
    }
});

// Start server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
