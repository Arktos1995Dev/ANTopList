const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const bcrypt = require('bcryptjs');
const { processAndDownloadImages } = require('./download-images.js');

const app = express();
const PORT = process.env.PORT || 8000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname))); // Serve static files from the root directory
app.use('/images', express.static(path.join(__dirname, 'images'))); // Serve images

// --- User Authentication and List Management ---

const usersDir = path.join(__dirname, 'users');
if (!fs.existsSync(usersDir)) {
    fs.mkdirSync(usersDir);
}

// In-memory session store (for simplicity, replace with a proper session store in production)
const sessions = {};

// Helper function to get user file path
const getUserFilePath = (username) => path.join(usersDir, `${username}.json`);

// Registration
app.post('/api/register', async (req, res) => {
    const { username, password } = req.body;
    if (!username || !password) {
        return res.status(400).json({ message: 'Username and password are required' });
    }

    const userFilePath = getUserFilePath(username);
    if (fs.existsSync(userFilePath)) {
        return res.status(409).json({ message: 'User already exists' });
    }

    const hashedPassword = await bcrypt.hash(password, 10);
    const userData = { username, password: hashedPassword, list: [], friends: [] };
    fs.writeFileSync(userFilePath, JSON.stringify(userData, null, 2));

    res.status(201).json({ message: 'User registered successfully' });
});

// Login
app.post('/api/login', async (req, res) => {
    const { username, password } = req.body;
    if (!username || !password) {
        return res.status(400).json({ message: 'Username and password are required' });
    }

    const userFilePath = getUserFilePath(username);
    if (!fs.existsSync(userFilePath)) {
        return res.status(401).json({ message: 'Invalid credentials' });
    }

    const userData = JSON.parse(fs.readFileSync(userFilePath, 'utf8'));

    const isMatch = await bcrypt.compare(password, userData.password);
    if (!isMatch) {
        return res.status(401).json({ message: 'Invalid credentials' });
    }
    
    // Create a simple session token (in a real app, use JWT)
    const sessionId = Buffer.from(username).toString('base64');
    sessions[sessionId] = { username: userData.username };

    res.json({ message: 'Login successful', token: sessionId });
});

// Middleware to check for authentication
const requireAuth = (req, res, next) => {
    const token = req.headers.authorization;
    if (!token || !sessions[token]) {
        return res.status(401).json({ message: 'Unauthorized' });
    }
    req.user = sessions[token];
    next();
};

// Save list
app.post('/api/list', requireAuth, (req, res) => {
    const { list } = req.body;
    const userFilePath = getUserFilePath(req.user.username);
    const userData = JSON.parse(fs.readFileSync(userFilePath, 'utf8'));
    userData.list = list;
    fs.writeFileSync(userFilePath, JSON.stringify(userData, null, 2));
    res.json({ message: 'List saved successfully' });
});

// Load list
app.get('/api/list', requireAuth, (req, res) => {
    const userFilePath = getUserFilePath(req.user.username);
    const userData = JSON.parse(fs.readFileSync(userFilePath, 'utf8'));
    res.json(userData.list);
});

// Load a specific user's list (public)
app.get('/api/list/:username', (req, res) => {
    const { username } = req.params;
    const userFilePath = getUserFilePath(username);

    if (!fs.existsSync(userFilePath)) {
        return res.status(404).json({ message: 'User not found' });
    }

    const userData = JSON.parse(fs.readFileSync(userFilePath, 'utf8'));
    res.json(userData.list || []);
});

// Add a friend
app.post('/api/friends/add', requireAuth, (req, res) => {
    const { friendUsername } = req.body;
    const currentUser = req.user.username;

    if (!friendUsername) {
        return res.status(400).json({ message: 'Friend username is required' });
    }

    if (friendUsername === currentUser) {
        return res.status(400).json({ message: 'You cannot add yourself as a friend' });
    }

    const friendFilePath = getUserFilePath(friendUsername);
    if (!fs.existsSync(friendFilePath)) {
        return res.status(404).json({ message: 'User to be added not found' });
    }
    
    const currentUserFilePath = getUserFilePath(currentUser);
    const currentUserData = JSON.parse(fs.readFileSync(currentUserFilePath, 'utf8'));

    if (!currentUserData.friends) {
        currentUserData.friends = [];
    }

    if (currentUserData.friends.includes(friendUsername)) {
        return res.status(409).json({ message: 'This user is already your friend' });
    }

    currentUserData.friends.push(friendUsername);
    fs.writeFileSync(currentUserFilePath, JSON.stringify(currentUserData, null, 2));

    res.status(200).json({ message: `Successfully added ${friendUsername} as a friend` });
});

// Get friends list
app.get('/api/friends', requireAuth, (req, res) => {
    const currentUser = req.user.username;
    const currentUserFilePath = getUserFilePath(currentUser);
    const currentUserData = JSON.parse(fs.readFileSync(currentUserFilePath, 'utf8'));
    
    res.json(currentUserData.friends || []);
});


// --- End of User Management ---

// Get app version from package.json
let appVersion = 'unknown';
try {
    const packageJson = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
    appVersion = packageJson.version || 'unknown';
} catch (error) {
    console.error('Could not read app version from package.json:', error);
}

// API Routes
app.get('/api/version', (req, res) => {
    res.json({ version: appVersion });
});

app.post('/download-images', async (req, res) => {
    try {
        const { animeTitles } = req.body;
        if (!Array.isArray(animeTitles)) {
            return res.status(400).json({ message: 'animeTitles must be an array' });
        }
        const downloadedFiles = await processAndDownloadImages(animeTitles);
        res.json(downloadedFiles);
    } catch (error) {
        console.error('Error processing download request:', error);
        res.status(500).json({ message: 'Internal Server Error', error: error.message });
    }
});

// Serve the main HTML file
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
