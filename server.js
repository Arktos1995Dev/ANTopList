require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const bcrypt = require('bcryptjs');
const { createClient } = require('@supabase/supabase-js');
const { processAndDownloadImages } = require('./download-images.js');

const app = express();
const PORT = process.env.PORT || 8000;

// --- Supabase Setup ---
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);
// --- End Supabase Setup ---

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname)));
app.use('/images', express.static(path.join(__dirname, 'images')));

// In-memory session store (for simplicity, replace with a proper session store in production)
const sessions = {};

// Registration
app.post('/api/register', async (req, res) => {
    const { username, password } = req.body;
    if (!username || !password) {
        return res.status(400).json({ message: 'Username and password are required' });
    }

    try {
        const hashedPassword = await bcrypt.hash(password, 10);
        const { error } = await supabase
            .from('users')
            .insert([{ username, password: hashedPassword }]);

        if (error) {
            if (error.code === '23505') { // unique_violation
                return res.status(409).json({ message: 'User already exists' });
            }
            throw error;
        }

        res.status(201).json({ message: 'User registered successfully' });
    } catch (error) {
        console.error('Registration error:', error);
        res.status(500).json({ message: 'Internal server error during registration' });
    }
});

// Login
app.post('/api/login', async (req, res) => {
    const { username, password } = req.body;
    if (!username || !password) {
        return res.status(400).json({ message: 'Username and password are required' });
    }

    try {
        const { data: users, error } = await supabase
            .from('users')
            .select('*')
            .eq('username', username);

        if (error || !users || users.length === 0) {
            return res.status(401).json({ message: 'Invalid credentials' });
        }

        const user = users[0];
        const isMatch = await bcrypt.compare(password, user.password);

        if (!isMatch) {
            return res.status(401).json({ message: 'Invalid credentials' });
        }

        const sessionId = Buffer.from(username).toString('base64');
        sessions[sessionId] = { username: user.username };

        res.json({ message: 'Login successful', token: sessionId });
    } catch (error) {
        console.error('Login error:', error);
        res.status(500).json({ message: 'Internal server error during login' });
    }
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
app.post('/api/list', requireAuth, async (req, res) => {
    const { list } = req.body;
    try {
        const { error } = await supabase
            .from('users')
            .update({ list: list })
            .eq('username', req.user.username);

        if (error) throw error;
        res.json({ message: 'List saved successfully' });
    } catch (error) {
        console.error('Save list error:', error);
        res.status(500).json({ message: 'Internal server error while saving list' });
    }
});

// Load list
app.get('/api/list', requireAuth, async (req, res) => {
    try {
        const { data: users, error } = await supabase
            .from('users')
            .select('list')
            .eq('username', req.user.username);

        if (error) throw error;
        res.json(users[0]?.list || []);
    } catch (error) {
        console.error('Load list error:', error);
        res.status(500).json({ message: 'Internal server error while loading list' });
    }
});

// Load a specific user's list (public)
app.get('/api/list/:username', async (req, res) => {
    const { username } = req.params;
    try {
        const { data: users, error } = await supabase
            .from('users')
            .select('list')
            .eq('username', username);
        
        if (error) throw error;

        if (!users || users.length === 0) {
            return res.status(404).json({ message: 'User not found' });
        }
        res.json(users[0].list || []);
    } catch (error) {
        console.error('Load user list error:', error);
        res.status(500).json({ message: 'Internal server error while loading user list' });
    }
});

// Add a friend
app.post('/api/friends/add', requireAuth, async (req, res) => {
    const { friendUsername } = req.body;
    const currentUser = req.user.username;

    if (!friendUsername) {
        return res.status(400).json({ message: 'Friend username is required' });
    }

    if (friendUsername === currentUser) {
        return res.status(400).json({ message: 'You cannot add yourself as a friend' });
    }

    try {
        const { data: friend, error: friendError } = await supabase
            .from('users')
            .select('username')
            .eq('username', friendUsername)
            .single();

        if (friendError || !friend) {
            return res.status(404).json({ message: 'User to be added not found' });
        }

        const { data: currentUserData, error: currentUserError } = await supabase
            .from('users')
            .select('friends')
            .eq('username', currentUser)
            .single();
        
        if (currentUserError) throw currentUserError;

        const friends = currentUserData.friends || [];

        if (friends.includes(friendUsername)) {
            return res.status(409).json({ message: 'This user is already your friend' });
        }

        friends.push(friendUsername);

        const { error: updateError } = await supabase
            .from('users')
            .update({ friends: friends })
            .eq('username', currentUser);

        if (updateError) throw updateError;

        res.status(200).json({ message: `Successfully added ${friendUsername} as a friend` });
    } catch (error) {
        console.error('Add friend error:', error);
        res.status(500).json({ message: 'Internal server error while adding friend' });
    }
});

// Get friends list
app.get('/api/friends', requireAuth, async (req, res) => {
    try {
        const { data: users, error } = await supabase
            .from('users')
            .select('friends')
            .eq('username', req.user.username);

        if (error) throw error;

        res.json(users[0]?.friends || []);
    } catch (error) {
        console.error('Get friends error:', error);
        res.status(500).json({ message: 'Internal server error while getting friends' });
    }
});


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
