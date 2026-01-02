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

// In-memory session store
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
        // Fetch user by username
        const { data: user, error } = await supabase
            .from('users')
            .select('id, username, password') // Select id as well
            .eq('username', username)
            .single();

        if (error || !user) {
            return res.status(401).json({ message: 'Invalid credentials' });
        }

        const isMatch = await bcrypt.compare(password, user.password);

        if (!isMatch) {
            return res.status(401).json({ message: 'Invalid credentials' });
        }

        // Create a session token and store user id and username
        const sessionId = Buffer.from(username).toString('base64');
        sessions[sessionId] = { id: user.id, username: user.username };

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
    // Attach user info (id and username) to the request
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
            .eq('id', req.user.id); // Use id for update

        if (error) throw error;
        res.json({ message: 'List saved successfully' });
    } catch (error) {
        console.error('Save list error:', error);
        res.status(500).json({ message: 'Internal server error while saving list' });
    }
});

// Load own list
app.get('/api/list', requireAuth, async (req, res) => {
    try {
        const { data: user, error } = await supabase
            .from('users')
            .select('list')
            .eq('id', req.user.id) // Use id to fetch list
            .single();

        if (error) throw error;
        res.json(user?.list || []);
    } catch (error) {
        console.error('Load list error:', error);
        res.status(500).json({ message: 'Internal server error while loading list' });
    }
});


// Load a specific user's list (public)
app.get('/api/list/:username', async (req, res) => {
    const { username } = req.params;
    try {
        const { data: user, error } = await supabase
            .from('users')
            .select('list')
            .eq('username', username)
            .single();
        
        if (error || !user) {
            return res.status(404).json({ message: 'User not found' });
        }
        res.json(user.list || []);
    } catch (error) {
        console.error('Load user list error:', error);
        res.status(500).json({ message: 'Internal server error while loading user list' });
    }
});

app.get('/api/image/:mal_id', async (req, res) => {
    const { mal_id } = req.params;
    try {
        const { data, error } = await supabase
            .from('images')
            .select('image_data, image_url')
            .eq('mal_id', mal_id)
            .single();

        if (data && data.image_data) {
            // Determine content type from URL, default to jpeg
            const extension = path.extname(data.image_url).toLowerCase();
            let contentType = 'image/jpeg';
            if (extension === '.png') {
                contentType = 'image/png';
            } else if (extension === '.gif') {
                contentType = 'image/gif';
            } else if (extension === '.webp') {
                contentType = 'image/webp';
            }
            res.setHeader('Content-Type', contentType);
            res.send(Buffer.from(data.image_data, 'base64')); // Send the binary data
        } else {
             if (error) {
                console.error(`Error fetching image for mal_id ${mal_id}:`, error);
            }
            res.status(404).send('Not Found');
        }
    } catch (error) {
        res.status(500).send('Internal Server Error');
    }
});

// --- New Friend Endpoints (3NF) ---

// Add a friend (using the new 'friendships' table)
app.post('/api/friends/add', requireAuth, async (req, res) => {
    const { friendUsername } = req.body;
    const currentUserId = req.user.id;
    const currentUsername = req.user.username;

    if (!friendUsername) {
        return res.status(400).json({ message: 'Friend username is required' });
    }

    if (friendUsername === currentUsername) {
        return res.status(400).json({ message: 'You cannot add yourself as a friend' });
    }

    try {
        // 1. Find the friend's user data
        const { data: friend, error: friendError } = await supabase
            .from('users')
            .select('id')
            .eq('username', friendUsername)
            .single();

        if (friendError || !friend) {
            return res.status(404).json({ message: 'User to be added not found' });
        }
        const friendId = friend.id;

        // 2. Check if the friendship already exists
        const { data: existingFriendship, error: checkError } = await supabase
            .from('friendships')
            .select('*')
            .or(`(user_id.eq.${currentUserId},friend_id.eq.${friendId}),(user_id.eq.${friendId},friend_id.eq.${currentUserId})`);

        if (checkError) throw checkError;

        if (existingFriendship && existingFriendship.length > 0) {
            return res.status(409).json({ message: 'This user is already your friend' });
        }

        // 3. Create the friendship (both ways for easier lookup)
        const { error: insertError } = await supabase
            .from('friendships')
            .insert([
                { user_id: currentUserId, friend_id: friendId },
                { user_id: friendId, friend_id: currentUserId } // Mutual friendship
            ]);

        if (insertError) {
            // Handle potential race conditions or other db errors
            if (insertError.code === '23505') { // unique_violation
                 return res.status(409).json({ message: 'This user is already your friend' });
            }
            throw insertError;
        }

        res.status(200).json({ message: `Successfully added ${friendUsername} as a friend` });
    } catch (error) {
        console.error('Add friend error:', error);
        res.status(500).json({ message: 'Internal server error while adding friend' });
    }
});

// Get friends list (from the new 'friendships' table)
app.get('/api/friends', requireAuth, async (req, res) => {
    const currentUserId = req.user.id;
    try {
        // Perform a join query
        const { data, error } = await supabase
            .from('friendships')
            .select('friend:users!friendships_friend_id_fkey(username)')
            .eq('user_id', currentUserId);

        if (error) throw error;

        const friendsUsernames = data.map(item => item.friend.username);
        res.json(friendsUsernames);

    } catch (error) {
        console.error('Get friends error:', error);
        res.status(500).json({ message: 'Internal server error while getting friends list' });
    }
});

// Get user profile (username and friends)
app.get('/api/profile', requireAuth, async (req, res) => {
    const currentUserId = req.user.id;
    const currentUsername = req.user.username;

    try {
        // Fetch friends' usernames
        const { data, error } = await supabase
            .from('friendships')
            .select('friend:users!friendships_friend_id_fkey(username)')
            .eq('user_id', currentUserId);

        if (error) throw error;

        const friendsUsernames = data.map(item => item.friend.username);
        
        res.json({
            username: currentUsername,
            friends: friendsUsernames
        });

    } catch (error) {
        console.error('Get profile error:', error);
        res.status(500).json({ message: 'Internal server error while getting profile' });
    }
});

// Search for users
app.get('/api/users/search', async (req, res) => {
    const { username } = req.query;

    if (!username) {
        return res.status(400).json({ message: 'Username query parameter is required' });
    }

    try {
        const { data, error } = await supabase
            .from('users')
            .select('username')
            .ilike('username', `%${username}%`); // Case-insensitive search

        if (error) throw error;

        res.json(data);
    } catch (error) {
        console.error('User search error:', error);
        res.status(500).json({ message: 'Internal server error during user search' });
    }
});



// --- End New Friend Endpoints ---


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
        const downloadedFiles = await processAndDownloadImages(supabase, animeTitles);
        res.json(downloadedFiles);
    } catch (error) {
        console.error('Error processing download request:', error);
        res.status(500).json({ message: 'Internal Server Error', error: error.message });
    }
});

// Route for profile pages, serves index.html
app.get('/профиль/:username', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/profile', (req, res) => {
    res.sendFile(path.join(__dirname, 'profile.html'));
});

// Serve the main HTML file for all other GET requests
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});