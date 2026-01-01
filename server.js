const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { processAndDownloadImages } = require('./download-images.js');

const app = express();
const PORT = process.env.PORT || 8000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname))); // Serve static files from the root directory
app.use('/images', express.static(path.join(__dirname, 'images'))); // Serve images

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
