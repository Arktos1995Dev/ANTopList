const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');
const { Pool } = require('pg');

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: {
    rejectUnauthorized: false
  }
});

const imagesDir = path.join(__dirname, 'images');
if (!fs.existsSync(imagesDir)) {
    fs.mkdirSync(imagesDir, { recursive: true });
}

async function createImageTable() {
    const client = await pool.connect();
    try {
        await client.query(`
            CREATE TABLE IF NOT EXISTS images (
                id SERIAL PRIMARY KEY,
                title TEXT UNIQUE NOT NULL,
                filename TEXT NOT NULL
            )
        `);
    } finally {
        client.release();
    }
}

createImageTable();

function downloadFile(url, filepath) {
    return new Promise((resolve, reject) => {
        const protocol = url.startsWith('https') ? https : http;
        const request = protocol.get(url, (response) => {
            if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                downloadFile(response.headers.location, filepath).then(resolve).catch(reject);
                return;
            }
            if (response.statusCode !== 200) {
                let errorBody = '';
                response.on('data', chunk => errorBody += chunk);
                response.on('end', () => {
                    const err = new Error(`HTTP Status: ${response.statusCode} ${response.statusMessage}`);
                    err.body = errorBody;
                    reject(err);
                });
                return;
            }
            const fileStream = fs.createWriteStream(filepath);
            response.pipe(fileStream);
            fileStream.on('finish', () => fileStream.close(resolve));
            fileStream.on('error', (err) => {
                fs.unlink(filepath, () => {});
                reject(err);
            });
        });
        request.on('error', reject);
    });
}

function makeRequest(url) {
    return new Promise((resolve, reject) => {
        const protocol = url.startsWith('https') ? https : http;
        protocol.get(url, (res) => {
            let data = '';
            res.on('data', (chunk) => data += chunk);
            res.on('end', () => {
                try {
                    if (!res.statusCode) return reject(new Error("No status code received."));
                    const isSuccess = res.statusCode >= 200 && res.statusCode < 300;
                    resolve({ ok: isSuccess, status: res.statusCode, json: () => Promise.resolve(JSON.parse(data)) });
                } catch (e) {
                    reject(new Error(`JSON parsing error: ${e.message}. Server response: ${data}`));
                }
            });
        }).on('error', reject);
    });
}

function getImageUrl(animeData) {
    if (!animeData || !animeData.images) return null;
    const { jpg, webp } = animeData.images;
    const imageUrl = (jpg && jpg.image_url) || (webp && webp.image_url);
    return imageUrl && !imageUrl.includes('questionmark') ? imageUrl : null;
}

function generateSearchAlternatives(title) {
    const alternatives = new Set([title.trim()]);
    alternatives.add(title.trim().replace(/\s*[\[(].*?[)\]]\s*/g, '').trim());
    alternatives.add(title.trim().split(':')[0].trim());
    alternatives.add(title.trim().replace(/[^a-zA-Z0-9\s-]/g, ' ').replace(/\s+/g, ' ').trim());
    alternatives.delete('');
    return Array.from(alternatives);
}

async function fetchAnimeImageWithAlternatives(animeTitle) {
    const alternatives = generateSearchAlternatives(animeTitle);
    for (const title of alternatives) {
        const searchUrl = `https://api.jikan.moe/v4/anime?q=${encodeURIComponent(title)}&limit=1`;
        try {
            await new Promise(resolve => setTimeout(resolve, 500));
            const response = await makeRequest(searchUrl);
            if (response.status === 429) {
                console.log(`Rate limit hit. Waiting 2 seconds...`);
                await new Promise(resolve => setTimeout(resolve, 2000));
                continue;
            }
            if (!response.ok) continue;
            const data = await response.json();
            if (data.data && data.data.length > 0) {
                const imageUrl = getImageUrl(data.data[0]);
                if (imageUrl) {
                    if (title !== animeTitle) console.log(`Found via "${title}"`);
                    return imageUrl;
                }
            }
        } catch (error) {
            console.error(`API request error for "${title}": ${error.message}`);
        }
    }
    return null;
}

function sanitizeFileName(name) {
    return name.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
}

async function processAndDownloadImages(animeTitles) {
    console.log(`Starting download for ${animeTitles.length} unique anime titles.`);
    const downloadedFiles = {};
    const client = await pool.connect();
    try {
        for (const animeTitle of animeTitles) {
            console.log(`Processing: "${animeTitle}"...`);

            const dbResult = await client.query('SELECT filename FROM images WHERE title = $1', [animeTitle]);
            if (dbResult.rows.length > 0) {
                downloadedFiles[animeTitle] = dbResult.rows[0].filename;
                console.log(`Found in DB: ${dbResult.rows[0].filename}`);
                continue;
            }

            const imageUrl = await fetchAnimeImageWithAlternatives(animeTitle);
            if (!imageUrl) {
                console.log(`Image not found for "${animeTitle}"`);
                continue;
            }
            const safeName = sanitizeFileName(animeTitle);
            const ext = path.extname(new URL(imageUrl).pathname) || '.jpg';
            const filename = `${safeName}${ext}`;
            const filepath = path.join(imagesDir, filename);
            try {
                await downloadFile(imageUrl, filepath);
                await client.query('INSERT INTO images (title, filename) VALUES ($1, $2)', [animeTitle, filename]);
                downloadedFiles[animeTitle] = filename;
                console.log(`Downloaded: ${filename}`);
            } catch (error) {
                console.error(`Failed to download file for "${animeTitle}" from ${imageUrl}: ${error.message}`);
            }
        }
    } finally {
        client.release();
    }

    if (Object.keys(downloadedFiles).length > 0) {
        fs.writeFileSync(path.join(__dirname, 'image-mapping.json'), JSON.stringify(downloadedFiles, null, 2));
        console.log('Image mapping file created.');
    }
    return downloadedFiles;
}

module.exports = { processAndDownloadImages };
