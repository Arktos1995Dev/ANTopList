const https = require('https');
const http = require('http');

function downloadFile(url) {
    return new Promise((resolve, reject) => {
        const protocol = url.startsWith('https') ? https : http;
        const request = protocol.get(url, (response) => {
            if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                downloadFile(response.headers.location).then(resolve).catch(reject);
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
            const chunks = [];
            response.on('data', (chunk) => {
                chunks.push(chunk);
            });
            response.on('end', () => {
                resolve(Buffer.concat(chunks));
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

async function fetchAnimeDataWithAlternatives(animeTitle) {
    const alternatives = generateSearchAlternatives(animeTitle);
    for (const title of alternatives) {
        const searchUrl = `https://api.jikan.moe/v4/anime?q=${encodeURIComponent(title)}&limit=1`;
        try {
            await new Promise(resolve => setTimeout(resolve, 500)); // Rate limit Jikan API
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
                    if (title !== animeTitle) console.log(`Found via \"${title}\"`);
                    return data.data[0];
                }
            }
        } catch (error) {
            console.error(`API request error for \"${title}\": ${error.message}`);
        }
    }
    return null;
}

async function processAndDownloadImages(supabase, animeTitles) {
    console.log(`Starting download for ${animeTitles.length} unique anime titles.`);
    const downloadedImages = {};

    for (const animeTitle of animeTitles) {
        console.log(`Processing: \"${animeTitle}\"...`);

        const animeData = await fetchAnimeDataWithAlternatives(animeTitle);
        if (!animeData) {
            console.log(`Image not found for \"${animeTitle}\"`);
            continue;
        }

        const mal_id = animeData.mal_id;
        const imageUrl = getImageUrl(animeData);

        const { data: existingImage, error: selectError } = await supabase
            .from('images')
            .select('mal_id, image_data') // Also select image_data to check for existence
            .eq('mal_id', mal_id)
            .single();

        // If image exists and has data, skip.
        if (existingImage && existingImage.image_data) { 
            downloadedImages[animeTitle] = mal_id;
            console.log(`Found in DB with image data: ${mal_id}`);
            continue;
        }

        if (selectError && selectError.code !== 'PGRST116') { // Ignore "0 rows" error
            console.error(`Error checking DB for \"${animeTitle}\":`, selectError.message);
            continue;
        }

        try {
            const imageData = await downloadFile(imageUrl); // Returns a Buffer
            
            // Convert buffer to base64 string for storing in bytea column
            const imageBase64 = imageData.toString('base64');

            const { error: insertError } = await supabase
                .from('images')
                .insert({
                    mal_id: mal_id,
                    image_url: imageUrl,
                    image_data: imageBase64 // Save as base64 string
                }, { upsert: true }); // Use upsert to avoid race conditions and update existing rows

            if (insertError) {
                console.error(`Failed to insert DB record for \"${animeTitle}\": ${insertError.message}`);
            }

            downloadedImages[animeTitle] = mal_id;
            console.log(`Downloaded and saved to DB: ${mal_id}`);

        } catch (error) {
            console.error(`Failed to download file for \"${animeTitle}\" from ${imageUrl}: ${error.message}`);
        }
    }
    return downloadedImages;
}

module.exports = { processAndDownloadImages };