const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');

// --- Directory Setup ---
// Ensure the 'images' directory exists in the project root.
const imagesDir = path.join(__dirname, 'images');
if (!fs.existsSync(imagesDir)) {
    try {
        fs.mkdirSync(imagesDir, { recursive: true });
        console.log(`Created directory: ${imagesDir}`);
    } catch (error) {
        console.error(`Failed to create directory ${imagesDir}:`, error);
    }
}
// --- End Directory Setup ---


function downloadFile(url) {
    return new Promise((resolve, reject) => {
        const protocol = url.startsWith('https') ? https : http;
        const request = protocol.get(url, (response) => {
            if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                // Handle redirects
                const redirectUrl = new URL(response.headers.location, url).href;
                downloadFile(redirectUrl).then(resolve).catch(reject);
                return;
            }
            if (response.statusCode !== 200) {
                let errorBody = '';
                response.on('data', chunk => errorBody += chunk);
                response.on('end', () => {
                    const err = new Error(`HTTP Status: ${response.statusCode} ${response.statusMessage} for URL: ${url}`);
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
        request.on('error', (err) => {
            reject(new Error(`Request error for ${url}: ${err.message}`));
        });
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
    // Filter out placeholder images
    return imageUrl && !imageUrl.includes('questionmark') && !imageUrl.includes('ma-icon') ? imageUrl : null;
}

function generateSearchAlternatives(title) {
    const alternatives = new Set([title.trim()]);
    // Remove content in brackets (e.g., " (TV)")
    alternatives.add(title.trim().replace(/\s*[\[(].*?[)\]]\s*/g, '').trim());
    // Take content before a colon
    alternatives.add(title.trim().split(':')[0].trim());
    // Sanitize by removing special characters
    alternatives.add(title.trim().replace(/[^a-zA-Z0-9\s-]/g, ' ').replace(/\s+/g, ' ').trim());
    alternatives.delete(''); // Remove empty strings
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
                continue; // Retry with the same title
            }
            if (!response.ok) continue;
            const data = await response.json();
            if (data.data && data.data.length > 0) {
                const imageUrl = getImageUrl(data.data[0]);
                if (imageUrl) {
                    if (title !== animeTitle) console.log(`Found \"${animeTitle}\" as \"${title}\"`);
                    return data.data[0];
                }
            }
        } catch (error) {
            console.error(`API request error for \"${title}\" at ${searchUrl}: ${error.message}`);
        }
    }
    return null;
}

async function processAndDownloadImages(supabase, animeTitles) {
    console.log(`Starting image processing for ${animeTitles.length} unique anime titles.`);
    const imagePathMapping = {}; // This will now map title -> local file path

    for (const animeTitle of animeTitles) {
        console.log(`Processing: \"${animeTitle}\"...`);

        const animeData = await fetchAnimeDataWithAlternatives(animeTitle);
        if (!animeData) {
            console.log(`Could not find anime data for \"${animeTitle}\"`);
            continue;
        }

        const mal_id = animeData.mal_id;
        const remoteImageUrl = getImageUrl(animeData);

        if (!remoteImageUrl) {
            console.log(`Image URL not found for \"${animeTitle}\" (mal_id: ${mal_id})`);
            continue;
        }

        // Check DB for an existing entry for this mal_id
        const { data: existingDbEntry, error: dbError } = await supabase
            .from('images')
            .select('mal_id, image_url')
            .eq('mal_id', mal_id)
            .single();

        if (dbError && dbError.code !== 'PGRST116') { // Ignore "0 rows" error
            console.error(`Error checking DB for mal_id ${mal_id}:`, dbError.message);
            continue;
        }

        // If an entry exists, check if its URL is local and if the file exists.
        if (existingDbEntry && existingDbEntry.image_url && existingDbEntry.image_url.startsWith('/images/')) {
            const localPath = path.join(__dirname, existingDbEntry.image_url);
            if (fs.existsSync(localPath)) {
                console.log(`Image found locally for \"${animeTitle}\": ${existingDbEntry.image_url}`);
                imagePathMapping[animeTitle] = existingDbEntry.image_url;
                continue; // Skip to next title
            }
        }

        // If we're here, we need to download the image.
        try {
            // Determine file extension. Default to .jpg.
            const extension = path.extname(new URL(remoteImageUrl).pathname).toLowerCase() || '.jpg';
            const validExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.webp'];
            const finalExtension = validExtensions.includes(extension) ? extension : '.jpg';
            
            const localFileName = `${mal_id}${finalExtension}`;
            const publicPath = `/images/${localFileName}`; // The path to be stored in DB and used in src attribute
            const absoluteFilePath = path.join(imagesDir, localFileName);

            console.log(`Downloading ${remoteImageUrl} to ${absoluteFilePath}`);
            const imageDataBuffer = await downloadFile(remoteImageUrl);
            
            fs.writeFileSync(absoluteFilePath, imageDataBuffer);
            console.log(`Successfully saved ${absoluteFilePath}`);

            // Now, upsert the DB with the new local path
            const { error: upsertError } = await supabase
                .from('images')
                .upsert({
                    mal_id: mal_id,
                    image_url: publicPath // Store the public URL path
                });

            if (upsertError) {
                console.error(`Failed to upsert DB record for \"${animeTitle}\" (mal_id: ${mal_id}): ${upsertError.message}`);
            } else {
                console.log(`Updated DB for ${mal_id} with path: ${publicPath}`);
            }

            imagePathMapping[animeTitle] = publicPath;

        } catch (error) {
            console.error(`Failed during download/save process for \"${animeTitle}\" from ${remoteImageUrl}: ${error.message}`);
        }
    }
    console.log("Finished image processing.");
    return imagePathMapping; // Return the new mapping: { "title": "/images/12345.jpg" }
}

module.exports = { processAndDownloadImages };