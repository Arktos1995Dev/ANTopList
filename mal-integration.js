const https = require('https');

function makeRequest(url) {
    return new Promise((resolve, reject) => {
        https.get(url, (res) => {
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

async function getMalUserAnimeList(username) {
    if (!username) {
        throw new Error("MyAnimeList username is required.");
    }

    const url = `https://api.jikan.moe/v4/users/${username}/animelist`;
    console.log(`Fetching anime list for user: ${username}`);

    try {
        const response = await makeRequest(url);

        if (!response.ok) {
            throw new Error(`Failed to fetch data from MAL. Status: ${response.status}`);
        }

        const data = await response.json();
        if (!data.data || !Array.isArray(data.data)) {
            console.warn("No anime list data found for this user or the data is in an unexpected format.");
            return [];
        }

        const animeTitles = data.data.map(item => item.anime.title);
        return [...new Set(animeTitles)]; // Return unique titles
    } catch (error) {
        console.error(`Error fetching MAL user list: ${error.message}`);
        throw error;
    }
}

module.exports = { getMalUserAnimeList };