const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');

// Создаем папку для изображений
const imagesDir = path.join(__dirname, 'images');
if (!fs.existsSync(imagesDir)) {
    fs.mkdirSync(imagesDir, { recursive: true });
    console.log('Создана папка images/');
}

// Функция для скачивания файла
function downloadFile(url, filepath) {
    return new Promise((resolve, reject) => {
        const protocol = url.startsWith('https') ? https : http;
        
        const request = protocol.get(url, (response) => {
            if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                // Обработка перенаправления
                return downloadFile(response.headers.location, filepath)
                    .then(resolve)
                    .catch(reject);
            }
            
            if (response.statusCode !== 200) {
                // Собираем тело ответа для лога
                let errorBody = '';
                response.on('data', chunk => errorBody += chunk);
                response.on('end', () => {
                  const err = new Error(`Статус ответа: ${response.statusCode} ${response.statusMessage}`);
                  err.body = errorBody;
                  reject(err);
                });
                return;
            }
            
            const fileStream = fs.createWriteStream(filepath);
            response.pipe(fileStream);
            
            fileStream.on('finish', () => {
                fileStream.close(resolve);
            });
            
            fileStream.on('error', (err) => {
                fs.unlink(filepath, () => {}); // Удаляем частичный файл
                reject(err);
            });
        });

        request.on('error', (err) => {
          reject(err);
        });
    });
}

// Функция для HTTP запроса
function makeRequest(url) {
    return new Promise((resolve, reject) => {
        const protocol = url.startsWith('https') ? https : http;
        
        protocol.get(url, (res) => {
            let data = '';
            res.on('data', (chunk) => data += chunk);
            res.on('end', () => {
                try {
                    if (!res.statusCode) {
                        return reject(new Error("Не удалось получить код состояния ответа."));
                    }
                    const isSuccess = res.statusCode >= 200 && res.statusCode < 300;
                    resolve({ ok: isSuccess, status: res.statusCode, statusText: res.statusMessage, json: () => Promise.resolve(JSON.parse(data)) });
                } catch (e) {
                    reject(new Error(`Ошибка парсинга JSON: ${e.message}. Ответ сервера: ${data}`));
                }
            });
        }).on('error', reject);
    });
}

// Функция для извлечения URL изображения
function getImageUrl(animeData) {
    if (!animeData || !animeData.images) return null;
    
    const { jpg, webp } = animeData.images;
    const imageUrl = (jpg && jpg.image_url) || (webp && webp.image_url);
    
    // Игнорируем стандартное изображение "заглушку"
    if (imageUrl && imageUrl.includes('questionmark')) {
        return null;
    }
    
    return imageUrl;
}

// Функция для генерации альтернативных вариантов поиска
function generateSearchAlternatives(title) {
    const alternatives = new Set([title]);
    const trimmed = title.trim();
    
    // Убираем скобки и содержимое в них
    alternatives.add(trimmed.replace(/\s*[\[(].*?[)\]]\s*/g, '').trim());
    // Часть до двоеточия
    alternatives.add(trimmed.split(':')[0].trim());
    // Очищенное от спецсимволов название
    alternatives.add(trimmed.replace(/[^a-zA-Z0-9\s-]/g, ' ').replace(/\s+/g, ' ').trim());

    alternatives.delete(''); // Удаляем пустые строки
    return Array.from(alternatives);
}

// Функция для получения изображения аниме через API
async function fetchAnimeImageWithAlternatives(animeTitle) {
    const alternatives = generateSearchAlternatives(animeTitle);
    
    for (const title of alternatives) {
        const searchUrl = `https://api.jikan.moe/v4/anime?q=${encodeURIComponent(title)}&limit=1`;
        
        try {
            // Задержка для соблюдения лимитов API
            await new Promise(resolve => setTimeout(resolve, 500));
            
            const response = await makeRequest(searchUrl);
            
            if (!response.ok) {
                 if (response.status === 429) { // Too Many Requests
                    console.log(`  🟡 Лимит запросов превышен. Ожидание 2 секунды...`);
                    await new Promise(resolve => setTimeout(resolve, 2000));
                    continue; // Повторяем попытку с тем же названием
                }
                // Не логируем как ошибку, просто пробуем следующий вариант
                continue; 
            }
            
            const data = await response.json();
            
            if (data.data && data.data.length > 0) {
                const imageUrl = getImageUrl(data.data[0]);
                if (imageUrl) {
                    if (title !== animeTitle) {
                        console.log(`  ℹ️ Найдено по альтернативному названию: "${title}"`);
                    }
                    return imageUrl;
                }
            }
        } catch (error) {
            console.error(`  ❌ Ошибка при запросе к API для "${title}":`);
            console.error(`     - URL: ${searchUrl}`);
            console.error(`     - Ошибка: ${error.message}\n`);
            // Продолжаем со следующим вариантом
        }
    }
    
    return null; // Не найдено ни по одному из вариантов
}

// Функция для создания безопасного имени файла
function sanitizeFileName(name) {
    return name.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
}

// Основная функция
async function downloadImages() {
    const files = fs.readdirSync(__dirname).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
    if (files.length === 0) {
        console.error('Ошибка: Не найден .xlsx или .xls файл в текущей директории.');
        return;
    }
    
    const xlsxFile = files[0];
    console.log(`\n📖 Читаю файл: ${xlsxFile}`);
    
    const workbook = XLSX.readFile(xlsxFile);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    const headerRow = jsonData.find(row => row.some(cell => cell)) || [];
    const animeColumnIndex = headerRow.findIndex(h => String(h).toLowerCase().trim() === 'anime');
    
    if (animeColumnIndex < 0) {
        console.error('Ошибка: В файле не найдена колонка с названием "anime".');
        return;
    }
    
    const animeTitles = [...new Set(jsonData.slice(jsonData.indexOf(headerRow) + 1)
        .map(row => String(row[animeColumnIndex]).trim())
        .filter(title => title))];
    
    console.log(`🔍 Найдено ${animeTitles.length} уникальных аниме для скачивания.\n`);
    
    let successCount = 0;
    let errorCount = 0;
    const downloadedFiles = {};
    
    for (let i = 0; i < animeTitles.length; i++) {
        const animeTitle = animeTitles[i];
        console.log(`[${i + 1}/${animeTitles.length}] Обрабатываю: "${animeTitle}"...`);
        
        const imageUrl = await fetchAnimeImageWithAlternatives(animeTitle);
        
        if (!imageUrl) {
            console.log(`  ❌ Изображение не найдено для "${animeTitle}"\n`);
            errorCount++;
            continue;
        }
        
        const safeName = sanitizeFileName(animeTitle);
        const ext = path.extname(new URL(imageUrl).pathname) || '.jpg';
        const filename = `${safeName}${ext}`;
        const filepath = path.join(imagesDir, filename);
        
        try {
            await downloadFile(imageUrl, filepath);
            downloadedFiles[animeTitle] = filename;
            console.log(`  ✅ Скачано: ${filename}\n`);
            successCount++;
        } catch (error) {
            console.error(`  ❌ Ошибка при скачивании файла для "${animeTitle}":`);
            console.error(`     - URL изображения: ${imageUrl}`);
            console.error(`     - Причина: ${error.message}`);
            if(error.body) {
              console.error(`     - Ответ сервера: ${error.body}`);
            }
            console.error(''); // Пустая строка для читаемости
            errorCount++;
        }
    }
    
    if (Object.keys(downloadedFiles).length > 0) {
        fs.writeFileSync(path.join(__dirname, 'image-mapping.json'), JSON.stringify(downloadedFiles, null, 2));
        console.log('✅ Маппинг названий на файлы сохранен в image-mapping.json');
    }

    console.log(`\n=== Итоги ===`);
    console.log(`👍 Успешно скачано: ${successCount}`);
    console.log(`👎 Ошибок: ${errorCount}`);
    console.log(`-`.repeat(15) + `\n`);
}

downloadImages().catch(error => {
    console.error('Критическая ошибка выполнения скрипта:', error);
    process.exit(1);
});
