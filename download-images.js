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
        
        protocol.get(url, (response) => {
            if (response.statusCode === 301 || response.statusCode === 302) {
                // Перенаправление
                return downloadFile(response.headers.location, filepath)
                    .then(resolve)
                    .catch(reject);
            }
            
            if (response.statusCode !== 200) {
                reject(new Error(`Ошибка HTTP: ${response.statusCode}`));
                return;
            }
            
            const fileStream = fs.createWriteStream(filepath);
            response.pipe(fileStream);
            
            fileStream.on('finish', () => {
                fileStream.close();
                resolve();
            });
            
            fileStream.on('error', (err) => {
                fs.unlink(filepath, () => {});
                reject(err);
            });
        }).on('error', reject);
    });
}

// Функция для HTTP запроса
function makeRequest(url) {
    return new Promise((resolve, reject) => {
        const protocol = url.startsWith('https') ? https : http;
        
        protocol.get(url, (res) => {
            let data = '';
            
            res.on('data', (chunk) => {
                data += chunk;
            });
            
            res.on('end', () => {
                try {
                    const jsonData = JSON.parse(data);
                    resolve({ ok: res.statusCode === 200, status: res.statusCode, statusText: res.statusMessage, json: () => Promise.resolve(jsonData) });
                } catch (error) {
                    reject(error);
                }
            });
        }).on('error', reject);
    });
}

// Функция для извлечения URL изображения
function getImageUrl(images, animeData) {
    if (!images) return null;
    
    let imageUrl = null;
    
    // Приоритет: jpg -> webp
    if (images.jpg && images.jpg.image_url) {
        imageUrl = images.jpg.image_url;
    } else if (images.webp && images.webp.image_url) {
        imageUrl = images.webp.image_url;
    } else if (animeData.images && animeData.images.image_url) {
        imageUrl = animeData.images.image_url;
    }
    
    if (imageUrl && imageUrl !== 'https://cdn.myanimelist.net/images/questionmark_23.gif') {
        return imageUrl;
    }
    
    return null;
}

// Функция для генерации альтернативных вариантов поиска
function generateSearchAlternatives(title) {
    const alternatives = [title];
    const trimmed = title.trim();
    
    // Убираем скобки и содержимое в них
    const withoutBrackets = trimmed.replace(/[\(\[].*?[\)\]]/g, '').trim();
    if (withoutBrackets && withoutBrackets !== trimmed) {
        alternatives.push(withoutBrackets);
    }
    
    // Берем только первое слово/часть до двоеточия
    const beforeColon = trimmed.split(':')[0].trim();
    if (beforeColon && beforeColon !== trimmed) {
        alternatives.push(beforeColon);
    }
    
    // Берем только первое слово
    const firstWord = trimmed.split(/\s+/)[0];
    if (firstWord && firstWord.length > 2 && firstWord !== trimmed) {
        alternatives.push(firstWord);
    }
    
    // Убираем специальные символы
    const cleaned = trimmed.replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim();
    if (cleaned && cleaned !== trimmed) {
        alternatives.push(cleaned);
    }
    
    return [...new Set(alternatives)]; // Убираем дубликаты
}

// Функция для получения изображения аниме через API с повторными попытками
async function fetchAnimeImage(animeTitle, retryCount = 0) {
    const maxRetries = 3;
    const baseDelay = 1000; // 1 секунда базовая задержка
    
    try {
        const searchUrl = `https://api.jikan.moe/v4/anime?q=${encodeURIComponent(animeTitle)}&limit=1`;
        
        const response = await makeRequest(searchUrl);
        
        // Обработка ошибки 429 (Too Many Requests)
        if (response.status === 429) {
            if (retryCount < maxRetries) {
                const delay = baseDelay * Math.pow(2, retryCount); // Экспоненциальная задержка
                console.log(`  ⏳ Лимит запросов превышен, ожидание ${delay/1000} сек...`);
                await new Promise(resolve => setTimeout(resolve, delay));
                return fetchAnimeImage(animeTitle, retryCount + 1);
            } else {
                throw new Error(`API ошибка 429: Превышен лимит запросов после ${maxRetries} попыток`);
            }
        }
        
        if (!response.ok) {
            throw new Error(`API ошибка: ${response.status} ${response.statusText}`);
        }
        
        const data = response.json();
        
        // Проверяем, есть ли результаты
        if (!data.data || data.data.length === 0) {
            // Пробуем альтернативные варианты поиска
            const alternatives = generateSearchAlternatives(animeTitle);
            for (const altTitle of alternatives) {
                if (altTitle === animeTitle) continue; // Уже проверили
                
                const altSearchUrl = `https://api.jikan.moe/v4/anime?q=${encodeURIComponent(altTitle)}&limit=1`;
                await new Promise(resolve => setTimeout(resolve, 500)); // Задержка между попытками
                const altResponse = await makeRequest(altSearchUrl);
                
                if (altResponse.ok && altResponse.status !== 429) {
                    const altData = altResponse.json();
                    if (altData.data && altData.data.length > 0) {
                        const images = altData.data[0].images;
                        let imageUrl = getImageUrl(images, altData.data[0]);
                        if (imageUrl) {
                            console.log(`  ✓ Найдено по альтернативному названию: "${altTitle}"`);
                            return imageUrl;
                        }
                    }
                }
            }
            
            return null; // Не найдено ни по одному варианту
        }
        
        // Если результаты есть, ищем изображение
        const images = data.data[0].images;
        let imageUrl = getImageUrl(images, data.data[0]);
        
        if (imageUrl) {
            return imageUrl;
        }
        
        // Если изображение не найдено, но аниме найдено
        return null;
    } catch (error) {
        console.error(`Ошибка при получении изображения для "${animeTitle}":`, error.message);
        return null;
    }
}

// Функция для создания безопасного имени файла
function sanitizeFileName(name) {
    return name
        .replace(/[<>:"/\\|?*]/g, '_')
        .replace(/\s+/g, '_')
        .substring(0, 100);
}

// Основная функция
async function downloadImages() {
    // Ищем xlsx файл
    const files = fs.readdirSync(__dirname).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
    
    if (files.length === 0) {
        console.error('Не найден xlsx файл в текущей директории');
        return;
    }
    
    const xlsxFile = files[0];
    console.log(`Читаю файл: ${xlsxFile}`);
    
    // Читаем Excel файл
    const workbook = XLSX.readFile(xlsxFile);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Конвертируем в JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    if (jsonData.length === 0) {
        console.error('Файл пуст');
        return;
    }
    
    // Находим заголовки
    let headerRowIndex = 0;
    for (let i = 0; i < jsonData.length; i++) {
        if (jsonData[i].some(cell => cell !== '')) {
            headerRowIndex = i;
            break;
        }
    }
    
    const headers = jsonData[headerRowIndex] || [];
    
    // Находим колонку "anime"
    const animeColumnIndex = headers.findIndex(h => 
        String(h).toLowerCase().trim() === 'anime'
    );
    
    if (animeColumnIndex < 0) {
        console.error('Не найдена колонка "anime"');
        return;
    }
    
    // Получаем список уникальных аниме
    const dataRows = jsonData.slice(headerRowIndex + 1);
    const animeTitles = [...new Set(
        dataRows
            .map(row => row[animeColumnIndex])
            .filter(title => title && String(title).trim())
            .map(title => String(title).trim())
    )];
    
    console.log(`Найдено ${animeTitles.length} уникальных аниме\n`);
    
    let successCount = 0;
    let errorCount = 0;
    const downloadedFiles = {};
    
    // Загружаем изображения
    for (let i = 0; i < animeTitles.length; i++) {
        const animeTitle = animeTitles[i];
        console.log(`[${i + 1}/${animeTitles.length}] Обрабатываю: "${animeTitle}"...`);
        
        try {
            const imageUrl = await fetchAnimeImage(animeTitle);
            
            if (!imageUrl) {
                console.log(`  ❌ Изображение не найдено для "${animeTitle}"`);
                console.log(`     Проверьте правильность названия или добавьте изображение вручную\n`);
                errorCount++;
                continue;
            }
            
            // Определяем расширение файла
            const urlParts = new URL(imageUrl);
            let ext = path.extname(urlParts.pathname);
            if (!ext || ext === '') {
                ext = '.jpg'; // По умолчанию
            }
            
            // Создаем имя файла
            const safeName = sanitizeFileName(animeTitle);
            const filename = `${safeName}${ext}`;
            const filepath = path.join(imagesDir, filename);
            
            // Скачиваем файл
            await downloadFile(imageUrl, filepath);
            
            downloadedFiles[animeTitle] = filename;
            console.log(`  ✓ Скачано: ${filename}\n`);
            successCount++;
            
            // Задержка между запросами (увеличена до 500ms для соблюдения лимита API - 3 запроса в секунду)
            if (i < animeTitles.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 500));
            }
        } catch (error) {
            console.error(`  ❌ Ошибка при скачивании "${animeTitle}":`, error.message, '\n');
            errorCount++;
        }
    }
    
    // Сохраняем маппинг названий на имена файлов (для использования в веб-приложении)
    const mappingFile = path.join(__dirname, 'image-mapping.json');
    fs.writeFileSync(mappingFile, JSON.stringify(downloadedFiles, null, 2), 'utf8');
    console.log(`\nМаппинг сохранен в: ${mappingFile}`);
    
    // Также создаем обратный маппинг (имя файла -> название) для быстрого поиска
    const reverseMapping = {};
    Object.entries(downloadedFiles).forEach(([title, filename]) => {
        reverseMapping[filename] = title;
    });
    
    const reverseMappingFile = path.join(__dirname, 'image-mapping-reverse.json');
    fs.writeFileSync(reverseMappingFile, JSON.stringify(reverseMapping, null, 2), 'utf8');
    console.log(`Обратный маппинг сохранен в: ${reverseMappingFile}`);
    
    console.log(`\n=== Итого ===`);
    console.log(`Успешно скачано: ${successCount}`);
    console.log(`Ошибок: ${errorCount}`);
    console.log(`Всего: ${animeTitles.length}`);
}

// Запускаем
downloadImages().catch(error => {
    console.error('Критическая ошибка:', error);
    process.exit(1);
});

