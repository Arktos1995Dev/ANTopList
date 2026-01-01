let workbook = null;
let currentSheetName = null;
let currentView = 'table';
let animeImageCache = {}; // Кеш для изображений аниме
let localImageMapping = {}; // Маппинг локальных изображений

const uploadBox = document.getElementById('uploadBox');
const fileInput = document.getElementById('fileInput');
const controlsSection = document.getElementById('controlsSection');
const dataSection = document.getElementById('dataSection');
const infoSection = document.getElementById('infoSection');
const sheetSelect = document.getElementById('sheetSelect');
const dataTable = document.getElementById('dataTable');
const cardsContainer = document.getElementById('cardsContainer');
const sheetCountEl = document.getElementById('sheetCount');
const rowCountEl = document.getElementById('rowCount');
const colCountEl = document.getElementById('colCount');
const genreHint = document.getElementById('genreHint');
const imageControls = document.getElementById('imageControls');
const loadImagesBtn = document.getElementById('loadImagesBtn');

// Event listeners
uploadBox.addEventListener('click', () => fileInput.click());
uploadBox.addEventListener('dragover', handleDragOver);
uploadBox.addEventListener('dragleave', handleDragLeave);
uploadBox.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);

document.querySelectorAll('.btn-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.btn-toggle').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentView = btn.dataset.view;
        renderData();
    });
});

sheetSelect.addEventListener('change', () => {
    currentSheetName = sheetSelect.value;
    renderData();
});

if (loadImagesBtn) {
    loadImagesBtn.addEventListener('click', async () => {
        await loadAndRenderImages();
    });
}

function handleDragOver(e) {
    e.preventDefault();
    uploadBox.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    uploadBox.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    uploadBox.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

function processFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            // Включаем чтение стилей и гиперссылок
            workbook = XLSX.read(data, { 
                type: 'array',
                cellStyles: true,
                cellHTML: false
            });
            
            // Populate sheet selector
            sheetSelect.innerHTML = '';
            workbook.SheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelect.appendChild(option);
            });
            
            currentSheetName = workbook.SheetNames[0];
            sheetSelect.value = currentSheetName;
            
            // Show controls and data sections
            controlsSection.style.display = 'flex';
            dataSection.style.display = 'block';
            infoSection.style.display = 'block';
            
            // Update info
            updateInfo();
            
            // Render data
            renderData();
        } catch (error) {
            alert('Ошибка при чтении файла: ' + error.message);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function updateInfo() {
    if (!workbook || !currentSheetName) return;
    
    const worksheet = workbook.Sheets[currentSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    sheetCountEl.textContent = workbook.SheetNames.length;
    rowCountEl.textContent = jsonData.length;
    colCountEl.textContent = jsonData.length > 0 ? jsonData[0].length : 0;
}

async function renderData() {
    if (!workbook || !currentSheetName) return;
    
    if (currentView === 'table') {
        await renderTableView();
    } else {
        await renderCardsView();
    }
}

function getCellStyle(cellAddress, worksheet) {
    const cell = worksheet[cellAddress];
    if (!cell) return null;
    
    const style = {};
    const cellStyle = cell.s;
    
    if (cellStyle) {
        // Цвет фона (fill foreground)
        if (cellStyle.fg) {
            if (cellStyle.fg.rgb) {
                style.backgroundColor = '#' + cellStyle.fg.rgb;
            } else if (cellStyle.fg.theme !== undefined && cellStyle.fg.tint !== undefined) {
                // Для тематических цветов можно использовать базовые
                // Это упрощенная обработка, полная требует таблицы тем
            }
        }
        
        // Альтернативный способ чтения цвета фона
        if (cellStyle.fill && cellStyle.fill.fgColor) {
            if (cellStyle.fill.fgColor.rgb) {
                style.backgroundColor = '#' + cellStyle.fill.fgColor.rgb;
            }
        }
        
        // Цвет текста
        if (cellStyle.font) {
            if (cellStyle.font.color) {
                if (cellStyle.font.color.rgb) {
                    style.color = '#' + cellStyle.font.color.rgb;
                }
            }
            
            // Жирный текст
            if (cellStyle.font.bold) {
                style.fontWeight = 'bold';
            }
            
            // Курсив
            if (cellStyle.font.italic) {
                style.fontStyle = 'italic';
            }
            
            // Подчеркивание
            if (cellStyle.font.underline) {
                style.textDecoration = 'underline';
            }
        }
        
        // Выравнивание
        if (cellStyle.alignment) {
            if (cellStyle.alignment.horizontal) {
                const align = cellStyle.alignment.horizontal;
                style.textAlign = align === 'center' ? 'center' : 
                                 align === 'right' ? 'right' : 'left';
            }
        }
    }
    
    // Также проверяем прямой доступ к свойствам ячейки (для некоторых форматов)
    if (cell.z && !cellStyle) {
        // Если есть формат числа, но нет стиля, это может указывать на форматирование
    }
    
    return Object.keys(style).length > 0 ? style : null;
}

function getCellHyperlink(cellAddress, worksheet) {
    // Проверяем гиперссылки в листе (массив гиперссылок)
    if (worksheet['!hyperlinks'] && Array.isArray(worksheet['!hyperlinks'])) {
        const hyperlink = worksheet['!hyperlinks'].find(hl => hl.ref === cellAddress);
        if (hyperlink) {
            // Может быть target (полный URL) или location (якорь на листе)
            return hyperlink.target || hyperlink.location || null;
        }
    }
    
    // Также проверяем в самой ячейке (для некоторых форматов)
    const cell = worksheet[cellAddress];
    if (cell && cell.l) {
        // Гиперссылка может быть объектом или строкой
        if (typeof cell.l === 'string') {
            return cell.l;
        }
        if (cell.l.Target) {
            return cell.l.Target;
        }
        if (cell.l.target) {
            return cell.l.target;
        }
        if (cell.l.href) {
            return cell.l.href;
        }
    }
    
    // Проверяем, является ли само значение ячейки URL-ом
    if (cell && cell.v) {
        const value = String(cell.v);
        // Проверяем, похоже ли значение на URL
        const urlPattern = /^(https?:\/\/|www\.|mailto:|ftp:\/\/)/i;
        if (urlPattern.test(value.trim())) {
            // Если начинается с www., добавляем http://
            if (value.trim().toLowerCase().startsWith('www.')) {
                return 'http://' + value.trim();
            }
            return value.trim();
        }
        // Проверяем наличие домена
        const domainPattern = /^[a-zA-Z0-9][a-zA-Z0-9-]{1,61}[a-zA-Z0-9]\.[a-zA-Z]{2,}(\/.*)?$/;
        if (domainPattern.test(value.trim()) && !value.includes(' ')) {
            return 'http://' + value.trim();
        }
    }
    
    return null;
}

function styleToString(style) {
    if (!style) return '';
    return Object.entries(style).map(([key, value]) => {
        const cssKey = key.replace(/([A-Z])/g, '-$1').toLowerCase();
        return `${cssKey}: ${value}`;
    }).join('; ');
}

function getGenreColor(genre) {
    if (!genre) return null;
    
    const genreLower = String(genre).toLowerCase().trim();
    
    // Цветовая палитра для различных жанров (приглушенные, неконтрастные цвета)
    const genreColors = {
        'action': '#f5e6e6',           // Приглушенный красный
        'adventure': '#f5f0e1',        // Приглушенный желтый
        'comedy': '#f5f1d9',           // Приглушенный светло-желтый
        'drama': '#e5f0e5',            // Светло-зеленый
        'fantasy': '#e8eef5',          // Приглушенный синий
        'horror': '#f5eef3',           // Приглушенный розовый
        'romance': '#f5e8eb',          // Приглушенный розовый
        'sci-fi': '#e8f5f5',           // Приглушенный голубой
        'science fiction': '#e8f5f5',
        'thriller': '#f0eef5',         // Приглушенный светло-фиолетовый
        'documentary': '#e5f0e8',      // Приглушенный зеленый
        'animation': '#f5f2e1',        // Приглушенный желтый
        'mystery': '#eaeaea',          // Приглушенный серый
        'crime': '#eeeeee',            // Приглушенный светло-серый
        'family': '#f5f0e1',           // Приглушенный бежевый
        'western': '#f5ede1',          // Приглушенный персиковый
        'war': '#f5e8e8',              // Приглушенный красный
        'musical': '#f5e8d9',          // Приглушенный золотистый
        'biography': '#e8eef5',        // Приглушенный светло-синий
        'history': '#efeef5',          // Приглушенный лавандовый
        'sport': '#e5f0e8',            // Приглушенный мятный
        'superhero': '#e8eef5',        // Приглушенный синий
        'super hero': '#e8eef5',
    };
    
    // Ищем точное совпадение (без учета регистра)
    if (genreColors.hasOwnProperty(genreLower)) {
        return genreColors[genreLower];
    }
    
    // Если точного совпадения нет, генерируем цвет на основе хэша
    return generateColorFromString(genreLower);
}

function generateColorFromString(str) {
    // Генерируем цвет на основе строки для неизвестных жанров
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    
    // Генерируем пастельный цвет (HSL: светлота > 85%)
    const hue = Math.abs(hash % 360);
    const saturation = 40 + (Math.abs(hash) % 30); // 40-70%
    const lightness = 85 + (Math.abs(hash) % 10);  // 85-95%
    
    return `hsl(${hue}, ${saturation}%, ${lightness}%)`;
}

// Функция для извлечения URL изображения из данных API
function getImageUrlFromData(animeData) {
    if (!animeData || !animeData.images) return null;
    
    const images = animeData.images;
    let imageUrl = null;
    
    if (images.jpg && images.jpg.image_url) {
        imageUrl = images.jpg.image_url;
    } else if (images.webp && images.webp.image_url) {
        imageUrl = images.webp.image_url;
    } else if (images.image_url) {
        imageUrl = images.image_url;
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

// Загружаем маппинг локальных изображений
async function loadLocalImageMapping() {
    try {
        const response = await fetch('image-mapping.json');
        if (response.ok) {
            localImageMapping = await response.json();
            console.log('Локальный маппинг изображений загружен:', Object.keys(localImageMapping).length, 'файлов');
        }
    } catch (error) {
        console.log('Локальный маппинг изображений не найден, используем API');
    }
}

async function fetchAnimeImage(animeTitle, retryCount = 0) {
    if (!animeTitle || typeof animeTitle !== 'string') {
        return { error: 'Неверное название аниме' };
    }
    
    const title = animeTitle.trim();
    if (!title) {
        return { error: 'Пустое название аниме' };
    }
    
    // Проверяем кеш
    if (animeImageCache[title]) {
        return animeImageCache[title];
    }
    
    // Проверяем локальные изображения
    if (localImageMapping[title]) {
        const localPath = `images/${localImageMapping[title]}`;
        animeImageCache[title] = { url: localPath };
        return { url: localPath };
    }
    
    const maxRetries = 3;
    const baseDelay = 2000; // 2 секунды базовая задержка
    
    try {
        // Используем Jikan API (MyAnimeList) для поиска аниме
        const searchUrl = `https://api.jikan.moe/v4/anime?q=${encodeURIComponent(title)}&limit=1`;
        const response = await fetch(searchUrl);
        
        // Обработка ошибки 429 (Too Many Requests)
        if (response.status === 429) {
            if (retryCount < maxRetries) {
                const delay = baseDelay * Math.pow(2, retryCount); // Экспоненциальная задержка
                console.log(`Лимит запросов превышен для "${title}", ожидание ${delay/1000} сек...`);
                await new Promise(resolve => setTimeout(resolve, delay));
                return fetchAnimeImage(animeTitle, retryCount + 1);
            } else {
                const errorMsg = `Ошибка API 429: Превышен лимит запросов после ${maxRetries} попыток`;
                animeImageCache[title] = { error: errorMsg };
                return { error: errorMsg };
            }
        }
        
        if (!response.ok) {
            const errorMsg = `Ошибка API: ${response.status} ${response.statusText}`;
            animeImageCache[title] = { error: errorMsg };
            return { error: errorMsg };
        }
        
        const data = await response.json();
        
        // Проверяем, есть ли результаты
        if (!data.data || data.data.length === 0) {
            // Пробуем альтернативные варианты поиска
            const alternatives = generateSearchAlternatives(title);
            for (const altTitle of alternatives) {
                if (altTitle === title) continue; // Уже проверили
                
                await new Promise(resolve => setTimeout(resolve, 500)); // Задержка между попытками
                const altSearchUrl = `https://api.jikan.moe/v4/anime?q=${encodeURIComponent(altTitle)}&limit=1`;
                const altResponse = await fetch(altSearchUrl);
                
                if (altResponse.ok && altResponse.status !== 429) {
                    const altData = await altResponse.json();
                    if (altData.data && altData.data.length > 0) {
                        const imageUrl = getImageUrlFromData(altData.data[0]);
                        if (imageUrl) {
                            animeImageCache[title] = { url: imageUrl };
                            return { url: imageUrl };
                        }
                    }
                }
            }
            
            const errorMsg = `Аниме "${title}" не найдено в базе MyAnimeList`;
            animeImageCache[title] = { error: errorMsg };
            return { error: errorMsg };
        }
        
        // Если результаты есть, ищем изображение
        const imageUrl = getImageUrlFromData(data.data[0]);
        
        if (imageUrl) {
            animeImageCache[title] = { url: imageUrl };
            return { url: imageUrl };
        }
        
        // Если изображение не найдено, но аниме найдено
        const errorMsg = `Изображение для "${title}" недоступно`;
        animeImageCache[title] = { error: errorMsg };
        return { error: errorMsg };
    } catch (error) {
        let errorMsg = 'Неизвестная ошибка';
        if (error instanceof TypeError && error.message.includes('fetch')) {
            errorMsg = 'Ошибка сети: не удалось подключиться к API';
        } else if (error.message) {
            errorMsg = `Ошибка: ${error.message}`;
        }
        
        console.log(`Error fetching image for "${title}":`, error);
        animeImageCache[title] = { error: errorMsg };
        return { error: errorMsg };
    }
}

async function loadAnimeImages(animeTitles) {
    // Загружаем изображения для всех уникальных аниме
    const uniqueTitles = [...new Set(animeTitles.filter(t => t && t.trim()))];
    let successCount = 0;
    let errorCount = 0;
    
    // Загружаем с задержкой, чтобы не превысить лимит API (3 запроса в секунду)
    // Увеличена задержка до 500ms для большей надежности
    for (let i = 0; i < uniqueTitles.length; i++) {
        const result = await fetchAnimeImage(uniqueTitles[i]);
        if (result && result.url) {
            successCount++;
        } else if (result && result.error) {
            errorCount++;
        }
        // Задержка 500ms между запросами (для соблюдения лимита 3 запроса в секунду)
        if (i < uniqueTitles.length - 1) {
            await new Promise(resolve => setTimeout(resolve, 500));
        }
    }
    
    return { successCount, errorCount, total: uniqueTitles.length };
}

async function loadAndRenderImages() {
    if (!workbook || !currentSheetName) return;
    
    const worksheet = workbook.Sheets[currentSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    if (jsonData.length === 0) return;
    
    // Определяем заголовки
    let headerRowIndex = 0;
    for (let i = 0; i < jsonData.length; i++) {
        if (jsonData[i].some(cell => cell !== '')) {
            headerRowIndex = i;
            break;
        }
    }
    
    const headers = jsonData[headerRowIndex] || [];
    const dataRows = jsonData.slice(headerRowIndex + 1);
    
    // Находим индекс колонки "anime"
    const animeColumnIndex = headers.findIndex(h => 
        String(h).toLowerCase().trim() === 'anime'
    );
    
    if (animeColumnIndex < 0) return;
    
    const animeTitles = dataRows
        .map(row => row[animeColumnIndex])
        .filter(title => title && title.trim());
    
    if (animeTitles.length === 0) return;
    
    // Отключаем кнопку и показываем загрузку
    if (loadImagesBtn) {
        loadImagesBtn.disabled = true;
        const btnText = loadImagesBtn.querySelector('.btn-text');
        if (btnText) {
            btnText.textContent = 'Загрузка...';
        }
    }
    
    try {
        // Загружаем изображения
        const result = await loadAnimeImages(animeTitles);
        
        // Перерисовываем данные с изображениями
        await renderData();
        
        // Обновляем текст кнопки с информацией о результатах
        if (loadImagesBtn) {
            const btnText = loadImagesBtn.querySelector('.btn-text');
            if (btnText) {
                if (result.errorCount > 0) {
                    btnText.textContent = `Загружено: ${result.successCount}/${result.total} (ошибок: ${result.errorCount})`;
                } else {
                    btnText.textContent = `Загружено: ${result.successCount}/${result.total} ✓`;
                }
            }
            loadImagesBtn.disabled = false;
            loadImagesBtn.style.opacity = '0.7';
        }
    } catch (error) {
        console.error('Error loading images:', error);
        if (loadImagesBtn) {
            const btnText = loadImagesBtn.querySelector('.btn-text');
            if (btnText) {
                const errorMsg = error.message || 'Неизвестная ошибка';
                btnText.textContent = `Ошибка: ${errorMsg}`;
            }
            loadImagesBtn.disabled = false;
        }
    }
}

async function renderTableView() {
    const worksheet = workbook.Sheets[currentSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    if (jsonData.length === 0) {
        dataTable.innerHTML = '<tr><td colspan="100" style="text-align: center; padding: 40px;">Нет данных</td></tr>';
        return;
    }
    
    // Determine header row (first non-empty row)
    let headerRowIndex = 0;
    for (let i = 0; i < jsonData.length; i++) {
        if (jsonData[i].some(cell => cell !== '')) {
            headerRowIndex = i;
            break;
        }
    }
    
    const headers = jsonData[headerRowIndex] || [];
    const dataRows = jsonData.slice(headerRowIndex + 1);
    
    // Находим индекс колонки "genre" (без учета регистра)
    const genreColumnIndex = headers.findIndex(h => 
        String(h).toLowerCase().trim() === 'genre'
    );
    
    // Находим индекс колонки "anime" (без учета регистра)
    const animeColumnIndex = headers.findIndex(h => 
        String(h).toLowerCase().trim() === 'anime'
    );
    
    // Показываем кнопку загрузки изображений, если найдена колонка anime
    if (animeColumnIndex >= 0 && imageControls) {
        imageControls.style.display = 'block';
    } else if (imageControls) {
        imageControls.style.display = 'none';
    }
    
    // Получаем диапазон ячеек
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Create table
    let html = '<thead><tr>';
    
    // Добавляем первую колонку
    if (headers.length > 0) {
        const firstCellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: 0 });
        const firstCellStyle = getCellStyle(firstCellAddress, worksheet);
        const firstStyleAttr = firstCellStyle ? ` style="${styleToString(firstCellStyle)}"` : '';
        html += `<th${firstStyleAttr}>${escapeHtml(headers[0] || '')}</th>`;
    }
    
    // Вставляем колонку Image второй, если есть колонка anime
    if (animeColumnIndex >= 0) {
        html += '<th>Image</th>';
    }
    
    // Добавляем остальные колонки (начиная со второй, если есть image, иначе с первой)
    headers.forEach((header, colIndex) => {
        if (colIndex === 0) return; // Первая колонка уже добавлена
        
        const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: colIndex });
        const cellStyle = getCellStyle(cellAddress, worksheet);
        const styleAttr = cellStyle ? ` style="${styleToString(cellStyle)}"` : '';
        html += `<th${styleAttr}>${escapeHtml(header || '')}</th>`;
    });
    
    html += '</tr></thead><tbody>';
    
    dataRows.forEach((row, rowIndex) => {
        if (row.some(cell => cell !== '')) {
            // Получаем значение genre для этой строки
            const genreValue = genreColumnIndex >= 0 && genreColumnIndex < row.length 
                ? row[genreColumnIndex] 
                : null;
            const genreColor = getGenreColor(genreValue);
            const rowStyle = genreColor ? ` style="background-color: ${genreColor};"` : '';
            
            html += `<tr${rowStyle} class="genre-row">`;
            
            // Добавляем первую ячейку
            if (headers.length > 0) {
                const firstCellValue = row[0] || '';
                const actualRowIndex = headerRowIndex + 1 + rowIndex;
                const firstCellAddress = XLSX.utils.encode_cell({ r: actualRowIndex, c: 0 });
                const firstCellStyle = getCellStyle(firstCellAddress, worksheet);
                const firstHyperlink = getCellHyperlink(firstCellAddress, worksheet);
                
                let firstCellContent = escapeHtml(String(firstCellValue));
                if (firstHyperlink) {
                    let url = firstHyperlink;
                    if (typeof firstHyperlink === 'string') {
                        const linkLower = firstHyperlink.toLowerCase().trim();
                        if (linkLower.startsWith('http://') || 
                            linkLower.startsWith('https://') || 
                            linkLower.startsWith('ftp://') || 
                            linkLower.startsWith('mailto:') ||
                            linkLower.startsWith('#')) {
                            url = firstHyperlink.trim();
                        } else if (linkLower.startsWith('www.')) {
                            url = 'http://' + firstHyperlink.trim();
                        } else if (firstHyperlink.trim().length > 0) {
                            const domainPattern = /^[a-zA-Z0-9][a-zA-Z0-9-_.]*\.[a-zA-Z]{2,}(\/.*)?$/;
                            if (domainPattern.test(firstHyperlink.trim()) && !firstHyperlink.includes(' ')) {
                                url = 'http://' + firstHyperlink.trim();
                            }
                        }
                    }
                    const linkText = firstCellContent || escapeHtml(url);
                    firstCellContent = `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" class="cell-link">${linkText}</a>`;
                }
                
                const firstStyleAttr = firstCellStyle ? ` style="${styleToString(firstCellStyle)}"` : '';
                html += `<td${firstStyleAttr}>${firstCellContent}</td>`;
            }
            
            // Вставляем ячейку с изображением второй, если есть колонка anime
            if (animeColumnIndex >= 0) {
                const animeTitle = row[animeColumnIndex] || '';
                const imageData = animeTitle ? animeImageCache[animeTitle.trim()] : null;
                
                if (imageData && imageData.url) {
                    html += `<td class="anime-image-cell"><img src="${escapeHtml(imageData.url)}" alt="${escapeHtml(animeTitle)}" class="anime-image" loading="lazy" onerror="this.onerror=null; this.parentElement.innerHTML='<span class=\\'image-error\\'>Ошибка загрузки изображения</span>';"></td>`;
                } else if (imageData && imageData.error) {
                    html += `<td class="anime-image-cell"><span class="image-error" title="${escapeHtml(imageData.error)}">${escapeHtml(imageData.error)}</span></td>`;
                } else {
                    html += `<td class="anime-image-cell">—</td>`;
                }
            }
            
            // Добавляем остальные ячейки (начиная со второй)
            headers.forEach((_, colIndex) => {
                if (colIndex === 0) return; // Первая ячейка уже добавлена
                
                const cellValue = row[colIndex] || '';
                const actualRowIndex = headerRowIndex + 1 + rowIndex;
                const cellAddress = XLSX.utils.encode_cell({ r: actualRowIndex, c: colIndex });
                const cellStyle = getCellStyle(cellAddress, worksheet);
                const hyperlink = getCellHyperlink(cellAddress, worksheet);
                
                let cellContent = escapeHtml(String(cellValue));
                if (hyperlink) {
                    let url = hyperlink;
                    // Обработка различных форматов ссылок
                    if (typeof hyperlink === 'string') {
                        const linkLower = hyperlink.toLowerCase().trim();
                        // Если ссылка уже полная (http/https/ftp/mailto), используем как есть
                        if (linkLower.startsWith('http://') || 
                            linkLower.startsWith('https://') || 
                            linkLower.startsWith('ftp://') || 
                            linkLower.startsWith('mailto:') ||
                            linkLower.startsWith('#')) {
                            url = hyperlink.trim();
                        } else if (linkLower.startsWith('www.')) {
                            // Если начинается с www., добавляем http://
                            url = 'http://' + hyperlink.trim();
                        } else if (hyperlink.trim().length > 0) {
                            // Другие случаи - пытаемся определить, это домен или нет
                            const domainPattern = /^[a-zA-Z0-9][a-zA-Z0-9-_.]*\.[a-zA-Z]{2,}(\/.*)?$/;
                            if (domainPattern.test(hyperlink.trim()) && !hyperlink.includes(' ')) {
                                url = 'http://' + hyperlink.trim();
                            }
                        }
                    }
                    // Используем текст ссылки или URL, если текст пустой
                    const linkText = cellContent || escapeHtml(url);
                    cellContent = `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" class="cell-link">${linkText}</a>`;
                }
                
                const styleAttr = cellStyle ? ` style="${styleToString(cellStyle)}"` : '';
                html += `<td${styleAttr}>${cellContent}</td>`;
            });
            
            html += '</tr>';
        }
    });
    
    html += '</tbody>';
    dataTable.innerHTML = html;
    
    // Show table view, hide cards view
    document.getElementById('tableView').style.display = 'block';
    document.getElementById('cardsView').style.display = 'none';
    
    // Показываем подсказку о цветовом выделении, если есть колонка genre
    showGenreColorHint(genreColumnIndex >= 0);
}

async function renderCardsView() {
    const worksheet = workbook.Sheets[currentSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    if (jsonData.length === 0) {
        cardsContainer.innerHTML = '<div style="text-align: center; padding: 40px; grid-column: 1/-1;">Нет данных</div>';
        return;
    }
    
    // Determine header row
    let headerRowIndex = 0;
    for (let i = 0; i < jsonData.length; i++) {
        if (jsonData[i].some(cell => cell !== '')) {
            headerRowIndex = i;
            break;
        }
    }
    
    const headers = jsonData[headerRowIndex] || [];
    const dataRows = jsonData.slice(headerRowIndex + 1).filter(row => row.some(cell => cell !== ''));
    
    // Находим индекс колонки "genre" (без учета регистра)
    const genreColumnIndex = headers.findIndex(h => 
        String(h).toLowerCase().trim() === 'genre'
    );
    
    // Находим индекс колонки "anime" (без учета регистра)
    const animeColumnIndex = headers.findIndex(h => 
        String(h).toLowerCase().trim() === 'anime'
    );
    
    // Показываем кнопку загрузки изображений, если найдена колонка anime
    if (animeColumnIndex >= 0 && imageControls) {
        imageControls.style.display = 'block';
    } else if (imageControls) {
        imageControls.style.display = 'none';
    }
    
    // Получаем диапазон ячеек
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Create cards
    let html = '';
    dataRows.forEach((row, rowIndex) => {
        const actualRowIndex = headerRowIndex + 1 + rowIndex;
        
        // Получаем значение genre для этой строки
        const genreValue = genreColumnIndex >= 0 && genreColumnIndex < row.length 
            ? row[genreColumnIndex] 
            : null;
        const genreColor = getGenreColor(genreValue);
        const cardStyle = genreColor ? ` style="background-color: ${genreColor};"` : '';
        
        html += `<div class="card genre-card"${cardStyle}>`;
        
        // Добавляем изображение аниме в начало карточки, если есть
        if (animeColumnIndex >= 0) {
            const animeTitle = row[animeColumnIndex] || '';
            const imageData = animeTitle ? animeImageCache[animeTitle.trim()] : null;
            
            if (imageData && imageData.url) {
                html += `<div class="card-image-wrapper"><img src="${escapeHtml(imageData.url)}" alt="${escapeHtml(animeTitle)}" class="card-anime-image" loading="lazy" onerror="this.onerror=null; this.parentElement.innerHTML='<span class=\\'image-error\\'>Ошибка загрузки изображения</span>';"></div>`;
            } else if (imageData && imageData.error) {
                html += `<div class="card-image-wrapper"><span class="image-error" title="${escapeHtml(imageData.error)}">${escapeHtml(imageData.error)}</span></div>`;
            }
        }
        headers.forEach((header, colIndex) => {
            const value = row[colIndex] || '';
            if (header) {
                const cellAddress = XLSX.utils.encode_cell({ r: actualRowIndex, c: colIndex });
                const cellStyle = getCellStyle(cellAddress, worksheet);
                const hyperlink = getCellHyperlink(cellAddress, worksheet);
                
                let valueContent = escapeHtml(String(value));
                if (hyperlink) {
                    let url = hyperlink;
                    // Обработка различных форматов ссылок
                    if (typeof hyperlink === 'string') {
                        const linkLower = hyperlink.toLowerCase().trim();
                        // Если ссылка уже полная (http/https/ftp/mailto), используем как есть
                        if (linkLower.startsWith('http://') || 
                            linkLower.startsWith('https://') || 
                            linkLower.startsWith('ftp://') || 
                            linkLower.startsWith('mailto:') ||
                            linkLower.startsWith('#')) {
                            url = hyperlink.trim();
                        } else if (linkLower.startsWith('www.')) {
                            // Если начинается с www., добавляем http://
                            url = 'http://' + hyperlink.trim();
                        } else if (hyperlink.trim().length > 0) {
                            // Другие случаи - пытаемся определить, это домен или нет
                            const domainPattern = /^[a-zA-Z0-9][a-zA-Z0-9-_.]*\.[a-zA-Z]{2,}(\/.*)?$/;
                            if (domainPattern.test(hyperlink.trim()) && !hyperlink.includes(' ')) {
                                url = 'http://' + hyperlink.trim();
                            }
                        }
                    }
                    // Используем текст ссылки или URL, если текст пустой
                    const linkText = valueContent || escapeHtml(url);
                    valueContent = `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" class="cell-link">${linkText}</a>`;
                }
                
                const styleAttr = cellStyle ? ` style="${styleToString(cellStyle)}"` : '';
                html += `
                    <div class="card-row">
                        <span class="card-label">${escapeHtml(header)}</span>
                        <span class="card-value"${styleAttr}>${valueContent}</span>
                    </div>
                `;
            }
        });
        html += '</div>';
    });
    
    cardsContainer.innerHTML = html;
    
    // Show cards view, hide table view
    document.getElementById('tableView').style.display = 'none';
    document.getElementById('cardsView').style.display = 'block';
    
    // Показываем подсказку о цветовом выделении, если есть колонка genre
    showGenreColorHint(genreColumnIndex >= 0);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function showGenreColorHint(show) {
    if (genreHint) {
        genreHint.style.display = show ? 'flex' : 'none';
    }
}

// Auto-load the Excel file if it's in the same directory
window.addEventListener('DOMContentLoaded', async () => {
    // Загружаем маппинг локальных изображений
    await loadLocalImageMapping();
    // Try to load QD Design Doc.xlsx if it exists
    fetch('QD Design Doc.xlsx')
        .then(response => {
            if (response.ok) {
                return response.arrayBuffer();
            }
            throw new Error('File not found');
        })
        .then(data => {
            const uint8Array = new Uint8Array(data);
            // Включаем чтение стилей и гиперссылок
            workbook = XLSX.read(uint8Array, { 
                type: 'array',
                cellStyles: true,
                cellHTML: false
            });
            
            // Populate sheet selector
            sheetSelect.innerHTML = '';
            workbook.SheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelect.appendChild(option);
            });
            
            currentSheetName = workbook.SheetNames[0];
            sheetSelect.value = currentSheetName;
            
            // Show controls and data sections
            controlsSection.style.display = 'flex';
            dataSection.style.display = 'block';
            infoSection.style.display = 'block';
            
            // Update info
            updateInfo();
            
            // Render data
            renderData();
        })
        .catch(() => {
            // File not found or couldn't be loaded, user will need to upload
            console.log('Excel file not found, waiting for user upload...');
        });
});

