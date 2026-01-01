document.addEventListener('DOMContentLoaded', async () => {
    await fetchVersion();
    // ... остальная логика DOMContentLoaded ...
});

async function fetchVersion() {
    try {
        const response = await fetch('/api/version');
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        const data = await response.json();
        const versionInfoEl = document.getElementById('version-info');
        if (versionInfoEl) {
            versionInfoEl.textContent = `Версия: ${data.version}`;
        }
    } catch (error) {
        console.error('Failed to fetch version:', error);
        const versionInfoEl = document.getElementById('version-info');
        if (versionInfoEl) {
            versionInfoEl.textContent = 'Версия: не удалось загрузить';
        }
    }
}

let workbook = null;
let currentSheetName = null;
let currentView = 'table';
let localImageMapping = {}; // Маппинг для локальных изображений, полученный с сервера
let imagesHidden = false;

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
const loadingIndicator = document.getElementById('loadingIndicator');
const hideImagesButton = document.getElementById('hideImagesButton');
const loadMalListButton = document.getElementById('loadMalList');
const malUsernameInput = document.getElementById('malUsername');

// Event listeners
uploadBox.addEventListener('click', () => fileInput.click());
uploadBox.addEventListener('dragover', (e) => { e.preventDefault(); uploadBox.classList.add('dragover'); });
uploadBox.addEventListener('dragleave', (e) => { e.preventDefault(); uploadBox.classList.remove('dragover'); });
uploadBox.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);
loadMalListButton.addEventListener('click', handleMalListLoad);

hideImagesButton.addEventListener('click', () => {
    imagesHidden = !imagesHidden;
    document.body.classList.toggle('hide-images');
    hideImagesButton.textContent = imagesHidden ? 'Show Images' : 'Hide Images';
});

document.querySelectorAll('.btn-toggle[data-view]').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.btn-toggle[data-view]').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentView = btn.dataset.view;
        renderData();
    });
});

sheetSelect.addEventListener('change', () => {
    currentSheetName = sheetSelect.value;
    renderData();
    // Повторно инициировать загрузку для нового листа
    initiateImageDownloads(); 
});

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

async function handleMalListLoad() {
    const username = malUsernameInput.value.trim();
    if (!username) {
        alert('Please enter a MyAnimeList username.');
        return;
    }

    if(loadingIndicator) loadingIndicator.style.display = 'flex';

    try {
        const response = await fetch('/api/mal-list', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username })
        });

        if (!response.ok) {
            throw new Error(`Server error: ${response.statusText}`);
        }

        localImageMapping = await response.json();
        console.log('Image mapping received from server:', localImageMapping);

        const animeTitles = Object.keys(localImageMapping);
        const jsonData = [['Anime'], ...animeTitles.map(title => [title])];
        
        // Create a fake workbook and sheet
        workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet(jsonData);
        XLSX.utils.book_append_sheet(workbook, worksheet, "MyAnimeList");

        currentSheetName = "MyAnimeList";
        sheetSelect.innerHTML = `<option value="${currentSheetName}">${currentSheetName}</option>`;

        controlsSection.style.display = 'flex';
        dataSection.style.display = 'block';
        infoSection.style.display = 'block';
        uploadBox.style.display = 'none';

        updateInfo();
        await renderData();

    } catch (error) {
        console.error('Error loading MAL list:', error);
    } finally {
        if(loadingIndicator) loadingIndicator.style.display = 'none';
    }
}

function processFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array', cellStyles: true });
            
            sheetSelect.innerHTML = '';
            workbook.SheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelect.appendChild(option);
            });
            
            currentSheetName = workbook.SheetNames[0];
            
            controlsSection.style.display = 'flex';
            dataSection.style.display = 'block';
            infoSection.style.display = 'block';
            uploadBox.style.display = 'none'; // Скрыть область загрузки
            
            updateInfo();
            renderData(); // Первичная отрисовка без картинок
            initiateImageDownloads(); // Запуск процесса скачивания на сервере

        } catch (error) {
            alert('Ошибка при чтении файла: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

async function initiateImageDownloads() {
    if (!workbook || !currentSheetName) return;

    const worksheet = workbook.Sheets[currentSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    let headerRowIndex = 0;
    for(let i=0; i < jsonData.length; i++) {
        if (jsonData[i].some(cell => cell !== '')) {
            headerRowIndex = i; break;
        }
    }
    const headers = jsonData[headerRowIndex] || [];
    const animeColumnIndex = headers.findIndex(h => String(h).toLowerCase().trim() === 'anime');

    if (animeColumnIndex < 0) {
        console.log('Колонка "anime" не найдена, скачивание изображений не требуется.');
        return;
    }

    const animeTitles = [...new Set(jsonData.slice(headerRowIndex + 1)
        .map(row => row[animeColumnIndex])
        .filter(title => title && String(title).trim()))];

    if (animeTitles.length === 0) return;

    if(loadingIndicator) loadingIndicator.style.display = 'flex';

    try {
        const response = await fetch('/download-images', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ animeTitles })
        });

        if (!response.ok) {
            throw new Error(`Ошибка сервера: ${response.statusText}`);
        }

        localImageMapping = await response.json();
        console.log('Маппинг изображений получен с сервера:', localImageMapping);

        // Перерисовываем данные с уже загруженными изображениями
        await renderData();

    } catch (error) {
        console.error('Ошибка при запросе на скачивание изображений:', error);
        // Можно показать сообщение об ошибке пользователю
    } finally {
        if(loadingIndicator) loadingIndicator.style.display = 'none';
    }
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
    if (!cell || !cell.s) return null;
    const style = {};
    if (cell.s.fill && cell.s.fill.fgColor && cell.s.fill.fgColor.rgb) {
        style.backgroundColor = '#' + cell.s.fill.fgColor.rgb;
    }
    if (cell.s.font) {
        if (cell.s.font.color && cell.s.font.color.rgb) style.color = '#' + cell.s.font.color.rgb;
        if (cell.s.font.bold) style.fontWeight = 'bold';
        if (cell.s.font.italic) style.fontStyle = 'italic';
        if (cell.s.font.underline) style.textDecoration = 'underline';
    }
    if (cell.s.alignment) {
        if (cell.s.alignment.horizontal) style.textAlign = cell.s.alignment.horizontal;
    }
    return Object.keys(style).length > 0 ? style : null;
}

function getCellHyperlink(cellAddress, worksheet) {
    const cell = worksheet[cellAddress];
    if (worksheet['!hyperlinks']) {
        const hyperlink = worksheet['!hyperlinks'].find(hl => hl.ref === cellAddress);
        if (hyperlink) return hyperlink.target || hyperlink.location;
    }
    return cell && cell.l ? (cell.l.Target || cell.l.target || cell.l.href) : null;
}

function styleToString(style) {
    if (!style) return '';
    return Object.entries(style).map(([key, value]) => `${key.replace(/([A-Z])/g, '-\1').toLowerCase()}: ${value}`).join('; ');
}

function getGenreColor(genre) {
    if (!genre) return null;
    const genreLower = String(genre).toLowerCase().trim();
    const genreColors = {
        'action': '#f5e6e6', 'adventure': '#f5f0e1', 'comedy': '#f5f1d9', 'drama': '#e5f0e5', 'fantasy': '#e8eef5', 
        'horror': '#f5eef3', 'romance': '#f5e8eb', 'sci-fi': '#e8f5f5', 'thriller': '#f0eef5', 'mystery': '#eaeaea'
    };
    return genreColors[genreLower] || generateColorFromString(genreLower);
}

function generateColorFromString(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) hash = str.charCodeAt(i) + ((hash << 5) - hash);
    const hue = Math.abs(hash % 360);
    return `hsl(${hue}, 60%, 90%)`;
}

function renderCellContent(value, hyperlink) {
    const escapedValue = escapeHtml(String(value));
    if (!hyperlink) return escapedValue;
    return `<a href="${escapeHtml(hyperlink)}" target="_blank" rel="noopener noreferrer" class="cell-link">${escapedValue || escapeHtml(hyperlink)}</a>`;
}

async function renderTableView() {
    const worksheet = workbook.Sheets[currentSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    if (jsonData.length === 0) {
        dataTable.innerHTML = '<tr><td colspan="100">Нет данных</td></tr>'; return;
    }

    let headerRowIndex = 0;
    for(let i=0; i < jsonData.length; i++) {
        if (jsonData[i].some(cell => cell !== '')) { headerRowIndex = i; break; }
    }
    const headers = jsonData[headerRowIndex] || [];
    const dataRows = jsonData.slice(headerRowIndex + 1);
    const genreColumnIndex = headers.findIndex(h => String(h).toLowerCase().trim() === 'genre');
    const animeColumnIndex = headers.findIndex(h => String(h).toLowerCase().trim() === 'anime');

    let html = '<thead><tr>';
    if (animeColumnIndex >= 0) html += '<th class="image-col-header">Image</th>';
    headers.forEach((header, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: colIndex });
        const styleAttr = styleToString(getCellStyle(cellAddress, worksheet));
        html += `<th style="${styleAttr}">${escapeHtml(header || '')}</th>`;
    });
    html += '</tr></thead><tbody>';

    dataRows.forEach((row, rowIndex) => {
        if (!row.some(cell => cell !== '')) return;
        const genreValue = genreColumnIndex >= 0 ? row[genreColumnIndex] : null;
        const rowStyle = getGenreColor(genreValue) ? ` style="background-color: ${getGenreColor(genreValue)};"` : '';
        html += `<tr${rowStyle}>`;

        if (animeColumnIndex >= 0) {
            const animeTitle = row[animeColumnIndex] || '';
            const imageName = localImageMapping[animeTitle.trim()];
            html += `<td class="anime-image-cell image-col">`;
            if (imageName) {
                html += `<img src="images/${escapeHtml(imageName)}" alt="${escapeHtml(animeTitle)}" class="anime-image" loading="lazy">`;
            }
            html += `</td>`;
        }

        row.forEach((cellValue, colIndex) => {
            const actualRowIndex = headerRowIndex + 1 + rowIndex;
            const cellAddress = XLSX.utils.encode_cell({ r: actualRowIndex, c: colIndex });
            const hyperlink = getCellHyperlink(cellAddress, worksheet);
            const styleAttr = styleToString(getCellStyle(cellAddress, worksheet));
            html += `<td style="${styleAttr}">${renderCellContent(cellValue, hyperlink)}</td>`;
        });
        html += '</tr>';
    });

    dataTable.innerHTML = html + '</tbody>';
    document.getElementById('tableView').style.display = 'block';
    document.getElementById('cardsView').style.display = 'none';
    showGenreColorHint(genreColumnIndex >= 0);
}

async function renderCardsView() {
    const worksheet = workbook.Sheets[currentSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    if (jsonData.length === 0) {
        cardsContainer.innerHTML = '<p>Нет данных</p>'; return;
    }
    
    let headerRowIndex = 0;
    for(let i=0; i < jsonData.length; i++) {
        if (jsonData[i].some(cell => cell !== '')) { headerRowIndex = i; break; }
    }
    const headers = jsonData[headerRowIndex] || [];
    const dataRows = jsonData.slice(headerRowIndex + 1).filter(row => row.some(cell => cell !== ''));
    const genreColumnIndex = headers.findIndex(h => String(h).toLowerCase().trim() === 'genre');
    const animeColumnIndex = headers.findIndex(h => String(h).toLowerCase().trim() === 'anime');

    let html = '';
    dataRows.forEach((row, rowIndex) => {
        const genreValue = genreColumnIndex >= 0 ? row[genreColumnIndex] : null;
        const cardStyle = getGenreColor(genreValue) ? ` style="background-color: ${getGenreColor(genreValue)};"` : '';
        html += `<div class="card"${cardStyle}>`;
        
        const animeTitle = animeColumnIndex >= 0 ? (row[animeColumnIndex] || '') : '';
        const imageName = localImageMapping[animeTitle.trim()];
        if (imageName) {
            html += `<div class="card-image-wrapper"><img src="images/${escapeHtml(imageName)}" alt="${escapeHtml(animeTitle)}" class="card-anime-image" loading="lazy"></div>`;
        }

        headers.forEach((header, colIndex) => {
            if (!header) return;
            const actualRowIndex = headerRowIndex + 1 + rowIndex;
            const cellAddress = XLSX.utils.encode_cell({ r: actualRowIndex, c: colIndex });
            const cellValue = row[colIndex] || '';
            const hyperlink = getCellHyperlink(cellAddress, worksheet);
            const styleAttr = styleToString(getCellStyle(cellAddress, worksheet));
            html += `<div class="card-row"><span class="card-label">${escapeHtml(header)}</span><span class="card-value" style="${styleAttr}">${renderCellContent(cellValue, hyperlink)}</span></div>`;
        });
        html += '</div>';
    });

    cardsContainer.innerHTML = html;
    document.getElementById('tableView').style.display = 'none';
    document.getElementById('cardsView').style.display = 'block';
    showGenreColorHint(genreColumnIndex >= 0);
}

function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
}

function showGenreColorHint(show) {
    if (genreHint) {
        genreHint.style.display = show ? 'flex' : 'none';
    }
}
