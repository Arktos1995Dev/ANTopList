document.addEventListener('DOMContentLoaded', async () => {
    await fetchVersion();
    checkAuthState();
});

// ... (rest of the existing code) ...

// --- Authentication --- //

const loginBtn = document.getElementById('loginBtn');
const registerBtn = document.getElementById('registerBtn');
const logoutBtn = document.getElementById('logoutBtn');
const saveListBtn = document.getElementById('saveListBtn');
const authModal = document.getElementById('authModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const showRegisterLink = document.getElementById('showRegisterLink');
const showLoginLink = document.getElementById('showLoginLink');
const loginForm = document.getElementById('loginForm');
const registerForm = document.getElementById('registerForm');
const loginSubmit = document.getElementById('loginSubmit');
const registerSubmit = document.getElementById('registerSubmit');
const userInfo = document.getElementById('userInfo');
const profileLink = document.getElementById('profileLink');

loginBtn.addEventListener('click', () => {
    authModal.style.display = 'block';
    loginForm.style.display = 'block';
    registerForm.style.display = 'none';
    showLoginLink.style.display = 'none';
    showRegisterLink.style.display = 'block';
});

registerBtn.addEventListener('click', () => {
    authModal.style.display = 'block';
    loginForm.style.display = 'none';
    registerForm.style.display = 'block';
    showLoginLink.style.display = 'block';
    showRegisterLink.style.display = 'none';
});

closeModalBtn.addEventListener('click', () => {
    authModal.style.display = 'none';
});

window.addEventListener('click', (event) => {
    if (event.target == authModal) {
        authModal.style.display = 'none';
    }
});

showRegisterLink.addEventListener('click', (e) => {
    e.preventDefault();
    loginForm.style.display = 'none';
    registerForm.style.display = 'block';
    showLoginLink.style.display = 'block';
    showRegisterLink.style.display = 'none';
});

showLoginLink.addEventListener('click', (e) => {
    e.preventDefault();
    registerForm.style.display = 'none';
    loginForm.style.display = 'block';
    showRegisterLink.style.display = 'block';
    showLoginLink.style.display = 'none';
});

registerSubmit.addEventListener('click', async () => {
    const username = document.getElementById('registerUsername').value;
    const password = document.getElementById('registerPassword').value;

    try {
        const response = await fetch('/api/register', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username, password })
        });
        const data = await response.json();
        alert(data.message);
        if (response.ok) {
            authModal.style.display = 'none';
        }
    } catch (error) {
        alert('Registration failed.');
    }
});

loginSubmit.addEventListener('click', async () => {
    const username = document.getElementById('loginUsername').value;
    const password = document.getElementById('loginPassword').value;

    try {
        const response = await fetch('/api/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username, password })
        });

        const data = await response.json();
        if (response.ok) {
            localStorage.setItem('token', data.token);
            localStorage.setItem('username', username);
            updateUIForAuth(username);
            authModal.style.display = 'none';
            loadUserList();
        } else {
            alert(data.message);
        }
    } catch (error) {
        alert('Login failed.');
    }
});

logoutBtn.addEventListener('click', () => {
    localStorage.removeItem('token');
    localStorage.removeItem('username');
    updateUIForGuest();
});

profileLink.addEventListener('click', (e) => {
    e.preventDefault();
    window.location.href = 'profile.html';
});

saveListBtn.addEventListener('click', async () => {
    if (!workbook || !currentSheetName) {
        alert('No list to save.');
        return;
    }

    const token = localStorage.getItem('token');
    const worksheet = workbook.Sheets[currentSheetName];
    const listData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    try {
        const response = await fetch('/api/list', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': token
            },
            body: JSON.stringify({ list: listData })
        });
        const data = await response.json();
        alert(data.message);
    } catch (error) {
        alert('Failed to save list.');
    }
});

function checkAuthState() {
    const token = localStorage.getItem('token');
    const username = localStorage.getItem('username');
    if (token && username) {
        updateUIForAuth(username);
        loadUserList();
    } else {
        updateUIForGuest();
    }
}

function updateUIForAuth(username) {
    loginBtn.style.display = 'none';
    registerBtn.style.display = 'none';
    userInfo.style.display = 'flex';
    profileLink.textContent = username;
    saveListBtn.style.display = 'block';
}

function updateUIForGuest() {
    loginBtn.style.display = 'block';
    registerBtn.style.display = 'block';
    userInfo.style.display = 'none';
    profileLink.textContent = '';
    saveListBtn.style.display = 'none';
}

async function loadUserList() {
    const token = localStorage.getItem('token');
    if (!token) return;

    try {
        const response = await fetch('/api/list', {
            headers: { 'Authorization': token }
        });
        if (response.ok) {
            const listData = await response.json();
            if (listData && listData.length > 0) {
                const worksheet = XLSX.utils.aoa_to_sheet(listData);
                workbook = XLSX.utils.book_new();
                const sheetName = 'My Saved List';
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

                currentSheetName = sheetName;
                sheetSelect.innerHTML = `<option value="${currentSheetName}">${currentSheetName}</option>`;

                controlsSection.style.display = 'flex';
                dataSection.style.display = 'block';
                infoSection.style.display = 'block';
                uploadBox.style.display = 'none';

                updateInfo();
                renderData();
                initiateImageDownloads();
            }
        }
    } catch (error) {
        console.error('Failed to load user list:', error);
    }
}
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
let localImageMapping = {}; 
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


uploadBox.addEventListener('click', () => fileInput.click());
uploadBox.addEventListener('dragover', (e) => { e.preventDefault(); uploadBox.classList.add('dragover'); });
uploadBox.addEventListener('dragleave', (e) => { e.preventDefault(); uploadBox.classList.remove('dragover'); });
uploadBox.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);

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

function processFile(file) {
    if (file.name.endsWith('.xml')) {
        processXmlFile(file);
    } else {
        processExcelFile(file);
    }
}

function processExcelFile(file) {
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
            uploadBox.style.display = 'none';
            
            updateInfo();
            renderData();
            initiateImageDownloads();

        } catch (error) {
            alert('Ошибка при чтении файла: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

function processXmlFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const xmlString = e.target.result;
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, "application/xml");
        const sheetName = file.name.replace(/\.xml$/, '');

        const animeList = Array.from(xmlDoc.getElementsByTagName('anime')).map(anime => {
            const series_title = anime.getElementsByTagName('series_title')[0]?.textContent || '';
            const my_score = anime.getElementsByTagName('my_score')[0]?.textContent || '0';
            const my_status = anime.getElementsByTagName('my_status')[0]?.textContent || anime.getElementsByTagName('shiki_status')[0]?.textContent ||'';
            const series_type = anime.getElementsByTagName('series_type')[0]?.textContent || '';
            const series_episodes = anime.getElementsByTagName('series_episodes')[0]?.textContent || '';

            return {
                'Anime': series_title,
                'Score': my_score,
                'Status': my_status,
                'Type': series_type,
                'Episodes': series_episodes,
            };
        });
        
        const headers = ['Anime', 'Score', 'Status', 'Type', 'Episodes'];
        const data = animeList.map(item => headers.map(header => item[header]));

        const worksheet = XLSX.utils.aoa_to_sheet([headers, ...data]);
        workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

        currentSheetName = sheetName;
        sheetSelect.innerHTML = `<option value="${currentSheetName}">${currentSheetName}</option>`;

        controlsSection.style.display = 'flex';
        dataSection.style.display = 'block';
        infoSection.style.display = 'block';
        uploadBox.style.display = 'none';

        updateInfo();
        renderData();
        initiateImageDownloads();
    };
    reader.readAsText(file);
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

        await renderData();

    } catch (error) {
        console.error('Ошибка при запросе на скачивание изображений:', error);
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
    initSortable();
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
    return Object.entries(style).map(([key, value]) => `${key.replace(/([A-Z])/g, '-$1').toLowerCase()}: ${value}`).join('; ');
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

    let html = '<thead><tr><th></th>';
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
        html += `<tr${rowStyle} data-row-index="${rowIndex}">`;
        html += '<td class="drag-handle">&#9776;</td>';

        if (animeColumnIndex >= 0) {
            const animeTitle = row[animeColumnIndex] || '';
            const mal_id = localImageMapping[animeTitle.trim()];
            html += `<td class="anime-image-cell image-col">`;
            if (mal_id) {
                html += `<img src="/api/image/${escapeHtml(mal_id)}" alt="${escapeHtml(animeTitle)}" class="anime-image" loading="lazy">`;
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
        html += `<div class="card"${cardStyle} data-row-index="${rowIndex}">`;
        html += '<span class="drag-handle">&#9776;</span>';
        
        const animeTitle = animeColumnIndex >= 0 ? (row[animeColumnIndex] || '') : '';
        const mal_id = localImageMapping[animeTitle.trim()];
        if (mal_id) {
            html += `<div class="card-image-wrapper"><img src="/api/image/${escapeHtml(mal_id)}" alt="${escapeHtml(animeTitle)}" class="card-anime-image" loading="lazy"></div>`;
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

let sortableInstance = null;

function initSortable() {
    if (sortableInstance) {
        sortableInstance.destroy();
    }

    const target = currentView === 'table' ? dataTable.querySelector('tbody') : cardsContainer;

    sortableInstance = new Sortable(target, {
        animation: 150,
        handle: '.drag-handle',
        ghostClass: 'sortable-ghost',
        chosenClass: 'sortable-chosen',
        onEnd: (evt) => {
            const { oldIndex, newIndex } = evt;
            if (oldIndex === newIndex) return;

            const worksheet = workbook.Sheets[currentSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            let headerRowIndex = 0;
            for(let i=0; i < jsonData.length; i++) {
                if (jsonData[i].some(cell => cell !== '')) { headerRowIndex = i; break; }
            }
            const headers = jsonData[headerRowIndex];
            const dataRows = jsonData.slice(headerRowIndex + 1);

            const [movedRow] = dataRows.splice(oldIndex, 1);
            dataRows.splice(newIndex, 0, movedRow);

            const newSheetData = [headers, ...dataRows];
            const newWorksheet = XLSX.utils.aoa_to_sheet(newSheetData);
            
            // Preserve hyperlinks and styles
            Object.keys(worksheet).forEach(cellAddress => {
                if (worksheet[cellAddress].l) {
                    const newCell = newWorksheet[cellAddress];
                    if(newCell) newCell.l = worksheet[cellAddress].l;
                }
                 if (worksheet[cellAddress].s) {
                    const newCell = newWorksheet[cellAddress];
                    if(newCell) newCell.s = worksheet[cellAddress].s;
                }
            });
            if (worksheet['!hyperlinks']) {
                newWorksheet['!hyperlinks'] = worksheet['!hyperlinks'];
            }

            workbook.Sheets[currentSheetName] = newWorksheet;

            renderData();
        }
    });
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
