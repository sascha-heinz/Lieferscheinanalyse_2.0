let selectedFiles = [];
let progressBar;
let tableContainer;
let infoBox;
let errorFiles = [];

const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const analyzeBtn = document.getElementById('analyzeBtn');
const exportBtn = document.getElementById('exportBtn');
const resetBtn = document.getElementById('resetBtn');

window.onload = () => {
    progressBar = document.getElementById('progressBar');
    tableContainer = document.getElementById('table-container');
    infoBox = document.getElementById('info-box');

    resetBtn.disabled = true;

    uploadBtn.addEventListener('click', () => fileInput.click());
    analyzeBtn.addEventListener('click', startAnalysis);
    exportBtn.addEventListener('click', exportToExcel);
    resetBtn.addEventListener('click', resetAll);

    fileInput.addEventListener('change', (event) => {
        selectedFiles = Array.from(event.target.files);
        const messageElement = document.getElementById('file-message');
        
        if (selectedFiles.length > 0) {
            messageElement.style.display = 'inline'; // Datei-Info sichtbar machen
            messageElement.textContent = `${selectedFiles.length} Bilder hinzugefügt.`;
    
            // Deaktiviere den "jpg's einfügen"-Button direkt nach der Dateiauswahl
            uploadBtn.disabled = true;

            // Aktiviere den Reset-Button, wenn Dateien hochgeladen wurden
            resetBtn.disabled = false;
        } else {
            messageElement.style.display = 'none'; // Datei-Info verstecken
        }
        
        analyzeBtn.disabled = selectedFiles.length === 0; // Analyse-Button aktivieren/deaktivieren
    });
    
};

// Analyse starten
async function startAnalysis() {
    if (selectedFiles.length === 0) {
        alert('Bitte wählen Sie zuerst Dateien aus.');
        return;
    }

    // Deaktiviere den "jpg's einfügen"-Button während der Analyse
    uploadBtn.disabled = true;
    analyzeBtn.disabled = true;
    resetBtn.disabled = true;
    



    progressBar.value = 0;
    progressBar.max = selectedFiles.length;
    errorFiles = [];

    clearTable();

    for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        try {
            const qrData = await processFileWithDynamicScaling(file);
            if (qrData) {
                addToTable(i + 1, file.name, qrData); // Hier fügen wir die laufende Nummer hinzu
            } else {
                errorFiles.push(`${file.name}: QR-Code nicht erkannt`);
            }
        } catch (e) {
            console.error(`Fehler beim Verarbeiten von ${file.name}:`, e);
            errorFiles.push(`${file.name}: ${e.message || e}`); // Fehlerdetails werden hinzugefügt
        }
        progressBar.value = i + 1;
    }

    showInfoBox();
    exportBtn.disabled = false;
    resetBtn.disabled = false;
    analyzeBtn.disabled = true;
    // Der Button bleibt deaktiviert, bis auf "Zurücksetzen" geklickt wird
}

// Bild skalieren
function scaleImage(image, scaleFactor) {
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    canvas.width = image.width * scaleFactor;
    canvas.height = image.height * scaleFactor;
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return canvas;
}

// QR-Code lesen
function tryToReadQRCode(image) {
    return new Promise((resolve, reject) => {
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.width = image.width;
        canvas.height = image.height;
        context.drawImage(image, 0, 0);
        
        const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
        const code = jsQR(imageData.data, canvas.width, canvas.height);
        if (code) {
            resolve(parseQRData(code.data));
        } else {
            resolve(null); // Kein QR-Code erkannt
        }
    });
}

// QR-Daten analysieren
function parseQRData(data) {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(data, 'text/xml');
        
        const parserError = xmlDoc.getElementsByTagName('parsererror');
        if (parserError.length > 0) {
            throw new Error('Fehler beim Parsen des QR-Codes: Ungültiges XML');
        }

        const tags = Array.from(xmlDoc.getElementsByTagName('*')).filter(el => el.children.length === 0);
        const result = {};
        tags.forEach(tag => {
            result[tag.tagName] = tag.textContent;
        });
        return result;

    } catch (error) {
        throw new Error(`Fehler bei der Analyse des QR-Codes: ${error.message}`);
    }
}

// Bild laden
function loadImage(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.onload = () => resolve(img);
            img.onerror = (err) => reject(err);
            img.src = e.target.result;
        };
        reader.readAsDataURL(file);
    });
}

// Dynamische Skalierung (Vergrößern und Verkleinern)
async function processFileWithDynamicScaling(file) {
    const img = await loadImage(file);
    let qrData = null;

    const originalWidth = img.width;
    const originalHeight = img.height;

    // Versuch mit Originalgröße
    qrData = await tryToReadQRCode(img);
    if (qrData) return qrData;

    // Maximale und minimale Skalierungswerte
    const minScale = 0.10;
    const maxScale = 2.0;
    
    // Dynamische Skalierung: Versuchen, den QR-Code schrittweise größer und kleiner zu machen
    for (let scale = 1.10; scale <= maxScale; scale += 0.10) {
        const scaledImage = scaleImage(img, scale);
        qrData = await tryToReadQRCode(scaledImage);
        if (qrData) return qrData;
    }

    // Wenn nicht erkannt, dann versuche es in umgekehrter Reihenfolge (Verkleinerung)
    for (let scale = 0.90; scale >= minScale; scale -= 0.10) {
        const scaledImage = scaleImage(img, scale);
        qrData = await tryToReadQRCode(scaledImage);
        if (qrData) return qrData;
    }

    return null;
}

// Ergebnisse in der Tabelle hinzufügen (mit laufender Nummer)
function addToTable(index, fileName, qrData) {
    const table = document.getElementById('result-table');
    if (!table) {
        createTable(Object.keys(qrData));
    }

    const row = document.createElement('tr');

    // Laufende Nummer
    const indexCell = document.createElement('td');
    indexCell.textContent = index;
    row.appendChild(indexCell);

    // Dateiname
    const fileNameCell = document.createElement('td');
    fileNameCell.textContent = fileName;
    row.appendChild(fileNameCell);

    Object.values(qrData).forEach(value => {
        const cell = document.createElement('td');
        cell.textContent = value;
        row.appendChild(cell);
    });

    document.getElementById('result-table').appendChild(row);
}

// Tabelle erstellen
function createTable(headers) {
    const table = document.createElement('table');
    table.id = 'result-table';

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    // Laufende Nummer Spalte
    const indexHeader = document.createElement('th');
    indexHeader.textContent = 'Nr.';
    headerRow.appendChild(indexHeader);

    // Dateiname
    const fileHeader = document.createElement('th');
    fileHeader.textContent = 'Dateiname';
    headerRow.appendChild(fileHeader);

    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);
    tableContainer.appendChild(table);
}

// Tabelle leeren
function clearTable() {
    tableContainer.innerHTML = '';
}

// Infokasten anzeigen
function showInfoBox() {
    if (errorFiles.length > 0) {
        infoBox.style.display = 'block';
        infoBox.innerHTML = `<p>Analyse abgeschlossen. Fehler bei folgenden Dateien:</p><ul>${errorFiles.map(file => `<li>${file}</li>`).join('')}</ul>`;
    } else {
        infoBox.style.display = 'none';
    }
}

// Alles zurücksetzen
function resetAll() {
    location.reload(); // Die Seite wird neu geladen
}


// Exportieren der Tabelle nach Excel
function exportToExcel() {
    const table = document.getElementById('result-table');
    if (!table) {
        alert('Keine Daten zum Exportieren.');
        return;
    }

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.table_to_sheet(table);

    // Hier fügen wir die Formatierung als Excel-Tabelle hinzu
    const range = { s: { r: 0, c: 0 }, e: { r: table.rows.length - 1, c: table.rows[0].cells.length - 1 } };
    worksheet['!ref'] = XLSX.utils.encode_range(range);
    worksheet['!autofilter'] = { ref: worksheet['!ref'] }; // Autofilter für die Tabelle aktivieren

    // Jetzt wird die Tabelle als Format 'Table' hinzugefügt
    const tableProperties = {
        range: worksheet['!ref'], // Bereich der Tabelle
        name: 'Ergebnisse',
        headerRow: true, // Erste Zeile als Kopfzeile markieren
        totalsRow: false, // Keine Summenzeile
        columns: Array(table.rows[0].cells.length).fill({}) // Jede Spalte standardmäßig ohne Formatierung
    };
    worksheet['!table'] = tableProperties;

    // Hinzufügen der Arbeitsblattdaten
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Ergebnisse');

    // Schreiben der Excel-Datei
    XLSX.writeFile(workbook, 'Ergebnisse.xlsx');
}
