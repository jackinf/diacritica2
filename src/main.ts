import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import * as path from 'path';
import * as fs from 'fs';
import * as XLSX from 'xlsx';

let mainWindow: BrowserWindow | null = null;

// Diacritics mapping
const diacriticsMap: { [key: string]: string } = {
    'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A', 'Ä': 'A', 'Å': 'A', 'Æ': 'AE',
    'Ç': 'C', 'È': 'E', 'É': 'E', 'Ê': 'E', 'Ë': 'E',
    'Ì': 'I', 'Í': 'I', 'Î': 'I', 'Ï': 'I', 'Ð': 'D', 'Ñ': 'N',
    'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O', 'Ö': 'O', 'Ø': 'O',
    'Ù': 'U', 'Ú': 'U', 'Û': 'U', 'Ü': 'U', 'Ý': 'Y', 'Þ': 'TH',
    'ß': 'ss', 'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a',
    'å': 'a', 'æ': 'ae', 'ç': 'c', 'è': 'e', 'é': 'e', 'ê': 'e',
    'ë': 'e', 'ì': 'i', 'í': 'i', 'î': 'i', 'ï': 'i', 'ð': 'd',
    'ñ': 'n', 'ò': 'o', 'ó': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
    'ø': 'o', 'ù': 'u', 'ú': 'u', 'û': 'u', 'ü': 'u', 'ý': 'y',
    'þ': 'th', 'ÿ': 'y'
};

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 500,
        height: 400,
        resizable: false,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        },
        icon: path.join(__dirname, 'assets/icon.png')
    });

    mainWindow.loadFile(path.join(__dirname, 'index.html'));

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (mainWindow === null) {
        createWindow();
    }
});

// Remove diacritics from a string
function removeDiacritics(text: string): string {
    if (typeof text !== 'string') {
        return text;
    }

    let result = '';
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        result += diacriticsMap[char] || char;
    }
    return result;
}

// Process Excel file
ipcMain.handle('process-file', async (event, filePath: string) => {
    try {
        // Read the Excel file
        const workbook = XLSX.readFile(filePath);

        // Process each worksheet
        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });

            // Process each cell
            const processedData = jsonData.map((row: any) => {
                if (Array.isArray(row)) {
                    return row.map((cell: any) => {
                        if (typeof cell === 'string') {
                            return removeDiacritics(cell);
                        }
                        return cell;
                    });
                }
                return row;
            });

            // Convert back to worksheet
            const newWorksheet = XLSX.utils.aoa_to_sheet(processedData);

            // Preserve cell properties (formulas, etc.)
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                    const originalCell = worksheet[cellAddress];
                    const newCell = newWorksheet[cellAddress];

                    if (originalCell && newCell) {
                        // Preserve formula if it exists
                        if (originalCell.f) {
                            newCell.f = originalCell.f;
                        }
                        // Only update value if it's not a formula
                        if (!originalCell.f && originalCell.v !== undefined) {
                            newCell.v = removeDiacritics(String(originalCell.v));
                            newCell.w = removeDiacritics(String(originalCell.w || originalCell.v));
                        }
                    }
                }
            }

            workbook.Sheets[sheetName] = newWorksheet;
        }

        // Create output filename
        const dir = path.dirname(filePath);
        const ext = path.extname(filePath);
        const baseName = path.basename(filePath, ext);
        const outputPath = path.join(dir, `${baseName}_fixed${ext}`);

        // Write the processed file
        XLSX.writeFile(workbook, outputPath);

        return {
            success: true,
            message: `File processed successfully!\nSaved as: ${path.basename(outputPath)}`
        };
    } catch (error) {
        return {
            success: false,
            message: `Error processing file: ${error}`
        };
    }
});

// Handle file dialog
ipcMain.handle('select-file', async () => {
    const result = await dialog.showOpenDialog(mainWindow!, {
        properties: ['openFile'],
        filters: [
            { name: 'Excel Files', extensions: ['xlsx', 'xlsm', 'xls'] },
            { name: 'All Files', extensions: ['*'] }
        ]
    });

    if (!result.canceled && result.filePaths.length > 0) {
        return result.filePaths[0];
    }
    return null;
});