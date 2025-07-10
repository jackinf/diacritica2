import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import * as path from 'path';
import * as fs from 'fs';
import * as XLSX from 'xlsx';

let mainWindow: BrowserWindow | null = null;

// Default diacritics mapping
const defaultDiacriticsMap: { [key: string]: string } = {
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
    'þ': 'th', 'ÿ': 'y', 'Š': 'S', 'š': 's', 'Ž': 'Z', 'ž': 'z',
    'ľ': 'l', 'ć': 'c', 'č': 'c', 'ř': 'r', 'ň': 'n', 'ť': 't', 'ď': 'd',
    // Special characters and symbols
    '°': 'o',     // Degree symbol
    'º': '',      // Masculine ordinal indicator
    'ª': '',      // Feminine ordinal indicator
    '`': '',      // Grave accent
    '´': '',      // Acute accent
    // Quotation marks
    '"': '',      // Left double quotation mark
    '""': '',
    // '"': '',      // Right double quotation mark
    // ''': '',      // Left single quotation mark
    // ''': '',      // Right single quotation mark (apostrophe)
    '„': '',      // Double low-9 quotation mark
    '‚': '',      // Single low-9 quotation mark
    // Common apostrophes and quotes
    "'": ' ',     // Straight apostrophe
    // '"': '',      // Straight double quote
    // Other common special characters
    '–': '-',     // En dash
    '—': '-',     // Em dash
    '…': '...',   // Ellipsis
    '•': '*',     // Bullet
    '·': '.',     // Middle dot
    '¸': '',      // Cedilla
    '¨': '',      // Diaeresis
    '˚': 'o',     // Ring above
    '˙': '',      // Dot above
    'ˇ': '',      // Caron
    '˘': '',      // Breve
    '¯': '',      // Macron
    '˛': '',      // Ogonek
    '˝': '',      // Double acute
};

let diacriticsMap: { [key: string]: string } = { ...defaultDiacriticsMap };

// Get the user data path for storing config
function getConfigPath(): string {
    return path.join(app.getPath('userData'), 'character-mappings.json');
}

// Load character mappings from config file
async function loadCharacterMappings(): Promise<void> {
    const configPath = getConfigPath();
    
    try {
        if (fs.existsSync(configPath)) {
            const data = fs.readFileSync(configPath, 'utf8');
            const customMappings = JSON.parse(data);
            
            // Merge with default mappings (custom mappings take precedence)
            diacriticsMap = { ...defaultDiacriticsMap, ...customMappings };
            
            console.log('Loaded custom character mappings');
        } else {
            // Create default config file if it doesn't exist
            await saveCharacterMappings(defaultDiacriticsMap);
            console.log('Created default character mappings file');
        }
    } catch (error) {
        console.error('Error loading character mappings:', error);
        // Fall back to default mappings
        diacriticsMap = { ...defaultDiacriticsMap };
    }
}

// Save character mappings to config file
async function saveCharacterMappings(mappings: { [key: string]: string }): Promise<void> {
    const configPath = getConfigPath();
    
    try {
        // Ensure the directory exists
        const dir = path.dirname(configPath);
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
        }
        
        // Write the mappings with pretty formatting
        fs.writeFileSync(configPath, JSON.stringify(mappings, null, 2));
    } catch (error) {
        console.error('Error saving character mappings:', error);
    }
}

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 500,
        height: 450,
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

app.whenReady().then(async () => {
    await loadCharacterMappings();
    createWindow();
});

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
    let unmappedChars = new Set<string>();
    
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        if (diacriticsMap[char] !== undefined) {
            result += diacriticsMap[char];
        } else {
            result += char;
            // Log unmapped special characters (excluding common ASCII)
            if (char.charCodeAt(0) > 127) {
                unmappedChars.add(char);
            }
        }
    }
    
    // Log unmapped characters for debugging
    if (unmappedChars.size > 0) {
        console.log('Unmapped special characters found:', Array.from(unmappedChars).map(c => ({
            char: c,
            code: c.charCodeAt(0),
            hex: '0x' + c.charCodeAt(0).toString(16).toUpperCase()
        })));
    }
    
    return result;
}

// Process Excel file
ipcMain.handle('process-file', async (event, filePath: string) => {
    try {
        // Reload mappings in case they were changed
        await loadCharacterMappings();
        
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

// Open config file in default editor
ipcMain.handle('open-config', async () => {
    const configPath = getConfigPath();
    
    try {
        // Ensure the config file exists
        if (!fs.existsSync(configPath)) {
            await saveCharacterMappings(defaultDiacriticsMap);
        }
        
        // Open the file with the default system editor
        const { shell } = require('electron');
        shell.openPath(configPath);
        
        return { success: true };
    } catch (error) {
        return { 
            success: false, 
            message: `Error opening config file: ${error}` 
        };
    }
});

// Analyze file for special characters
ipcMain.handle('analyze-file', async (event, filePath: string) => {
    try {
        const workbook = XLSX.readFile(filePath);
        const specialChars = new Map<string, number>();
        
        // Process each worksheet
        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
            
            // Analyze each cell
            jsonData.forEach((row: any) => {
                if (Array.isArray(row)) {
                    row.forEach((cell: any) => {
                        if (typeof cell === 'string') {
                            for (let i = 0; i < cell.length; i++) {
                                const char = cell[i];
                                if (char.charCodeAt(0) > 127) { // Non-ASCII character
                                    specialChars.set(char, (specialChars.get(char) || 0) + 1);
                                }
                            }
                        }
                    });
                }
            });
        }
        
        // Convert to array and sort by frequency
        const charArray = Array.from(specialChars.entries())
            .map(([char, count]) => ({
                char,
                count,
                code: char.charCodeAt(0),
                hex: '0x' + char.charCodeAt(0).toString(16).toUpperCase(),
                mapped: diacriticsMap[char] !== undefined ? diacriticsMap[char] : 'NOT MAPPED'
            }))
            .sort((a, b) => b.count - a.count);
        
        return {
            success: true,
            characters: charArray
        };
    } catch (error) {
        return {
            success: false,
            message: `Error analyzing file: ${error}`
        };
    }
});

// Update a single mapping
ipcMain.handle('update-mapping', async (event, oldChar: string, newChar: string) => {
    try {
        if (oldChar) {
            diacriticsMap[oldChar] = newChar;
            await saveCharacterMappings(diacriticsMap);
            return { success: true };
        }
        return { success: false, message: 'Invalid character' };
    } catch (error) {
        return { success: false, message: `Error updating mapping: ${error}` };
    }
});

// Delete a mapping
ipcMain.handle('delete-mapping', async (event, char: string) => {
    try {
        delete diacriticsMap[char];
        await saveCharacterMappings(diacriticsMap);
        return { success: true };
    } catch (error) {
        return { success: false, message: `Error deleting mapping: ${error}` };
    }
});