// Type definitions for the electron API
interface ElectronAPI {
    processFile: (filePath: string) => Promise<{ success: boolean; message: string }>;
    selectFile: () => Promise<string | null>;
    openConfig: () => Promise<{ success: boolean; message?: string }>;
    getMappings: () => Promise<{ [key: string]: string }>;
    updateMapping: (oldChar: string, newChar: string) => Promise<{ success: boolean; message?: string }>;
    deleteMapping: (char: string) => Promise<{ success: boolean; message?: string }>;
}

// @ts-ignore
declare global {
    interface Window {
        electronAPI: ElectronAPI;
    }
}

class DiacriticsRemoverApp {
    private dropZone: HTMLElement;
    private fileLabel: HTMLElement;
    private selectedFileLabel: HTMLElement;
    private fixButton: HTMLButtonElement;
    private browseButton: HTMLButtonElement;
    private configButton: HTMLButtonElement;
    private settingsButton: HTMLElement;
    private currentFile: string | null = null;

    constructor() {
        this.dropZone = document.getElementById('dropZone')!;
        this.fileLabel = document.getElementById('fileLabel')!;
        this.selectedFileLabel = document.getElementById('selectedFileLabel')!;
        this.fixButton = document.getElementById('fixButton') as HTMLButtonElement;
        this.browseButton = document.getElementById('browseButton') as HTMLButtonElement;
        this.configButton = document.getElementById('configButton') as HTMLButtonElement;
        this.settingsButton = document.getElementById('settingsButton')!;

        this.setupEventListeners();
    }

    private setupEventListeners(): void {
        // Drag and drop events
        this.dropZone.addEventListener('drop', this.handleDrop.bind(this));
        this.dropZone.addEventListener('dragover', this.handleDragOver.bind(this));
        this.dropZone.addEventListener('dragleave', this.handleDragLeave.bind(this));
        this.dropZone.addEventListener('dragenter', this.handleDragEnter.bind(this));

        // Button events
        this.browseButton.addEventListener('click', this.handleBrowse.bind(this));
        this.fixButton.addEventListener('click', this.handleFix.bind(this));
        this.configButton.addEventListener('click', this.handleOpenConfig.bind(this));
        this.settingsButton.addEventListener('click', this.handleOpenSettings.bind(this));
    }

    private handleDragOver(e: DragEvent): void {
        e.preventDefault();
        e.stopPropagation();
    }

    private handleDragEnter(e: DragEvent): void {
        e.preventDefault();
        this.dropZone.classList.add('drag-over');
    }

    private handleDragLeave(e: DragEvent): void {
        e.preventDefault();
        this.dropZone.classList.remove('drag-over');
    }

    private handleDrop(e: DragEvent): void {
        e.preventDefault();
        e.stopPropagation();
        this.dropZone.classList.remove('drag-over');

        const files = e.dataTransfer?.files;
        if (files && files.length > 0) {
            const file = files[0];
            if (this.isValidExcelFile(file.name)) {
                this.setSelectedFile(file.path);
            } else {
                this.showError('Please drop a valid Excel file (.xlsx, .xlsm, .xls)');
            }
        }
    }

    private async handleBrowse(): Promise<void> {
        // @ts-ignore
        const filePath = await window.electronAPI.selectFile();
        if (filePath) {
            this.setSelectedFile(filePath);
        }
    }

    private async handleFix(): Promise<void> {
        if (!this.currentFile) return;

        this.fixButton.disabled = true;
        this.fixButton.textContent = 'Processing...';

        try {
            // @ts-ignore
            const result = await window.electronAPI.processFile(this.currentFile);

            if (result.success) {
                this.showSuccess(result.message);
                this.resetInterface();
            } else {
                this.showError(result.message);
            }
        } catch (error) {
            this.showError(`Error: ${error}`);
        } finally {
            this.fixButton.disabled = false;
            this.fixButton.textContent = 'Fix';
        }
    }

    private async handleOpenConfig(): Promise<void> {
        try {
            // @ts-ignore
            const result = await window.electronAPI.openConfig();
            if (!result.success) {
                this.showError(result.message || 'Failed to open config file');
            }
        } catch (error) {
            this.showError(`Error opening config: ${error}`);
        }
    }

    private async handleOpenSettings(): Promise<void> {
        const modal = document.getElementById('settingsModal')!;
        modal.style.display = 'flex';
        await this.loadMappings();
    }

    private async loadMappings(): Promise<void> {
        try {
            // @ts-ignore
            const mappings = await window.electronAPI.getMappings();
            const container = document.getElementById('mappingsContainer')!;
            container.innerHTML = '';

            Object.entries(mappings).forEach(([char, replacement]) => {
                const item = this.createMappingItem(char, replacement as string);
                container.appendChild(item);
            });
        } catch (error) {
            console.error('Error loading mappings:', error);
        }
    }

    private createMappingItem(char: string, replacement: string): HTMLElement {
        const item = document.createElement('div');
        item.className = 'mapping-item';
        item.innerHTML = `
            <span class="char-display">${this.escapeHtml(char)}</span>
            <span class="arrow">→</span>
            <input type="text" class="replacement-input" value="${this.escapeHtml(replacement)}" data-char="${this.escapeHtml(char)}">
            <button class="delete-button" data-char="${this.escapeHtml(char)}">Delete</button>
        `;

        const input = item.querySelector('.replacement-input') as HTMLInputElement;
        const deleteBtn = item.querySelector('.delete-button') as HTMLButtonElement;

        input.addEventListener('change', async (e) => {
            const target = e.target as HTMLInputElement;
            const originalChar = target.getAttribute('data-char')!;
            const newValue = target.value;

            try {
                // @ts-ignore
                const result = await window.electronAPI.updateMapping(originalChar, newValue);
                if (!result.success) {
                    this.showError(result.message || 'Failed to update mapping');
                    target.value = replacement; // Revert on error
                }
            } catch (error) {
                this.showError(`Error updating mapping: ${error}`);
                target.value = replacement; // Revert on error
            }
        });

        deleteBtn.addEventListener('click', async () => {
            const charToDelete = deleteBtn.getAttribute('data-char')!;
            if (confirm(`Delete mapping for "${charToDelete}"?`)) {
                try {
                    // @ts-ignore
                    const result = await window.electronAPI.deleteMapping(charToDelete);
                    if (result.success) {
                        item.remove();
                    } else {
                        this.showError(result.message || 'Failed to delete mapping');
                    }
                } catch (error) {
                    this.showError(`Error deleting mapping: ${error}`);
                }
            }
        });

        return item;
    }

    private setSelectedFile(filePath: string): void {
        this.currentFile = filePath;
        const fileName = filePath.split(/[\\/]/).pop() || '';
        this.selectedFileLabel.textContent = `Selected: ${fileName}`;
        this.fileLabel.innerHTML = `✓ ${fileName}`;
        this.fixButton.disabled = false;
    }

    private resetInterface(): void {
        this.currentFile = null;
        this.selectedFileLabel.textContent = 'No file selected';
        this.fileLabel.innerHTML = 'Drag and drop Excel file here<br>(.xlsx, .xlsm, .xls)';
        this.fixButton.disabled = true;
    }

    private isValidExcelFile(filename: string): boolean {
        const validExtensions = ['.xlsx', '.xlsm', '.xls'];
        const ext = filename.toLowerCase().slice(filename.lastIndexOf('.'));
        return validExtensions.includes(ext);
    }

    private escapeHtml(text: string): string {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    private showSuccess(message: string): void {
        const modal = document.getElementById('modal')!;
        const modalTitle = document.getElementById('modalTitle')!;
        const modalMessage = document.getElementById('modalMessage')!;

        modalTitle.textContent = 'Done';
        modalTitle.className = 'modal-title success';
        modalMessage.textContent = message;
        modal.style.display = 'flex';
    }

    private showError(message: string): void {
        const modal = document.getElementById('modal')!;
        const modalTitle = document.getElementById('modalTitle')!;
        const modalMessage = document.getElementById('modalMessage')!;

        modalTitle.textContent = 'Error';
        modalTitle.className = 'modal-title error';
        modalMessage.textContent = message;
        modal.style.display = 'flex';
    }
}

// Initialize the app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new DiacriticsRemoverApp();

    // Modal close button
    document.getElementById('modalClose')?.addEventListener('click', () => {
        document.getElementById('modal')!.style.display = 'none';
    });

    // Settings modal close button
    document.getElementById('closeSettings')?.addEventListener('click', () => {
        document.getElementById('settingsModal')!.style.display = 'none';
    });

    // Add new mapping functionality
    const addButton = document.getElementById('addButton');
    const newCharInput = document.getElementById('newChar') as HTMLInputElement;
    const newReplacementInput = document.getElementById('newReplacement') as HTMLInputElement;

    addButton?.addEventListener('click', async () => {
        const char = newCharInput.value.trim();
        const replacement = newReplacementInput.value.trim();

        if (!char) {
            alert('Please enter a character to map');
            return;
        }

        try {
            // @ts-ignore
            const result = await window.electronAPI.updateMapping(char, replacement);
            if (result.success) {
                // Reload mappings
                const app = new DiacriticsRemoverApp();
                await app['loadMappings']();
                
                // Clear inputs
                newCharInput.value = '';
                newReplacementInput.value = '';
            } else {
                alert(result.message || 'Failed to add mapping');
            }
        } catch (error) {
            alert(`Error adding mapping: ${error}`);
        }
    });
});