// Type definitions for the electron API

interface ElectronAPI {
    processFile: (filePath: string) => Promise<{ success: boolean; message: string }>;
    selectFile: () => Promise<string | null>;
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
    private currentFile: string | null = null;

    constructor() {
        this.dropZone = document.getElementById('dropZone')!;
        this.fileLabel = document.getElementById('fileLabel')!;
        this.selectedFileLabel = document.getElementById('selectedFileLabel')!;
        this.fixButton = document.getElementById('fixButton') as HTMLButtonElement;
        this.browseButton = document.getElementById('browseButton') as HTMLButtonElement;

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
});