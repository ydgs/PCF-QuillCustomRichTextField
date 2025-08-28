


import { IInputs, IOutputs } from "./generated/ManifestTypes";
import Quill from "quill";

// Register font whitelist before creating the Quill instance
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const Font = Quill.import('formats/font') as any;
Font.whitelist = [
    'sans-serif',
    'serif',
    'monospace',
    'arial',
    'times-new-roman',
    'comic-sans'
];
Quill.register(Font, true);

export class CustomRichTextField implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private errorDiv: HTMLDivElement | undefined;

    // Simple HTML validation using DOMParser
    private isValidHTML(html: string): boolean {
        try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(html, 'text/html');
            // If parsererror is present, HTML is invalid
            if (doc.querySelector('parsererror')) {
                return false;
            }
            return true;
        } catch {
            return false;
        }
    }

    private showError(message: string) {
        if (!this.container) return;
        if (!this.errorDiv) {
            this.errorDiv = document.createElement('div');
            this.errorDiv.style.color = 'red';
            this.errorDiv.style.margin = '8px 0';
            this.container.insertBefore(this.errorDiv, this.container.firstChild);
        }
        this.errorDiv.textContent = message;
    }

    private clearError() {
        if (this.errorDiv && this.container) {
            this.container.removeChild(this.errorDiv);
            this.errorDiv = undefined;
        }
    }
    /**
     * Empty constructor.
     */

    private quill: Quill | undefined;
    private lastToolbarConfig = "";
    private container: HTMLDivElement | undefined;
    private _lastHeadingLevelsRaw = "";
    private htmlText = "";
    private lastDefaultValue = "";
    private notifyOutputChanged: (() => void) | undefined;
    private isProgrammaticUpdate = false;
    private _lastTogglesConfig = "";

    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;
        this.createOrUpdateQuill(context);
    }

    private applyBorderColors(element: HTMLElement, context: ComponentFramework.Context<IInputs>, prefix: string, defaultColor = "#ccc") {
        const generalBorderColor = context.parameters[`${prefix}BorderColor` as keyof IInputs]?.raw;
        const topColor = context.parameters[`${prefix}BorderTopColor` as keyof IInputs]?.raw || generalBorderColor || defaultColor;
        const rightColor = context.parameters[`${prefix}BorderRightColor` as keyof IInputs]?.raw || generalBorderColor || defaultColor;
        const bottomColor = context.parameters[`${prefix}BorderBottomColor` as keyof IInputs]?.raw || generalBorderColor || defaultColor;
        const leftColor = context.parameters[`${prefix}BorderLeftColor` as keyof IInputs]?.raw || generalBorderColor || defaultColor;
        
        element.style.borderTop = `1px solid ${topColor}`;
        element.style.borderRight = `1px solid ${rightColor}`;
        element.style.borderBottom = `1px solid ${bottomColor}`;
        element.style.borderLeft = `1px solid ${leftColor}`;
    }

    private createOrUpdateQuill(context: ComponentFramework.Context<IInputs>) {
        // Remove old editor if exists
        if (this.quill) {
            this.quill = undefined;
            if (this.container) {
                this.container.innerHTML = "";
            }
        }
    if (!this.container) return;
    // Create editor div directly in the main container
    const editorDiv = document.createElement("div");
    editorDiv.id = "quill-editor";
    editorDiv.style.width = "100%";
    editorDiv.style.height = "100%";
    editorDiv.style.boxSizing = "border-box";
    // Apply individual border colors to the Quill editor div
    this.applyBorderColors(editorDiv, context, "editor");
    this.container.appendChild(editorDiv);

        // Determine heading levels from property (default to 1,2,3,4)
        let headingLevels: (number | false)[] = [1, 2, 3, 4, false];
        let headingLevelsRaw: string | undefined = undefined;
        if ('headingLevels' in context.parameters && context.parameters['headingLevels'] && typeof context.parameters['headingLevels'].raw === 'string') {
            headingLevelsRaw = context.parameters['headingLevels'].raw as string;
        }
        if (headingLevelsRaw) {
            const parsed = headingLevelsRaw.split(',').map((x: string) => parseInt(x.trim(), 10)).filter((x: number) => !isNaN(x));
            headingLevels = [...parsed, false];
        }

        // Type-safe access to boolean toggles
        function getToggle(param: keyof IInputs, defaultValue = true) {
            const prop = context.parameters[param];
            if (prop && typeof prop.raw === 'boolean') {
                return prop.raw;
            }
            return defaultValue;
        }

        const showBold = getToggle('showBold');
        const showItalic = getToggle('showItalic');
        const showUnderline = getToggle('showUnderline');
        const showLink = getToggle('showLink');
        const showBlockquote = getToggle('showBlockquote');
        const showCodeBlock = getToggle('showCodeBlock');
        const showImage = getToggle('showImage');
        const showOrderedList = getToggle('showOrderedList');
        const showBulletList = getToggle('showBulletList');
        const showFont = getToggle('showFont');
        const showColor = getToggle('showColor');
        const showAlign = getToggle('showAlign');
        const showClean = getToggle('showClean');

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const toolbar: any[] = [];
        toolbar.push([{ header: headingLevels }]);
        if (showFont) toolbar.push([{ font: [] }]);
        const inline: string[] = [];
        if (showBold) inline.push("bold");
        if (showItalic) inline.push("italic");
        if (showUnderline) inline.push("underline");
        if (inline.length) toolbar.push(inline);
        if (showColor) toolbar.push([{ color: [] }, { background: [] }]);
        const block: string[] = [];
        if (showLink) block.push("link");
        if (showBlockquote) block.push("blockquote");
        if (showCodeBlock) block.push("code-block");
        if (showImage) block.push("image");
        if (block.length) toolbar.push(block);
        if (showAlign) toolbar.push([{ align: [] }]);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const lists: any[] = [];
        if (showOrderedList) lists.push({ list: "ordered" });
        if (showBulletList) lists.push({ list: "bullet" });
        if (lists.length) toolbar.push(lists);
        if (showClean) toolbar.push(["clean"]);

        this.quill = new Quill(editorDiv, {
            theme: "snow",
            modules: {
                toolbar
            }
        });

        // Apply user-selected colors efficiently
        setTimeout(() => {
            this.applyColorStyles(context);
        }, 0);

        // Set initial value from Default property
        const initialValue = context.parameters.Default?.raw || "";
        this.htmlText = initialValue;
        this.lastDefaultValue = initialValue;
        this.isProgrammaticUpdate = true;
        try {
            const delta = this.quill.clipboard.convert({ html: initialValue });
            this.quill.setContents(delta, 'silent');
        } finally {
            this.isProgrammaticUpdate = false;
        }

        // Notify that HtmlText output is available with initial value
        if (initialValue && this.notifyOutputChanged) {
            this.notifyOutputChanged();
        }

        // Update HtmlText live on every text change
        this.quill.on("text-change", () => {
            if (this.isProgrammaticUpdate) return;
            const html = this.quill!.root.innerHTML;
            if (html !== this.htmlText) {
                this.htmlText = html;
                if (this.notifyOutputChanged) {
                    this.notifyOutputChanged();
                }
            }
        });
        // Save toolbar config for change detection
        this.lastToolbarConfig = JSON.stringify(toolbar);
    }

    private setupToolbarColorMonitoring(
        toolbarEl: HTMLElement, 
        backgroundColor?: string,
        textColor?: string, 
        iconColor?: string, 
        dropdownColor?: string, 
        dropdownTextColor?: string
    ) {
        // Clean up any existing monitoring
        this.cleanupColorMonitoring(toolbarEl);

        // Simple, efficient color application function
        const applyColors = () => {
            try {
                if (backgroundColor) {
                    document.documentElement.style.setProperty('--toolbar-background-color', backgroundColor);
                }
                
                if (textColor) {
                    document.documentElement.style.setProperty('--toolbar-text-color', textColor);
                }
                
                if (iconColor) {
                    document.documentElement.style.setProperty('--toolbar-icon-color', iconColor);
                }
                
                if (dropdownColor) {
                    document.documentElement.style.setProperty('--toolbar-dropdown-color', dropdownColor);
                }
                
                if (dropdownTextColor) {
                    document.documentElement.style.setProperty('--toolbar-dropdown-text-color', dropdownTextColor);
                }
            } catch (error) {
                console.warn('Error applying toolbar colors:', error);
            }
        };

        // Apply colors immediately
        applyColors();

        // Set up minimal monitoring - only when necessary
        const observer = new MutationObserver((mutations) => {
            let shouldUpdate = false;
            for (const mutation of mutations) {
                if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
                    shouldUpdate = true;
                    break;
                }
            }
            if (shouldUpdate) {
                // Debounce the color application
                setTimeout(applyColors, 50);
            }
        });

        observer.observe(toolbarEl, {
            childList: true,
            subtree: true
        });

        // Store cleanup function
        (toolbarEl as HTMLElement & { _colorCleanup?: () => void })._colorCleanup = () => {
            observer.disconnect();
        };
    }

    private cleanupColorMonitoring(toolbarEl: HTMLElement) {
        const cleanup = (toolbarEl as HTMLElement & { _colorCleanup?: () => void })._colorCleanup;
        if (cleanup) {
            cleanup();
            delete (toolbarEl as HTMLElement & { _colorCleanup?: () => void })._colorCleanup;
        }
    }

    private isValidCSSColor(color: string): boolean {
        if (!color || color.trim() === '') return false;
        
        // Create a temporary element to test the color
        const tempElement = document.createElement('div');
        tempElement.style.color = color;
        return tempElement.style.color !== '';
    }

    private applyColorStyles(context: ComponentFramework.Context<IInputs>) {
        if (!this.container) return;

        // Get all color values with defaults and validation
        const getValidColor = (paramName: keyof IInputs, defaultColor: string): string => {
            const value = context.parameters[paramName]?.raw as string;
            if (value && this.isValidCSSColor(value)) {
                return value;
            }
            return defaultColor;
        };

        const editorBackgroundColor = getValidColor('editorBackgroundColor', '#ffffff');
        const toolbarColor = getValidColor('toolbarColor', '#ffffff');
        const toolbarTextColor = getValidColor('toolbarTextColor', '#444444');
        const toolbarIconColor = getValidColor('toolbarIconColor', '#444444');
        const toolbarDropdownColor = getValidColor('toolbarDropdownColor', '#ffffff');
        const toolbarDropdownTextColor = getValidColor('toolbarDropdownTextColor', '#333333');

        // Apply editor styles
        const editorDiv = this.container.querySelector('#quill-editor') as HTMLDivElement | null;
        if (editorDiv) {
            this.applyBorderColors(editorDiv, context, "editor");
        }
        
        const qlEditorEl = this.container.querySelector('.ql-editor') as HTMLElement | null;
        if (qlEditorEl) {
            qlEditorEl.style.background = editorBackgroundColor;
        }
        
        // Apply toolbar styles
        const toolbarEl = this.container.querySelector('.ql-toolbar') as HTMLElement | null;
        if (toolbarEl) {
            this.applyBorderColors(toolbarEl, context, "toolbar");
            
            // Set up efficient color monitoring
            this.setupToolbarColorMonitoring(
                toolbarEl, 
                toolbarColor,
                toolbarTextColor, 
                toolbarIconColor,
                toolbarDropdownColor,
                toolbarDropdownTextColor
            );
        }
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Check if we need to rebuild the toolbar configuration
        const shouldRebuildToolbar = this.shouldRebuildToolbar(context);
        
        if (shouldRebuildToolbar) {
            this.createOrUpdateQuill(context);
            return;
        }

        if (!this.quill) return;

        // Apply color changes immediately and efficiently
        this.applyColorStyles(context);

        // Handle content updates
        const newValue = context.parameters.Default?.raw || "";
        
        // Only update Quill if the Default property changed from outside (not from user typing)
        if (
            newValue !== this.lastDefaultValue &&
            newValue !== this.htmlText &&
            newValue !== this.quill.root.innerHTML &&
            !this.isProgrammaticUpdate
        ) {
            if (!this.isValidHTML(newValue)) {
                this.showError("Invalid HTML code. Please correct your input.");
                return;
            } else {
                this.clearError();
            }
            
            this.isProgrammaticUpdate = true;
            try {
                this.htmlText = newValue;
                this.lastDefaultValue = newValue;
                // Use dangerouslyPasteHTML to render HTML directly in Quill
                this.quill.clipboard.dangerouslyPasteHTML(newValue, 'silent');
            } finally {
                this.isProgrammaticUpdate = false;
            }
            
            // Notify that HtmlText output has changed
            if (this.notifyOutputChanged) {
                this.notifyOutputChanged();
            }
        }
    }

    private shouldRebuildToolbar(context: ComponentFramework.Context<IInputs>): boolean {
        // Recompute toolbar config
        let headingLevels: (number | false)[] = [1, 2, 3, 4, false];
        let headingLevelsRaw: string | undefined = undefined;
        if ('headingLevels' in context.parameters && context.parameters['headingLevels'] && typeof context.parameters['headingLevels'].raw === 'string') {
            headingLevelsRaw = context.parameters['headingLevels'].raw as string;
        }
        if (headingLevelsRaw) {
            const parsed = headingLevelsRaw.split(',').map((x: string) => parseInt(x.trim(), 10)).filter((x: number) => !isNaN(x));
            headingLevels = [...parsed, false];
        }
        
        function getToggle(param: keyof IInputs, defaultValue = true) {
            const prop = context.parameters[param];
            if (prop && typeof prop.raw === 'boolean') {
                return prop.raw;
            }
            return defaultValue;
        }
        
        // Declare all toggles before toolbar config
        const showBold = getToggle('showBold');
        const showItalic = getToggle('showItalic');
        const showUnderline = getToggle('showUnderline');
        const showLink = getToggle('showLink');
        const showBlockquote = getToggle('showBlockquote');
        const showCodeBlock = getToggle('showCodeBlock');
        const showImage = getToggle('showImage');
        const showOrderedList = getToggle('showOrderedList');
        const showBulletList = getToggle('showBulletList');
        const showFont = getToggle('showFont');
        const showColor = getToggle('showColor');
        const showAlign = getToggle('showAlign');
        const showClean = getToggle('showClean');
        
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const toolbar: any[] = [];
        toolbar.push([{ header: headingLevels }]);
        if (showFont) toolbar.push([{ font: [] }]);
        const inline: string[] = [];
        if (showBold) inline.push("bold");
        if (showItalic) inline.push("italic");
        if (showUnderline) inline.push("underline");
        if (inline.length) toolbar.push(inline);
        if (showColor) toolbar.push([{ color: [] }, { background: [] }]);
        const block: string[] = [];
        if (showLink) block.push("link");
        if (showBlockquote) block.push("blockquote");
        if (showCodeBlock) block.push("code-block");
        if (showImage) block.push("image");
        if (block.length) toolbar.push(block);
        if (showAlign) toolbar.push([{ align: [] }]);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const lists: any[] = [];
        if (showOrderedList) lists.push({ list: "ordered" });
        if (showBulletList) lists.push({ list: "bullet" });
        if (lists.length) toolbar.push(lists);
        if (showClean) toolbar.push(["clean"]);
        
        // Track all toolbar-related toggles for robust detection
        const toolbarConfig = JSON.stringify(toolbar);
        const togglesConfig = [
            showBold, showItalic, showUnderline, showLink, showBlockquote, showCodeBlock, showImage,
            showOrderedList, showBulletList, showFont, showColor, showAlign, showClean
        ].join("|");
        const headingLevelsRawForCompare = headingLevelsRaw || "";
        
        const lastConfig = this.lastToolbarConfig + "|" + this._lastHeadingLevelsRaw + "|" + (this._lastTogglesConfig || "");
        const newConfig = toolbarConfig + "|" + headingLevelsRawForCompare + "|" + togglesConfig;
        
        if (newConfig !== lastConfig) {
            this.lastToolbarConfig = toolbarConfig;
            this._lastHeadingLevelsRaw = headingLevelsRawForCompare;
            this._lastTogglesConfig = togglesConfig;
            return true;
        }
        
        return false;
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        // Return the HTML content as output
        return {
            HtmlText: this.htmlText
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Cleanup color monitoring
        if (this.container) {
            const toolbarEl = this.container.querySelector('.ql-toolbar') as HTMLElement | null;
            if (toolbarEl) {
                this.cleanupColorMonitoring(toolbarEl);
            }
        }
        
        // Cleanup Quill instance and error div
        if (this.quill) {
            // Quill does not have a destroy method, but we can remove listeners and DOM
            if (this.container) {
                this.container.innerHTML = "";
            }
            this.quill = undefined;
        }
        this.clearError();
    }
}
