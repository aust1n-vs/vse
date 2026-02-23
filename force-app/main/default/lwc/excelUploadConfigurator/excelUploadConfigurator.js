import { LightningElement, track } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import { loadScript } from 'lightning/platformResourceLoader';
import VSSheetJS from '@salesforce/resourceUrl/VSSheetJS';
import getConfigById from '@salesforce/apex/ExcelImportConfigController.getConfigById';
import getConfigMappings from '@salesforce/apex/ExcelImportConfigController.getConfigMappings';
import saveConfig from '@salesforce/apex/ExcelImportConfigController.saveConfig';

const TARGET_OBJECT_OPTIONS = [
    { label: 'Sellthru', value: 'Sellthru__c' },
    { label: 'Stock', value: 'Stock__c' }
];

const SALESFORCE_FIELD_OPTIONS = {
    Sellthru__c: [
        { label: 'Country', value: 'Country__c' },
        { label: 'Distributor', value: 'Distributor__c' },
        { label: 'Reseller', value: 'Reseller__c' },
        { label: 'Product', value: 'Product__c' },
        { label: 'Invoice Date', value: 'Invoice_Date__c' },
        { label: 'Qty', value: 'Qty__c' },
        { label: 'Unit Pirce', value: 'Unit_Price__c' },
        { label: 'Amount', value: 'Amount__c' },
        { label: 'Unit Price(USD)', value: 'Unit_Price_USD__c' },
        { label: 'Amount(USD)', value: 'Amount_USD__c' }
    ],
    Stock__c: [
        { label: 'NS', value: 'NS__c' },
        { label: 'Country', value: 'Country__c' },
        { label: 'Tier(Distributor)', value: 'Tier_Distributor__c' },
        { label: 'Channel Partner', value: 'Channel_Partner__c' },
        { label: 'Date', value: 'Date__c' },
        { label: 'Week', value: 'Week__c' },
        { label: 'Product Model', value: 'Product_Model__c' },
        { label: 'Stock Qty', value: 'Stock_Qty__c' },
        { label: 'Sold Qty', value: 'Sold_Qty__c' },
        { label: 'Open PO', value: 'Open_PO__c' },
        { label: 'Open SO', value: 'Open_SO__c' },
        { label: 'Deliver Stock', value: 'Deliver_Stock__c' },
        { label: 'Ordered Stock', value: 'Ordered_Stock__c' },
        { label: 'Product Name', value: 'Product_Name__c' },
        { label: 'Product code', value: 'Product_Code__c' },
        { label: 'Stock Report', value: 'Stock_Report__c' },
        { label: 'Updateded BY', value: 'Updateded_BY__c' }
    ]
};

const DATA_TYPE_OPTIONS = [
    { label: 'Auto', value: 'Auto' },
    { label: 'Text', value: 'Text' },
    { label: 'Number', value: 'Number' },
    { label: 'Date', value: 'Date' },
    { label: 'DateTime', value: 'DateTime' },
    { label: 'Currency', value: 'Currency' },
    { label: 'Boolean', value: 'Boolean' }
];

function makeExcelColumnOptions(maxColumns) {
    const options = [];
    for (let idx = 0; idx < maxColumns; idx += 1) {
        let n = idx + 1;
        let col = '';
        while (n > 0) {
            const rem = (n - 1) % 26;
            col = String.fromCharCode(65 + rem) + col;
            n = Math.floor((n - rem) / 26);
        }
        options.push({ label: col, value: col });
    }
    return options;
}

const EXCEL_COLUMN_OPTIONS = makeExcelColumnOptions(702);
const HEADER_REFRESH_TRIGGER_FIELDS = ['Sheet_Name__c', 'Header_Row_Number__c', 'Skip_Columns__c'];

export default class ExcelUploadConfigurator extends LightningElement {
    targetObjectOptions = TARGET_OBJECT_OPTIONS;
    dataTypeOptions = DATA_TYPE_OPTIONS;
    excelColumnOptions = EXCEL_COLUMN_OPTIONS;

    @track configForm = this.defaultConfigForm();
    @track mappingRows = [];
    @track parsedHeaders = [];
    @track sheetOptions = [];

    selectedConfigId = '';
    fieldFilter = '';
    uploadedFileName = '';

    errorMessage = '';
    isLoadingMappings = false;
    isSaving = false;
    isParsingFile = false;

    workbook = null;
    sheetJsReady = false;
    sheetJsLoadPromise = null;

    get fullFieldOptions() {
        return SALESFORCE_FIELD_OPTIONS[this.configForm.Target_Object_API_Name__c] || [];
    }

    get filteredFieldOptions() {
        const filter = this.fieldFilter.trim().toLowerCase();
        if (!filter) {
            return this.fullFieldOptions;
        }

        return this.fullFieldOptions.filter(
            (opt) =>
                opt.label.toLowerCase().includes(filter) ||
                opt.value.toLowerCase().includes(filter)
        );
    }

    get hasMappings() {
        return this.mappingRows.length > 0;
    }

    get hasParsedHeaders() {
        return this.parsedHeaders.length > 0;
    }

    get hasSheetOptions() {
        return this.sheetOptions.length > 0;
    }

    get excelColumnPickerOptions() {
        if (!this.hasParsedHeaders) {
            return this.excelColumnOptions;
        }

        const optionMap = new Map();
        this.parsedHeaders.forEach((item) => {
            optionMap.set(item.column, {
                label: `${item.column} - ${item.header}`,
                value: item.column
            });
        });

        this.mappingRows.forEach((row) => {
            const column = String(row.Excel_Column__c || '').toUpperCase();
            if (column && !optionMap.has(column)) {
                optionMap.set(column, { label: column, value: column });
            }
        });

        return Array.from(optionMap.values());
    }

    get canRefreshHeaders() {
        return Boolean(this.workbook) && !this.isParsingFile;
    }

    get disableRefreshHeaders() {
        return !this.canRefreshHeaders;
    }

    get disableSave() {
        return this.isSaving || !String(this.configForm.Name || '').trim();
    }

    get configPickerFilter() {
        return {
            criteria: [
                {
                    fieldPath: 'Target_Object_API_Name__c',
                    operator: 'eq',
                    value: this.configForm.Target_Object_API_Name__c
                }
            ]
        };
    }

    handleTargetObjectChange(event) {
        const newTarget = event.detail.value;
        if (newTarget === this.configForm.Target_Object_API_Name__c) {
            return;
        }

        this.configForm = {
            ...this.configForm,
            Target_Object_API_Name__c: newTarget
        };

        this.fieldFilter = '';
        this.resetSelection(true);
        if (this.hasParsedHeaders) {
            this.mergeMappingRowsFromHeaders(this.parsedHeaders);
        }
    }

    async handleRecordPickerChange(event) {
        this.selectedConfigId = event.detail.recordId || '';
        this.errorMessage = '';

        if (!this.selectedConfigId) {
            this.resetSelection(true);
            return;
        }

        this.isLoadingMappings = true;

        try {
            const selected = await getConfigById({ configId: this.selectedConfigId });

            if (!selected) {
                this.resetSelection(true);
                return;
            }

            this.configForm = {
                Id: selected.Id,
                Name: selected.Name,
                Target_Object_API_Name__c: selected.Target_Object_API_Name__c,
                Sheet_Name__c: selected.Sheet_Name__c || '',
                Skip_Rows__c: selected.Skip_Rows__c == null ? 0 : selected.Skip_Rows__c,
                Skip_Columns__c: selected.Skip_Columns__c == null ? 0 : selected.Skip_Columns__c,
                Header_Row_Number__c: selected.Header_Row_Number__c == null ? 1 : selected.Header_Row_Number__c,
                Data_Start_Row_Number__c:
                    selected.Data_Start_Row_Number__c == null ? 2 : selected.Data_Start_Row_Number__c,
                Active__c: true,
                Notes__c: selected.Notes__c || ''
            };

            if (this.workbook && window.XLSX) {
                this.parsedHeaders = this.extractHeadersFromWorkbook(this.workbook);
            }

            await this.loadMappings(this.selectedConfigId);
            this.syncIncomingHeadersToMappingRows();
        } catch (error) {
            this.mappingRows = [];
            this.errorMessage = this.getErrorMessage(error);
        } finally {
            this.isLoadingMappings = false;
        }
    }

    async loadMappings(configId) {
        const rows = await getConfigMappings({ configId });
        this.mappingRows = rows.map((row, idx) => this.normalizeMappingRow(row, idx));
    }

    handleConfigFieldChange(event) {
        const fieldName = event.target.dataset.field;
        const value = event.detail ? event.detail.value : event.target.value;

        this.configForm = {
            ...this.configForm,
            [fieldName]: value
        };

        if (this.workbook && HEADER_REFRESH_TRIGGER_FIELDS.includes(fieldName)) {
            this.rebuildHeadersFromWorkbook();
        }
    }

    handleFieldFilterChange(event) {
        this.fieldFilter = event.detail.value || '';
    }

    async handleExcelFileChange(event) {
        const [file] = event.target.files || [];
        event.target.value = '';

        if (!file) {
            return;
        }

        this.errorMessage = '';
        this.isParsingFile = true;
        this.uploadedFileName = file.name;

        try {
            await this.ensureSheetJsLoaded();
            const fileBuffer = await this.readFileAsArrayBuffer(file);
            const workbook = window.XLSX.read(fileBuffer, {
                type: 'array',
                raw: false
            });

            this.workbook = workbook;
            this.sheetOptions = (workbook.SheetNames || []).map((name) => ({
                label: name,
                value: name
            }));

            const selectedSheet = this.resolveSheetName(workbook);
            this.configForm = {
                ...this.configForm,
                Sheet_Name__c: selectedSheet
            };

            this.rebuildHeadersFromWorkbook();
            this.showToast('File parsed', `Loaded headers from ${file.name}`, 'success');
        } catch (error) {
            this.workbook = null;
            this.sheetOptions = [];
            this.parsedHeaders = [];
            this.mappingRows = [];
            this.errorMessage = this.getErrorMessage(error);
            this.showToast('Parse failed', this.errorMessage, 'error');
        } finally {
            this.isParsingFile = false;
        }
    }

    handleRefreshHeaders() {
        this.rebuildHeadersFromWorkbook();
    }

    handleAddMappingRow() {
        const newRow = this.normalizeMappingRow(
            {
                Sequence__c: this.nextSequence(),
                Excel_Column__c: '',
                Excel_Column_Index__c: null,
                Salesforce_Field_API_Name__c: '',
                Salesforce_Field_Label__c: '',
                Data_Type__c: 'Auto',
                Default_Value__c: '',
                Trim_Value__c: true,
                Ignore_If_Blank__c: false,
                Is_Enabled__c: true
            },
            this.mappingRows.length
        );

        this.mappingRows = [...this.mappingRows, newRow];
    }

    handleRemoveMappingRow(event) {
        const rowKey = event.currentTarget.dataset.rowKey;
        this.mappingRows = this.mappingRows.filter((row) => row._rowKey !== rowKey);
    }

    handleRowValueChange(event) {
        const rowKey = event.target.dataset.rowKey;
        const fieldName = event.target.dataset.field;

        if (!rowKey || !fieldName) {
            return;
        }

        let value;
        if (event.target.type === 'checkbox') {
            value = event.target.checked;
        } else {
            value = event.detail ? event.detail.value : event.target.value;
        }

        if (fieldName === 'Excel_Column__c') {
            value = String(value || '').toUpperCase().trim();
        }

        if (fieldName === 'Sequence__c' || fieldName === 'Excel_Column_Index__c') {
            value = this.normalizeNumber(value, null);
        }

        this.mappingRows = this.mappingRows.map((row) => {
            if (row._rowKey !== rowKey) {
                return row;
            }

            const updated = {
                ...row,
                [fieldName]: value
            };

            if (fieldName === 'Excel_Column__c') {
                const header = this.findHeaderByColumn(value);
                updated.Incoming_Header__c = header ? header.header : '';
                updated.Excel_Column_Index__c = header ? header.columnIndex : null;
                if (header) {
                    updated.Sequence__c = header.columnIndex;
                }
            }

            if (fieldName === 'Salesforce_Field_API_Name__c') {
                const option = this.fullFieldOptions.find((opt) => opt.value === value);
                updated.Salesforce_Field_Label__c = option ? option.label : '';
            }

            return this.withCardTitle(updated);
        });
    }

    async handleSave() {
        this.isSaving = true;
        this.errorMessage = '';

        const payloadConfig = {
            Id: this.configForm.Id,
            Name: this.configForm.Name,
            Target_Object_API_Name__c: this.configForm.Target_Object_API_Name__c,
            Sheet_Name__c: this.configForm.Sheet_Name__c,
            Skip_Rows__c: this.normalizeNumber(this.configForm.Skip_Rows__c, 0),
            Skip_Columns__c: this.normalizeNumber(this.configForm.Skip_Columns__c, 0),
            Header_Row_Number__c: this.normalizeNumber(this.configForm.Header_Row_Number__c, 1),
            Data_Start_Row_Number__c: this.normalizeNumber(this.configForm.Data_Start_Row_Number__c, 2),
            Active__c: true,
            Notes__c: this.configForm.Notes__c
        };

        const payloadMappings = this.mappingRows
            .map((row) => ({
                Sequence__c: this.normalizeNumber(row.Sequence__c, null),
                Excel_Column__c: String(row.Excel_Column__c || '').trim().toUpperCase(),
                Excel_Column_Index__c: this.normalizeNumber(row.Excel_Column_Index__c, null),
                Salesforce_Field_API_Name__c: String(row.Salesforce_Field_API_Name__c || '').trim(),
                Salesforce_Field_Label__c: String(row.Salesforce_Field_Label__c || '').trim(),
                Data_Type__c: row.Data_Type__c || 'Auto',
                Default_Value__c: row.Default_Value__c,
                Trim_Value__c: row.Trim_Value__c,
                Ignore_If_Blank__c: row.Ignore_If_Blank__c,
                Is_Enabled__c: row.Is_Enabled__c
            }))
            .filter((row) => row.Excel_Column__c)
            .filter((row) => row.Salesforce_Field_API_Name__c);

        try {
            const result = await saveConfig({
                config: payloadConfig,
                mappings: payloadMappings,
                saveAsNew: false
            });

            this.configForm = {
                ...this.defaultConfigForm(),
                ...result.config,
                Active__c: true
            };

            this.selectedConfigId = result.config.Id;
            this.mappingRows = (result.mappings || []).map((row, idx) => this.normalizeMappingRow(row, idx));

            this.showToast('Saved', 'Config updated successfully.', 'success');
        } catch (error) {
            this.errorMessage = this.getErrorMessage(error);
            this.showToast('Save failed', this.errorMessage, 'error');
        } finally {
            this.isSaving = false;
        }
    }

    normalizeMappingRow(rawRow, idx) {
        return this.withCardTitle({
            _rowKey: rawRow.Id || `row_${idx}_${Date.now()}`,
            Id: rawRow.Id,
            Sequence__c: this.normalizeNumber(rawRow.Sequence__c, idx + 1),
            Excel_Column__c: String(rawRow.Excel_Column__c || '').toUpperCase(),
            Excel_Column_Index__c: this.normalizeNumber(rawRow.Excel_Column_Index__c, null),
            Incoming_Header__c: rawRow.Incoming_Header__c || '',
            Salesforce_Field_API_Name__c: rawRow.Salesforce_Field_API_Name__c || '',
            Salesforce_Field_Label__c: rawRow.Salesforce_Field_Label__c || '',
            Data_Type__c: rawRow.Data_Type__c || 'Auto',
            Default_Value__c: rawRow.Default_Value__c || '',
            Trim_Value__c: rawRow.Trim_Value__c == null ? true : rawRow.Trim_Value__c,
            Ignore_If_Blank__c: rawRow.Ignore_If_Blank__c == null ? false : rawRow.Ignore_If_Blank__c,
            Is_Enabled__c: rawRow.Is_Enabled__c == null ? true : rawRow.Is_Enabled__c
        });
    }

    nextSequence() {
        if (!this.mappingRows.length) {
            return 1;
        }

        return Math.max(...this.mappingRows.map((row) => this.normalizeNumber(row.Sequence__c, 0))) + 1;
    }

    normalizeNumber(value, fallbackValue) {
        if (value === null || value === undefined || value === '') {
            return fallbackValue;
        }

        const parsed = Number(value);
        return Number.isNaN(parsed) ? fallbackValue : parsed;
    }

    defaultConfigForm() {
        return {
            Id: null,
            Name: '',
            Target_Object_API_Name__c: 'Sellthru__c',
            Sheet_Name__c: '',
            Skip_Rows__c: 0,
            Skip_Columns__c: 0,
            Header_Row_Number__c: 1,
            Data_Start_Row_Number__c: 2,
            Active__c: true,
            Notes__c: ''
        };
    }

    resetSelection(keepFormTarget) {
        const target = keepFormTarget ? this.configForm.Target_Object_API_Name__c : 'Sellthru__c';

        this.selectedConfigId = '';
        this.mappingRows = [];
        this.configForm = {
            ...this.defaultConfigForm(),
            Target_Object_API_Name__c: target,
            Active__c: true
        };
    }

    async ensureSheetJsLoaded() {
        if (window.XLSX) {
            this.sheetJsReady = true;
            return;
        }

        if (this.sheetJsReady && window.XLSX) {
            return;
        }

        if (!this.sheetJsLoadPromise) {
            this.sheetJsLoadPromise = this.loadFirstAvailableScript([
                VSSheetJS,
                `${VSSheetJS}/xlsx.full.min.js`,
                `${VSSheetJS}/dist/xlsx.full.min.js`,
                `${VSSheetJS}/xlsx.min.js`
            ]).finally(() => {
                this.sheetJsLoadPromise = null;
            });
        }

        await this.sheetJsLoadPromise;

        if (!window.XLSX) {
            throw new Error('SheetJS did not load. Verify static resource VSSheetJS path.');
        }

        this.sheetJsReady = true;
    }

    async loadFirstAvailableScript(urls) {
        let lastError;

        for (const url of urls) {
            try {
                await loadScript(this, url);
                if (window.XLSX) {
                    return;
                }
            } catch (error) {
                lastError = error;
            }
        }

        throw lastError || new Error('Unable to load SheetJS static resource.');
    }

    readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => reject(new Error('Unable to read uploaded file.'));
            reader.readAsArrayBuffer(file);
        });
    }

    resolveSheetName(workbook) {
        const names = workbook && workbook.SheetNames ? workbook.SheetNames : [];
        const current = String(this.configForm.Sheet_Name__c || '').trim();

        if (current && names.includes(current)) {
            return current;
        }

        return names.length ? names[0] : '';
    }

    rebuildHeadersFromWorkbook() {
        if (!this.workbook || !window.XLSX) {
            return;
        }

        const parsed = this.extractHeadersFromWorkbook(this.workbook);
        this.parsedHeaders = parsed;
        this.mergeMappingRowsFromHeaders(parsed);
    }

    extractHeadersFromWorkbook(workbook) {
        const sheetName = this.resolveSheetName(workbook);
        if (!sheetName) {
            return [];
        }

        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
            return [];
        }

        const allRows = window.XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            raw: false,
            defval: '',
            blankrows: false
        });

        const headerRowNumber = Math.max(1, this.normalizeNumber(this.configForm.Header_Row_Number__c, 1));
        const skipColumns = Math.max(0, this.normalizeNumber(this.configForm.Skip_Columns__c, 0));
        const headerRow = Array.isArray(allRows[headerRowNumber - 1]) ? allRows[headerRowNumber - 1] : [];
        const headers = [];

        for (let idx = skipColumns; idx < headerRow.length; idx += 1) {
            const rawHeader = headerRow[idx];
            const text = String(rawHeader || '').trim();
            if (!text) {
                continue;
            }

            headers.push({
                column: this.toExcelColumn(idx + 1),
                columnIndex: idx + 1,
                header: text
            });
        }

        return headers;
    }

    mergeMappingRowsFromHeaders(headers) {
        const existingByColumn = new Map(
            this.mappingRows.map((row) => [String(row.Excel_Column__c || '').toUpperCase(), row])
        );

        this.mappingRows = headers.map((item, idx) => {
            const existing = existingByColumn.get(item.column);
            const autoMatch = this.findAutoMatchedField(item.header);
            const merged = {
                Sequence__c: existing ? existing.Sequence__c : item.columnIndex,
                Excel_Column__c: item.column,
                Excel_Column_Index__c: item.columnIndex,
                Incoming_Header__c: item.header,
                Salesforce_Field_API_Name__c:
                    existing && existing.Salesforce_Field_API_Name__c
                        ? existing.Salesforce_Field_API_Name__c
                        : autoMatch
                          ? autoMatch.value
                          : '',
                Salesforce_Field_Label__c:
                    existing && existing.Salesforce_Field_Label__c
                        ? existing.Salesforce_Field_Label__c
                        : autoMatch
                          ? autoMatch.label
                          : '',
                Data_Type__c: existing ? existing.Data_Type__c : 'Auto',
                Default_Value__c: existing ? existing.Default_Value__c : '',
                Trim_Value__c: existing ? existing.Trim_Value__c : true,
                Ignore_If_Blank__c: existing ? existing.Ignore_If_Blank__c : false,
                Is_Enabled__c: existing ? existing.Is_Enabled__c : true
            };

            return this.normalizeMappingRow(merged, idx);
        });
    }

    syncIncomingHeadersToMappingRows() {
        this.mappingRows = this.mappingRows.map((row) => {
            const header = this.findHeaderByColumn(row.Excel_Column__c);
            return this.withCardTitle({
                ...row,
                Incoming_Header__c: header ? header.header : row.Incoming_Header__c,
                Excel_Column_Index__c: header ? header.columnIndex : row.Excel_Column_Index__c
            });
        });
    }

    withCardTitle(row) {
        const sequence = this.normalizeNumber(row.Sequence__c, null);
        const columnIndex = this.normalizeNumber(row.Excel_Column_Index__c, null);
        const incomingHeader = String(row.Incoming_Header__c || '').trim();
        const excelColumn = String(row.Excel_Column__c || '').trim();
        const displayIndex = columnIndex !== null && columnIndex !== undefined ? columnIndex : sequence;

        let title;
        if (displayIndex !== null && displayIndex !== undefined && incomingHeader) {
            title = `#${displayIndex} - ${incomingHeader}`;
        } else if (displayIndex !== null && displayIndex !== undefined && excelColumn) {
            title = `#${displayIndex} - ${excelColumn}`;
        } else if (displayIndex !== null && displayIndex !== undefined) {
            title = `#${displayIndex}`;
        } else if (incomingHeader) {
            title = incomingHeader;
        } else {
            title = 'New Mapping';
        }

        return {
            ...row,
            Card_Title__c: title
        };
    }

    findHeaderByColumn(columnName) {
        const key = String(columnName || '').toUpperCase();
        return this.parsedHeaders.find((item) => item.column === key);
    }

    findAutoMatchedField(headerText) {
        const headerKey = this.normalizeHeaderKey(headerText);
        return this.fullFieldOptions.find(
            (option) =>
                this.normalizeHeaderKey(option.label) === headerKey ||
                this.normalizeHeaderKey(option.value) === headerKey
        );
    }

    normalizeHeaderKey(value) {
        return String(value || '')
            .toLowerCase()
            .replace(/[^a-z0-9]/g, '');
    }

    toExcelColumn(columnIndex) {
        let n = columnIndex;
        let col = '';

        while (n > 0) {
            const rem = (n - 1) % 26;
            col = String.fromCharCode(65 + rem) + col;
            n = Math.floor((n - rem) / 26);
        }

        return col;
    }

    showToast(title, message, variant) {
        this.dispatchEvent(
            new ShowToastEvent({
                title,
                message,
                variant
            })
        );
    }

    getErrorMessage(error) {
        if (!error) {
            return 'Unexpected error.';
        }

        if (Array.isArray(error.body)) {
            return error.body.map((item) => item.message).join(', ');
        }

        if (error.body && error.body.message) {
            return error.body.message;
        }

        if (error.message) {
            return error.message;
        }

        return 'Unexpected error.';
    }
}
