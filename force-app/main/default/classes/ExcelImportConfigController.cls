public with sharing class ExcelImportConfigController {
    private static final Integer DEFAULT_RECENT_LIMIT = 10;
    private static final Integer SEARCH_LIMIT = 50;

    public class SaveResult {
        @AuraEnabled public Excel_Import_Config__c config;
        @AuraEnabled public List<Excel_Column_Mapping__c> mappings;
    }

    @AuraEnabled(cacheable=true)
    public static List<Excel_Import_Config__c> getConfigs(String searchTerm, String targetObjectApiName) {
        String trimmedSearch = (searchTerm == null) ? '' : searchTerm.trim();
        String trimmedTarget = (targetObjectApiName == null) ? '' : targetObjectApiName.trim();
        String prefix = '';

        String selectClause =
            'SELECT Id, Name, Target_Object_API_Name__c, Sheet_Name__c, Skip_Rows__c, Skip_Columns__c, ' +
            'Header_Row_Number__c, Data_Start_Row_Number__c, Active__c, Notes__c, LastModifiedDate ' +
            'FROM Excel_Import_Config__c';

        List<String> whereParts = new List<String>();

        if (!String.isBlank(trimmedTarget)) {
            whereParts.add('Target_Object_API_Name__c = :trimmedTarget');
        }

        if (!String.isBlank(trimmedSearch)) {
            prefix = trimmedSearch + '%';
            whereParts.add('Name LIKE :prefix');
        }

        String soql = selectClause;
        if (!whereParts.isEmpty()) {
            soql += ' WHERE ' + String.join(whereParts, ' AND ');
        }
        soql += ' ORDER BY LastModifiedDate DESC';
        soql += String.isBlank(trimmedSearch) ? ' LIMIT :DEFAULT_RECENT_LIMIT' : ' LIMIT :SEARCH_LIMIT';

        try {
            return Database.query(soql);
        } catch (Exception ex) {
            throw new AuraHandledException('Failed to load configs: ' + ex.getMessage());
        }
    }

    @AuraEnabled(cacheable=true)
    public static List<Excel_Column_Mapping__c> getConfigMappings(Id configId) {
        if (configId == null) {
            return new List<Excel_Column_Mapping__c>();
        }

        try {
            return [
                SELECT Id, Config__c, Sequence__c, Excel_Column__c, Excel_Column_Index__c,
                       Salesforce_Field_API_Name__c, Salesforce_Field_Label__c, Data_Type__c,
                       Default_Value__c, Trim_Value__c, Ignore_If_Blank__c, Is_Enabled__c
                FROM Excel_Column_Mapping__c
                WHERE Config__c = :configId
                ORDER BY Sequence__c ASC, CreatedDate ASC
            ];
        } catch (Exception ex) {
            throw new AuraHandledException('Failed to load mappings: ' + ex.getMessage());
        }
    }

    @AuraEnabled(cacheable=true)
    public static Excel_Import_Config__c getConfigById(Id configId) {
        if (configId == null) {
            return null;
        }

        try {
            return [
                SELECT Id, Name, Target_Object_API_Name__c, Sheet_Name__c, Skip_Rows__c, Skip_Columns__c,
                       Header_Row_Number__c, Data_Start_Row_Number__c, Active__c, Notes__c, LastModifiedDate
                FROM Excel_Import_Config__c
                WHERE Id = :configId
                LIMIT 1
            ];
        } catch (Exception ex) {
            throw new AuraHandledException('Failed to load config detail: ' + ex.getMessage());
        }
    }

    @AuraEnabled
    public static SaveResult saveConfig(
        Excel_Import_Config__c config,
        List<Excel_Column_Mapping__c> mappings,
        Boolean saveAsNew
    ) {
        if (config == null) {
            throw new AuraHandledException('Config payload is required.');
        }

        Boolean createNew = (saveAsNew == true) || config.Id == null;

        Excel_Import_Config__c configToSave = new Excel_Import_Config__c();
        if (!createNew) {
            configToSave.Id = config.Id;
        }

        configToSave.Name = config.Name;
        configToSave.Target_Object_API_Name__c = config.Target_Object_API_Name__c;
        configToSave.Sheet_Name__c = config.Sheet_Name__c;
        configToSave.Skip_Rows__c = config.Skip_Rows__c;
        configToSave.Skip_Columns__c = config.Skip_Columns__c;
        configToSave.Header_Row_Number__c = config.Header_Row_Number__c;
        configToSave.Data_Start_Row_Number__c = config.Data_Start_Row_Number__c;
        configToSave.Active__c = config.Active__c;
        configToSave.Notes__c = config.Notes__c;

        List<Excel_Column_Mapping__c> inputMappings = (mappings == null)
            ? new List<Excel_Column_Mapping__c>()
            : mappings;

        try {
            if (createNew) {
                insert configToSave;
            } else {
                update configToSave;
            }

            Id savedConfigId = configToSave.Id;

            if (!createNew) {
                List<Excel_Column_Mapping__c> existingRows = [
                    SELECT Id
                    FROM Excel_Column_Mapping__c
                    WHERE Config__c = :savedConfigId
                ];
                if (!existingRows.isEmpty()) {
                    delete existingRows;
                }
            }

            List<Excel_Column_Mapping__c> rowsToInsert = new List<Excel_Column_Mapping__c>();
            for (Excel_Column_Mapping__c row : inputMappings) {
                Excel_Column_Mapping__c newRow = new Excel_Column_Mapping__c();
                newRow.Config__c = savedConfigId;
                newRow.Sequence__c = row.Sequence__c;
                newRow.Excel_Column__c = row.Excel_Column__c;
                newRow.Excel_Column_Index__c = row.Excel_Column_Index__c;
                newRow.Salesforce_Field_API_Name__c = row.Salesforce_Field_API_Name__c;
                newRow.Salesforce_Field_Label__c = row.Salesforce_Field_Label__c;
                newRow.Data_Type__c = row.Data_Type__c;
                newRow.Default_Value__c = row.Default_Value__c;
                newRow.Trim_Value__c = row.Trim_Value__c;
                newRow.Ignore_If_Blank__c = row.Ignore_If_Blank__c;
                newRow.Is_Enabled__c = row.Is_Enabled__c;
                rowsToInsert.add(newRow);
            }

            if (!rowsToInsert.isEmpty()) {
                insert rowsToInsert;
            }

            SaveResult result = new SaveResult();
            result.config = [
                SELECT Id, Name, Target_Object_API_Name__c, Sheet_Name__c, Skip_Rows__c, Skip_Columns__c,
                       Header_Row_Number__c, Data_Start_Row_Number__c, Active__c, Notes__c, LastModifiedDate
                FROM Excel_Import_Config__c
                WHERE Id = :savedConfigId
                LIMIT 1
            ];
            result.mappings = [
                SELECT Id, Config__c, Sequence__c, Excel_Column__c, Excel_Column_Index__c,
                       Salesforce_Field_API_Name__c, Salesforce_Field_Label__c, Data_Type__c,
                       Default_Value__c, Trim_Value__c, Ignore_If_Blank__c, Is_Enabled__c
                FROM Excel_Column_Mapping__c
                WHERE Config__c = :savedConfigId
                ORDER BY Sequence__c ASC, CreatedDate ASC
            ];
            return result;
        } catch (DmlException dmlEx) {
            throw new AuraHandledException('Failed to save config: ' + dmlEx.getDmlMessage(0));
        } catch (Exception ex) {
            throw new AuraHandledException('Failed to save config: ' + ex.getMessage());
        }
    }
}
