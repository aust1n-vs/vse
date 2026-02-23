@IsTest
private class ExcelImportConfigControllerTest {
    private static Excel_Import_Config__c makeConfig(String name, String target) {
        return new Excel_Import_Config__c(
            Name = name,
            Target_Object_API_Name__c = target,
            Sheet_Name__c = 'Sheet1',
            Skip_Rows__c = 0,
            Skip_Columns__c = 0,
            Header_Row_Number__c = 1,
            Data_Start_Row_Number__c = 2,
            Active__c = true,
            Notes__c = 'test'
        );
    }

    private static Excel_Column_Mapping__c makeMapping(
        Decimal seq,
        String excelColumn,
        String sfField,
        String sfLabel
    ) {
        return new Excel_Column_Mapping__c(
            Sequence__c = seq,
            Excel_Column__c = excelColumn,
            Salesforce_Field_API_Name__c = sfField,
            Salesforce_Field_Label__c = sfLabel,
            Data_Type__c = 'Auto',
            Default_Value__c = null,
            Trim_Value__c = true,
            Ignore_If_Blank__c = false,
            Is_Enabled__c = true
        );
    }

    @IsTest
    static void testGetConfigsRecentAndSearch() {
        List<Excel_Import_Config__c> rows = new List<Excel_Import_Config__c>();
        for (Integer i = 0; i < 12; i++) {
            rows.add(makeConfig('Sell Config ' + i, 'Sellthru__c'));
        }
        insert rows;

        Test.startTest();
        List<Excel_Import_Config__c> recent = ExcelImportConfigController.getConfigs('', 'Sellthru__c');
        List<Excel_Import_Config__c> searched = ExcelImportConfigController.getConfigs('Sell Config 1', 'Sellthru__c');
        Test.stopTest();

        System.assertEquals(10, recent.size(), 'Empty search should return the most recent 10 rows.');
        System.assert(searched.size() >= 1, 'Prefix search should find at least one row.');
        System.assertEquals('Sellthru__c', searched[0].Target_Object_API_Name__c);
    }

    @IsTest
    static void testGetConfigMappings() {
        Excel_Import_Config__c cfg = makeConfig('Mapping Fetch', 'Stock__c');
        insert cfg;

        insert new List<Excel_Column_Mapping__c>{
            new Excel_Column_Mapping__c(
                Config__c = cfg.Id,
                Sequence__c = 1,
                Excel_Column__c = 'A',
                Salesforce_Field_API_Name__c = 'NS__c',
                Salesforce_Field_Label__c = 'NS',
                Data_Type__c = 'Auto',
                Trim_Value__c = true,
                Ignore_If_Blank__c = false,
                Is_Enabled__c = true
            ),
            new Excel_Column_Mapping__c(
                Config__c = cfg.Id,
                Sequence__c = 2,
                Excel_Column__c = 'B',
                Salesforce_Field_API_Name__c = 'Country__c',
                Salesforce_Field_Label__c = 'Country',
                Data_Type__c = 'Auto',
                Trim_Value__c = true,
                Ignore_If_Blank__c = false,
                Is_Enabled__c = true
            )
        };

        Test.startTest();
        Excel_Import_Config__c loaded = ExcelImportConfigController.getConfigById(cfg.Id);
        List<Excel_Column_Mapping__c> result = ExcelImportConfigController.getConfigMappings(cfg.Id);
        Test.stopTest();

        System.assertEquals(cfg.Id, loaded.Id);
        System.assertEquals(2, result.size());
        System.assertEquals('A', result[0].Excel_Column__c);
    }

    @IsTest
    static void testSaveConfigUpdateAndSaveAsNew() {
        Excel_Import_Config__c cfg = makeConfig('Base Config', 'Sellthru__c');
        insert cfg;

        insert new Excel_Column_Mapping__c(
            Config__c = cfg.Id,
            Sequence__c = 1,
            Excel_Column__c = 'A',
            Salesforce_Field_API_Name__c = 'Country__c',
            Salesforce_Field_Label__c = 'Country',
            Data_Type__c = 'Auto',
            Trim_Value__c = true,
            Ignore_If_Blank__c = false,
            Is_Enabled__c = true
        );

        Excel_Import_Config__c updatePayload = makeConfig('Base Config Updated', 'Sellthru__c');
        updatePayload.Id = cfg.Id;
        updatePayload.Skip_Rows__c = 2;

        List<Excel_Column_Mapping__c> newMappings = new List<Excel_Column_Mapping__c>{
            makeMapping(1, 'A', 'Distributor__c', 'Distributor'),
            makeMapping(2, 'B', 'Reseller__c', 'Reseller')
        };

        Test.startTest();
        ExcelImportConfigController.SaveResult updated = ExcelImportConfigController.saveConfig(
            updatePayload,
            newMappings,
            false
        );
        Test.stopTest();

        System.assertEquals(cfg.Id, updated.config.Id);
        System.assertEquals('Base Config Updated', updated.config.Name);
        System.assertEquals(2, updated.mappings.size());
        System.assertEquals('Distributor__c', updated.mappings[0].Salesforce_Field_API_Name__c);

        Excel_Import_Config__c clonePayload = makeConfig('Base Config Copy', 'Sellthru__c');

        Test.startTest();
        ExcelImportConfigController.SaveResult created = ExcelImportConfigController.saveConfig(
            clonePayload,
            newMappings,
            true
        );
        Test.stopTest();

        System.assertNotEquals(cfg.Id, created.config.Id, 'Save as new should create a new config.');
        System.assertEquals(2, created.mappings.size());
    }
}
