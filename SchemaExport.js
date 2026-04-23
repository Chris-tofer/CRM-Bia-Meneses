function exportSchema() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        if (!ss) {
            // If we are running as a standalone script or container-bound, we might need
            // to open by ID if getActiveSpreadsheet() returns null in certain execution contexts (like API Executable).
            // However, usually clasp run on a container-bound script works fine.
            // If execute via API, we might need explicitly open by ID if it's not bound correctly in context.
            // Let's assume standard container-bound behavior first.
            return JSON.stringify({ error: "No active spreadsheet found." });
        }

        var sheets = ss.getSheets();
        var schema = {};

        sheets.forEach(function (sheet) {
            var name = sheet.getName();
            var lastRow = sheet.getLastRow();
            var lastCol = sheet.getLastColumn();

            var sheetData = {
                headers: [],
                sample: []
            };

            if (lastRow > 0 && lastCol > 0) {
                // Get headers
                var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
                sheetData.headers = headers;

                // Get sample data (up to 10 rows)
                if (lastRow > 1) {
                    var numRows = Math.min(10, lastRow - 1);
                    var sampleValues = sheet.getRange(2, 1, numRows, lastCol).getValues();
                    sheetData.sample = sampleValues;
                }
            }

            schema[name] = sheetData;
        });

        var json = JSON.stringify(schema, null, 2);

        // Create file in Drive Root
        var file = DriveApp.createFile('schema_dump.json', json);
        console.log("Arquivo salvo no Drive: " + file.getUrl());
        return file.getUrl();

    } catch (e) {
        var errorJson = JSON.stringify({ error: e.toString(), stack: e.stack });
        console.error(errorJson);
        return errorJson;
    }
}
