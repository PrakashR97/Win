DataTable sheets =
    excelConn.GetOleDbSchemaTable(
        OleDbSchemaGuid.Tables,
        null);

foreach (DataRow row in sheets.Rows)
{
    string sheetName =
        row["TABLE_NAME"].ToString();

    Console.WriteLine(sheetName);

    // Processed tab
    if (sheetName.Contains("Processed$"))
    {
        LoadSheetData(
            excelConn,
            oracleConn,
            sheetName,
            "PROCESSED_TABLE",
            fileName);
    }

    // Unprocessed tab
    else if (sheetName.Contains("Unprocessed$"))
    {
        LoadSheetData(
            excelConn,
            oracleConn,
            sheetName,
            "UNPROCESSED_TABLE",
            fileName);
    }
}
