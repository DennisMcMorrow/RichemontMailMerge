using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using static RichemontMailMerge.Form1;

public class NewHire
{
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public DateTime DateOfHire { get; set; }
    public string Position { get; set; }
    // Add other relevant properties as needed
}

public class NewHireProcessor
{
    public List<string> Errors { get; private set; } = new List<string>();
    public string InputCsvPath { get; set; }
    public string OutputExcelPath { get; set; }
    public string OutputWordPath { get; set; }

    public List<Record> ReadCsvData(string filePath)
    {
        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            var records = csv.GetRecords<Record>().ToList();
            return records;
        }
    }

    public void GenerateMailMergeExcel(string inputFilePath, string outputFilePath)
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook inputWorkbook = excelApp.Workbooks.Open(inputFilePath);
        Excel.Worksheet inputWorksheet = inputWorkbook.Sheets[1];
        // ... (rest of the logic to generate the Excel file for mail merge)
        inputWorkbook.Close(false);
        excelApp.Quit();
    }

    public void GenerateWordDocument(string excelDataSourcePath, string templatePath, string outputWordFilePath)
    {
        Word.Application wordApp = new Word.Application();
        Word.Document document = wordApp.Documents.Open(templatePath);
        document.MailMerge.MainDocumentType = Word.WdMailMergeMainDocType.wdFormLetters;
        document.MailMerge.OpenDataSource(excelDataSourcePath);
        document.MailMerge.Execute(Pause: true);
        wordApp.ActiveDocument.SaveAs2(outputWordFilePath);
        wordApp.ActiveDocument.Close();
        wordApp.Quit();
    }

    public bool ProcessNewHireData(List<NewHire> newHires)
    {
        Errors.Clear();

        foreach (var hire in newHires)
        {
            if (!ValidateNewHire(hire))
            {
                Errors.Add($"Validation failed for {hire.FirstName} {hire.LastName}");
                continue;
            }

            // Store the validated new hire data
            // This can be in a data structure, database, etc.
            StoreNewHireData(hire);

            // Generate any required reports or notifications
            GenerateReports(hire);
        }

        return !Errors.Any();
    }

    private bool ValidateNewHire(NewHire hire)
    {
        // Implement validation logic here
        // For example, check if all properties have values, if the date of hire is valid, etc.
        if (string.IsNullOrEmpty(hire.FirstName) || string.IsNullOrEmpty(hire.LastName))
        {
            return false;
        }

        // Add other validation checks as needed

        return true;
    }

    private void StoreNewHireData(NewHire hire)
    {
        // Implement logic to store the new hire data
        // This can be saving to a database, adding to a list, etc.
    }

    private void GenerateReports(NewHire hire)
    {
        // Implement logic to generate reports or notifications for the new hire
        // This can be sending an email, generating a PDF report, etc.
    }
}
