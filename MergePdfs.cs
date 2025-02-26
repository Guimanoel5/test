#r "PDFSharp.dll"
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.IO;
using System.Collections.Generic;
using Vinyl.Sdk;
using Vinyl.Sdk.Events;
using Vinyl.Sdk.Filtering;
using Microsoft.Extensions.Logging;

var logger = Services.GetRequiredService<ILogger>();

string resultingFilename = "merged.pdf";

try
{
    resultingFilename = Row["MergedFilename"].GetValueAsString();
}catch
{
    logger.LogWarning("MergedFilename parameter column not found, using default filename merged.pdf");
}

if (resultingFilename == "")
{
    logger.LogWarning("MergedFilename parameter column empty, using default filename merged.pdf");
    resultingFilename = "merged.pdf";
}

Guid inputTableId = Row["InputTableId"].GetValueAsGuid();

Guid filterColumn = Row["ManagedBindingID"].GetValueAsGuid();

var tableService = Services.GetService<ITableService>();
var eventService = Services.GetService<IEventService>();

var inputTableFilter = Services.GetService<FilterBuilder>().From(inputTableId);

inputTableFilter.Filter.Limit = 1000;
// Where cause that considers the binding
inputTableFilter.Where("ManagedBindingID", ComparisonOperator.Equals, filterColumn);

// This is how you order a table by an especific column
inputTableFilter.Filter.Sorting.Add("Seq nr", SortDirection.Ascending);

EventTable inputTable = await eventService.InvokeFilterEventAsync(inputTableFilter);

List<Dictionary<string,object>> pdfsToBeMergedOrdered = new List<Dictionary<string,object>>();

int howManyRecords = inputTable.Rows.Count;
logger.LogWarning($"Records recieved: {howManyRecords.ToString()}");

foreach(var row in inputTable.Rows)
{
    int seq = row["Seq nr"].GetValueAsInteger();
    
    byte[] pdfContent = row["PDF Content"].GetValueAsByteArray();
    
    if (pdfContent.Length == 0)
    {
        logger.LogError($"Missing pdf content on Row of sequence: {seq}");
        continue;
    }
    string fileName = "";
    try
    {
        fileName = row["Filename"].GetValueAsString();    
    } 
    catch (Exception ex)
    {
        logger.LogWarning($"File name not found: {ex.Message}");    
    }
    

    var pdfEntry = new Dictionary<string, object>{
        {"fileName", fileName},
        {"pdfContent", pdfContent},
        {"sequence", seq}
    };
    
    pdfsToBeMergedOrdered.Add(pdfEntry);
}

Guid outputTableId = Row["OutputTableId"].GetValueAsGuid();

var outputTableFilter = Services.GetService<FilterBuilder>().From(outputTableId);

outputTableFilter.Filter.Limit = 1;

EventTable outputTable = await eventService.InvokeFilterEventAsync(outputTableFilter);

byte[] resultingPdf = MergePDFs(pdfsToBeMergedOrdered);

EventRow newRow = await eventService.InvokeNewEventAsync(outputTable);

newRow["Filename"].Value = resultingFilename;
newRow["Content"].Value = resultingPdf;
newRow["ManagedBindingID"].Value = filterColumn;

await eventService.InvokeInsertEventAsync(newRow);

logger.LogWarning("Added merged pdf to table!");

public static byte[] MergePDFs(List<Dictionary<string, object>> pdfEntries)
{
    using (var mergedDocument = new PdfDocument())
    {
        foreach (var entry in pdfEntries)
        {
            // Extract filename, sequence number, and PDF content from the dictionary
            string fileName = (string)entry["fileName"];
            byte[] pdfBytes = (byte[])entry["pdfContent"];
            int sequence = (int)entry["sequence"];

            try
            {
                using (var stream = new MemoryStream(pdfBytes))
                {
                    var sourceDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Import);

                    for (int i = 0; i < sourceDocument.PageCount; i++)
                    {
                        mergedDocument.AddPage(sourceDocument.Pages[i]);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error processing file '{fileName}' (Sequence: {sequence}): {ex.Message}");
            }
        }

        using (var outputStream = new MemoryStream())
        {
            mergedDocument.Save(outputStream, false);
            return outputStream.ToArray();
        }
    }
}