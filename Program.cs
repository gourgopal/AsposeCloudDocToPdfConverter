using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words.Cloud.Sdk;
using Aspose.Words.Cloud.Sdk.Model.Requests;

namespace AsposeDocxToPdfConverter
{
    class Program
    {
        // Replace with your credentials from https://dashboard.aspose.cloud/
        private const string ClientId = "YOUR_CLIENT_ID";
        private const string ClientSecret = "YOUR_CLIENT_SECRET";

        static async Task Main(string[] args)
        {
            Console.WriteLine("Aspose Cloud DOCX to PDF Converter");
            Console.WriteLine("-----------------------------------");

            try
            {
                // Get file paths
                Console.Write("Enter source DOCX path: ");
                string docxPath = Console.ReadLine();

                Console.Write("Enter target PDF path: ");
                string pdfPath = Console.ReadLine();

                if (!ValidatePaths(docxPath, pdfPath)) return;

                var stopwatch = Stopwatch.StartNew();

                // Initialize API
                var config = new Aspose.Words.Cloud.Sdk.Configuration
                {
                    ClientId = ClientId,
                    ClientSecret = ClientSecret
                };
                var wordsApi = new WordsApi(config);

                // Perform conversion
                await ConvertDocxToPdf(wordsApi, docxPath, pdfPath);

                stopwatch.Stop();

                Console.WriteLine($"\nConversion successful in {stopwatch.Elapsed.TotalSeconds:0.00}s");
                Console.WriteLine($"Output file: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
                Console.WriteLine("Check credentials and internet connection");
            }
        }

        static bool ValidatePaths(string docxPath, string pdfPath)
        {
            if (!File.Exists(docxPath))
            {
                Console.WriteLine("Error: Source file not found");
                return false;
            }

            if (Path.GetExtension(docxPath).ToLower() != ".docx")
            {
                Console.WriteLine("Error: Only .docx files supported");
                return false;
            }

            try
            {
                var dir = Path.GetDirectoryName(pdfPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                return true;
            }
            catch
            {
                Console.WriteLine("Error: Invalid output path");
                return false;
            }
        }

        static async Task ConvertDocxToPdf(WordsApi wordsApi, string docxPath, string pdfPath)
        {
            using (var inputStream = File.OpenRead(docxPath))
            {
                // Create conversion request WITHOUT outPath
                var convertRequest = new ConvertDocumentRequest(
                    document: inputStream,
                    format: "pdf"
                );

                // Get result stream
                var resultStream = await wordsApi.ConvertDocument(convertRequest);

                // Save locally
                using (var outputStream = File.Create(pdfPath))
                {
                    await resultStream.CopyToAsync(outputStream);
                }
            }
        }
    }
}
