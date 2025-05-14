using System;
using System.Collections.Generic;
using Amazon;
using Amazon.CostExplorer;
using Amazon.CostExplorer.Model;
using OfficeOpenXml;
using System.IO;

namespace AWSMonthlyCostReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize the AWS Cost Explorer client
            var client = new AmazonCostExplorerClient(RegionEndpoint.USEast1);

            // Get the current date and first day of the month
            DateTime now = DateTime.UtcNow;
            DateTime firstDayOfMonth = new DateTime(now.Year, now.Month, 1);

            // Define the request
            var request = new GetCostAndUsageRequest
            {
                TimePeriod = new DateInterval
                {
                    Start = firstDayOfMonth.ToString("yyyy-MM-dd"),
                    End = now.ToString("yyyy-MM-dd")
                },
                Granularity = Granularity.DAILY,
                Metrics = new List<string> { "UnblendedCost" }
            };

            // Fetch the cost and usage data
            var response = client.GetCostAndUsageAsync(request).Result;

            // Create an Excel package
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Cost Report");

                // Add headers
                worksheet.Cells[1, 1].Value = "Date";
                worksheet.Cells[1, 2].Value = "Cost";

                // Add data to the worksheet
                int row = 2;
                foreach (var result in response.ResultsByTime)
                {
                    worksheet.Cells[row, 1].Value = result.TimePeriod.Start;
                    worksheet.Cells[row, 2].Value = result.Total["UnblendedCost"].Amount;
                    row++;
                }

                // Save the Excel package
                var fileInfo = new FileInfo("CostReport.xlsx");
                package.SaveAs(fileInfo);
            }

            Console.WriteLine("Cost report has been generated successfully!");
        }
    }
}
