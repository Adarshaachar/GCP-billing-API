using Google.Cloud.BigQuery.V2;
using ClosedXML.Excel;
using Google.Apis.Auth.OAuth2;
using Newtonsoft.Json;
try
{
    Console.Write("Please Provide ProjectId: ");
    string projectId = Console.ReadLine();
    Console.Write("Please Provide BillingId: ");
    string billingId = Console.ReadLine();
    Console.Write("Enter the path to your service account key JSON file: ");
    string keyFilePath = Console.ReadLine();
    Console.Write("Enter the full path where you want to save the Excel file: ");
    string path = Console.ReadLine();
    GoogleCredential credential;
    try
    {
        using (var jsonStream = new FileStream(keyFilePath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(jsonStream);
        }

        Console.WriteLine(new string('-', 20));

        BigQueryClient client = BigQueryClient.Create(projectId, credential);
        string query = $"SELECT * FROM `{projectId}.{projectId}.gcp_billing_export_resource_v1_{billingId}`";
        BigQueryResults results = client.ExecuteQuery(query, parameters: null);

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("BigQuery Data");
            var headers = results.Schema.Fields;
          
            for (int i = 0; i < headers.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = headers[i].Name;
            }
            //values row 
            int row = 2;
            foreach (BigQueryRow rowData in results)
            {
                int col = 1;

                foreach (var field in headers)
                {
                    object value = rowData[field.Name];
                    if (value == null)
                    {
                        worksheet.Cell(row, col++).Value = "null";
                    }
                    else if (value is System.Collections.Generic.Dictionary<string, object> dictionaryValue)
                    {
                        worksheet.Cell(row, col++).Value = JsonConvert.SerializeObject(value);

                    }
                    else
                    {
                        worksheet.Cell(row, col++).Value = value.ToString();
                    }
                }
                row++;
            }

            workbook.SaveAs(path);
        }

        Console.WriteLine("Data has been exported successfully.");
    }
    catch (Google.GoogleApiException ex)
    {
        Console.WriteLine($"Google API Error: {ex.Message}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}

