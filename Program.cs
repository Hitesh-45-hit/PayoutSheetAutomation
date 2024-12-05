using System.Diagnostics;
using System.Text;
using Npgsql;
using Microsoft.Extensions.DependencyInjection;
using InvestorSystem.Controllers;
using AutoMapper;
using InvestorSystem.Infrastructure.Areas.Payout.Services;
using InvestorSystem.Infrastructure.DB;
using Microsoft.EntityFrameworkCore;
using MimeKit;
using MailKit.Net.Smtp;
using System.Data;
using OfficeOpenXml;
using ClosedXML.Excel;
using Json.Schema;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            StopIIS();
            UpdateDatabase();
            GetBusinessPartnerDataView();
            GenerateReferenceSheets();
            GeneratePayoutSheets();
            SendmailForPayoutRefSheets("payout").GetAwaiter().GetResult();
            SendmailForPayoutRefSheets("reference").GetAwaiter().GetResult();
            StartIIS();
        }
        catch (Exception ex)
        {
            string message = $"Error stopping IIS site: {ex.Message}";
            Console.WriteLine(message);
            LogMessage(message);
        }
    }
    private static void UpdateDatabase()
    {
        string connectionString = "User ID=postgres;Password=Jazz@2702;Host=localhost;Port=5432;Database=InvestorSystem;Pooling=true;";

        var dbContextOptions = new DbContextOptionsBuilder<AppDBContext>()
            .UseNpgsql(connectionString)
            .Options;

        using (var dbContext = new AppDBContext(dbContextOptions))
        {
            var mapperConfig = new MapperConfiguration(cfg =>
            {
                // Configure AutoMapper mappings here, e.g.:
                // cfg.CreateMap<Source, Destination>();
            });
            IMapper mapper = mapperConfig.CreateMapper();

            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    //string currentDate = DateTime.Now.ToString("yyyy-MM-05");
                    string currentDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 5).AddMonths(-1).ToString("yyyy-MM-dd");
                    //string currentDate = DateTime.Now.ToString("yyyy-11-05");
                    connection.Open();
                    Console.WriteLine("Connected to the database successfully!");

                    // Update queries
                    string updateInvestorPayoutQuery = $@"UPDATE ""Investor_Payout_History"" SET ""PaidOn"" = '{currentDate}' WHERE ""PaidOn"" = '-infinity';";
                    string updateEmployeePayoutQuery = $@"UPDATE ""Employee_Payout_History"" SET ""PaidOn"" = '{currentDate}' WHERE ""PaidOn"" = '-infinity';";

                    using (NpgsqlCommand command = new NpgsqlCommand())
                    {
                        command.Connection = connection;

                        // Execute the first query
                        command.CommandText = updateInvestorPayoutQuery;
                        int investorRowsAffected = command.ExecuteNonQuery();
                        Console.WriteLine($"{investorRowsAffected} rows updated in Investor_Payout_History.");

                        // Execute the second query
                        command.CommandText = updateEmployeePayoutQuery;
                        int employeeRowsAffected = command.ExecuteNonQuery();
                        Console.WriteLine($"{employeeRowsAffected} rows updated in Employee_Payout_History.");
                    }

                    // Determine the last day of the current month
                    //DateTime lastDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
                    //DateTime lastDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(-1).AddDays(-1);
                    DateTime lastDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1);

                    //DateTime lastDayOfMonth = new DateTime(2024, 11, DateTime.DaysInMonth(2024, 11));
                    Console.WriteLine($"Last day of the month: {lastDayOfMonth.ToShortDateString()}");
                    //string apiUrl = $"https://localhost:7223/api/Payout/CalculateAndPayMonthlyPayout?transactionDate={lastDayOfMonth:yyyy-MM-dd}";
                    //Console.WriteLine($"{apiUrl}");
                    //HitApi(apiUrl).GetAwaiter().GetResult();
                    //Console.WriteLine($"{apiUrl}");
                    // Call the CalculateAndPayMonthlyPayout method
                    try
                    {
                        InvestorPayoutService investorpayoutService = new InvestorPayoutService(dbContext, mapper);
                        investorpayoutService.CalculateAndPayMonthlyPayout(lastDayOfMonth);
                        Console.WriteLine("Monthly payout calculation and payment for investor processed successfully.");

                        EmployeePayoutService employeePayoutService = new EmployeePayoutService(dbContext);
                        employeePayoutService.CalculateAndPayMonthlyPayout(lastDayOfMonth);
                        Console.WriteLine("Monthly payout calculation and payment for employee processed successfully.");

                        string message = ($"Monthly payout Job run sucessfully");
                        SendMessageToWhatsapp(message);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error at Monthly payout Job run: {ex.Message}");
                        string message = $"Error at Monthly payout Job run: {ex.Message}";
                        SendMessageToWhatsapp(message);
                        LogMessage(message);

                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error updating database: {ex.Message}");
                    string message = $"Error updating database: {ex.Message}";
                    SendMessageToWhatsapp(message);
                    LogMessage(message);
                }
            }
        }
    }
    public static void GenerateReferenceSheets()
    {
        string connectionString = "Host=localhost;Port=5432;Database=InvestorSystem;Username=postgres;Password=Jazz@2702";
        try
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                string query = @"(SELECT
                            ""Person"".""FirstName"",
                            ""Person"".""LastName"",
                            ""Person"".""PANCardNo"",
                            ""BankDetails"".""AccountName"",
                            ""BankDetails"".""AccountNumber"",
                            ""AccountType"".""Name"" as ""AccountType"",
                            ""BankDetails"".""BankName"",
                            ""BankDetails"".""IFSC"",
                            ROUND(""Investor_Payout_History"".""Amount""::numeric, 2) as ""PayoutAmount"",
                            ROUND(""Investor_Payout_Investment"".""Amount""::numeric, 2) as ""Payout_Investment"",
                            ROUND(""Investor_Comp_Investment"".""Amount""::numeric, 2) as ""Comp_Investment"",
                            ROUND(""Investor_Payout_History"".""TDS""::numeric, 2) as ""PayoutTDS"",
                            ""Branch"".""Name"" as ""BranchName"",
                            'Investor' as ""Emp/Inv""
                        FROM
                            ""Investor""
                            INNER JOIN ""Person"" ON ""Investor"".""PersonID"" = ""Person"".""ID""
                            INNER JOIN ""Investor_Payout_History"" ON ""Investor"".""ID"" = ""Investor_Payout_History"".""InvestorID""
                            INNER JOIN ""Investor_Payout_Investment"" ON ""Investor"".""ID"" = ""Investor_Payout_Investment"".""InvestorID""
                            LEFT JOIN ""Employee"" ON ""Employee"".""ID"" = ""Investor"".""ReferredByID""
                            LEFT JOIN ""BankDetails"" ON ""Investor"".""BankDetailsID"" = ""BankDetails"".""ID""
                            LEFT JOIN ""AccountType"" ON ""AccountType"".""ID"" = ""BankDetails"".""AccountTypeID""
                            LEFT JOIN ""Investor_Comp_Investment"" ON ""Investor"".""ID"" = ""Investor_Comp_Investment"".""InvestorID""
                            LEFT JOIN ""InvestorReferralThirdTable"" IRTT ON ""Investor"".""ID"" = IRTT.""InvestorID""
                            LEFT JOIN ""Branch"" ON ""Investor_Payout_History"".""BranchId"" = ""Branch"".""ID""
                        WHERE
                            ""Investor_Payout_History"".""PaidOn"" = '-infinity'
                            AND ""Investor"".""IsActive""
                            AND ""Investor_Payout_History"".""IsRef"" = 'true'
                        ORDER BY
                            ""Employee"".""ID"", IRTT.""ReferredByInvestorID"")

                        UNION ALL

                        (SELECT
                            ""Person"".""FirstName"",
                            ""Person"".""LastName"",
                            ""Person"".""PANCardNo"",
                            ""BankDetails"".""AccountName"",
                            ""BankDetails"".""AccountNumber"",
                            ""AccountType"".""Name"" as ""AccountType"",
                            ""BankDetails"".""BankName"",
                            ""BankDetails"".""IFSC"",
                            ROUND(""EPH"".""Amount""::numeric, 2) as ""PayoutAmount"",
                            ROUND(""Employee_Payout_Investment"".""Amount""::numeric, 2) as ""Payout_Investment"",
                            ROUND(""Employee_Comp_Investment"".""Amount""::numeric, 2) as ""Comp_Investment"",
                            ROUND(""EPH"".""TDS""::numeric, 2) as ""PayoutTDS"",
                            ""Branch"".""Name"" AS ""BranchName"",
                            'Employee' as ""Emp/Inv""
                        FROM
                            ""Employee_Payout_History"" as ""EPH""
                            LEFT JOIN ""Employee"" ON ""EPH"".""EmployeeID"" = ""Employee"".""ID""
                            LEFT JOIN ""Person"" ON ""Employee"".""PersonID"" = ""Person"".""ID""
                            LEFT JOIN ""BankDetails"" ON ""Employee"".""BankDetailsID"" = ""BankDetails"".""ID""
                            LEFT JOIN ""AccountType"" ON ""AccountType"".""ID"" = ""BankDetails"".""AccountTypeID""
                            LEFT JOIN ""Branch"" ON ""EPH"".""BranchID"" = ""Branch"".""ID""
                            LEFT JOIN ""Employee_Payout_Investment"" ON ""Employee"".""ID"" = ""Employee_Payout_Investment"".""EmployeeID""
                            LEFT JOIN ""Employee_Comp_Investment"" ON ""Employee"".""ID"" = ""Employee_Comp_Investment"".""EmployeeID""
                        WHERE
                            ""EPH"".""IsReferral"" = 'true'
                            AND ""Employee"".""IsActive""
                            AND ""EPH"".""PaidOn"" = '-infinity'
                        ORDER BY
                            ""Employee"".""ID"")";
                      
                using (var cmd = new NpgsqlCommand(query, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    var dataTable = new DataTable();
                    try
                    {
                        dataTable.Load(reader);
                    }
                    catch (Exception ex)
                    {
                        string errormessage = $"Error loading data into DataTable  for referencequery: {ex.Message}";
                        Console.WriteLine(errormessage);
                        LogMessage(errormessage);
                        return;
                    }

                    // Add Flag column and remove "Comp_Investment" column
                  
                   
                    // Create output folder
                    //string outputFolder = $@"C:\Reference_Sheets_{DateTime.Now.ToString("MMMM")}";
                    string outputFolder = $@"C:\Reference_Sheets_{DateTime.Now.AddMonths(-1).ToString("MMMM")}";
                    try
                    {
                        Directory.CreateDirectory(outputFolder);
                    }
                    catch (Exception ex)
                    {
                        string errormessage = $"Error creating output folder  for referencequery: {ex.Message}";
                        Console.WriteLine(errormessage);
                        LogMessage(errormessage);
                        return;
                    }
                    // Group by Branch
                    var branches = dataTable.AsEnumerable().Select(row => row["BranchName"].ToString()).Distinct();
                    foreach (var branch in branches)
                    {
                        try
                        {
                            var branchDataRows = dataTable.AsEnumerable().Where(row => row["BranchName"].ToString() == branch);
                            if (!branchDataRows.Any())
                            {
                                string errormessage = $"No data found for branch {branch}, skipping Excel generation.  for referencequery";
                                Console.WriteLine(errormessage);
                                LogMessage(errormessage);
                                continue;
                            }

                            //var branchData = branchDataRows.CopyToDataTable();

                            //// Remove unnecessary columns
                            //string[] columnsToRemove = { "BranchName", "Emp/Inv" };
                            //RemoveColumns(branchData, columnsToRemove);
                            var investorData = branchDataRows.Where(row => row["Emp/Inv"].ToString() == "Investor").CopyToDataTable();
                            var employeeData = branchDataRows.Where(row => row["Emp/Inv"].ToString() == "Employee").CopyToDataTable();

                            // Remove unnecessary columns
                            string[] columnsToRemove = { "BranchName", "Emp/Inv" };
                            RemoveColumns(investorData, columnsToRemove);
                            RemoveColumns(employeeData, columnsToRemove);
                            // Create a new workbook
                            var workbook = new XLWorkbook();

                            // Add Investor sheet
                            var investorSheet = workbook.Worksheets.Add("InvestorReferral");
                            investorSheet.Cell(1, 1).InsertTable(investorData);

                            var employeeSheet = workbook.Worksheets.Add("EmployeeReferral");
                            employeeSheet.Cell(1, 1).InsertTable(employeeData);

                            // Adjust column widths
                            foreach (var column in investorSheet.ColumnsUsed()) column.AdjustToContents();
                            foreach (var column in employeeSheet.ColumnsUsed()) column.AdjustToContents();

                            // Save the workbook
                            string filename = Path.Combine(outputFolder, $"{branch}.xlsx");
                            workbook.SaveAs(filename);
                            Console.WriteLine($"Excel file generated successfully for branch {branch}: {filename}");
                        }
                        catch (Exception ex)
                        {
                            string errormessage = $"Error generating Excel file for branch {branch}: {ex.Message}  for referencequery";
                            Console.WriteLine(errormessage);
                            LogMessage(errormessage);
                        }
                    }

                }

            }

        }
        catch (Exception ex)
        {
            string errormessage = $"A general error occurred in ReferenceSheets: {ex.Message}  for referencequery";
            Console.WriteLine(errormessage);
            LogMessage(errormessage);
        }

    }
    public static void GeneratePayoutSheets()
    {
        string connectionString = "Host=localhost;Port=5432;Database=InvestorSystem;Username=postgres;Password=Jazz@2702";
        try
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                string query = @"(SELECT
                        ""Person"".""FirstName"",
                        ""Person"".""LastName"",
                        ""Person"".""PANCardNo"",
                        ""BankDetails"".""AccountName"",
                        ""BankDetails"".""AccountNumber"",
                        ""AccountType"".""Name"" as ""AccountType"",
                        ""BankDetails"".""BankName"",
                        ""BankDetails"".""IFSC"",
                        ROUND(""Investor_Payout_History"".""Amount""::numeric, 2) as ""PayoutAmount"",
                        ROUND(""Investor_Payout_Investment"".""Amount""::numeric, 2) as ""Payout_Investment"",
                        ROUND(""Investor_Comp_Investment"".""Amount""::numeric, 2) as ""Comp_Investment"",
                        ROUND(""Investor_Payout_History"".""TDS""::numeric, 2) as ""PayoutTDS"",
                        ""Branch"".""Name"" as ""BranchName"",
                        'Investor' as ""Emp/Inv""
                    FROM
                        ""Investor""
                        INNER JOIN ""Person"" ON ""Investor"".""PersonID"" = ""Person"".""ID""
                        INNER JOIN ""Branch"" ON ""Investor"".""BranchId"" = ""Branch"".""ID""
                        INNER JOIN ""Investor_Payout_History"" ON ""Investor"".""ID"" = ""Investor_Payout_History"".""InvestorID""
                        INNER JOIN ""Investor_Payout_Investment"" ON ""Investor"".""ID"" = ""Investor_Payout_Investment"".""InvestorID""
                        LEFT JOIN ""Employee"" ON ""Employee"".""ID"" = ""Investor"".""ReferredByID""
                        LEFT JOIN ""BankDetails"" ON ""Investor"".""BankDetailsID"" = ""BankDetails"".""ID""
                        LEFT JOIN ""AccountType"" ON ""AccountType"".""ID"" = ""BankDetails"".""AccountTypeID""
                        LEFT JOIN ""Investor_Comp_Investment"" ON ""Investor"".""ID"" = ""Investor_Comp_Investment"".""InvestorID""
	                    LEFT JOIN ""InvestorReferralThirdTable"" IRTT ON ""Investor"".""ID""=IRTT.""InvestorID""
                    WHERE
                        ""Investor_Payout_History"".""PaidOn"" = '-infinity'
                        AND ""Investor_Payout_History"".""IsRef"" = 'false'
                        AND ""Investor"".""IsActive""
                    ORDER BY
                        ""Employee"".""ID"",IRTT.""ReferredByInvestorID"")

                    UNION ALL

                    (SELECT
                        ""Person"".""FirstName"",
                        ""Person"".""LastName"",
                        ""Person"".""PANCardNo"",
                        ""BankDetails"".""AccountName"",
                        ""BankDetails"".""AccountNumber"",
                        ""AccountType"".""Name"" as ""AccountType"",
                        ""BankDetails"".""BankName"",
                        ""BankDetails"".""IFSC"",
                        ROUND(""Employee_Payout_History"".""Amount""::numeric, 2) as ""PayoutAmount"",
                        ROUND(""Employee_Payout_Investment"".""Amount""::numeric, 2) as ""Payout_Investment"",
                        ROUND(""Employee_Comp_Investment"".""Amount""::numeric, 2) as ""Comp_Investment"",
                        ROUND(""Employee_Payout_History"".""TDS""::numeric, 2) as ""PayoutTDS"",
                        ""Branch"".""Name"" as ""BranchName"",
                        'Employee' as ""Emp/Inv""
                    FROM
                        ""Employee""
                        INNER JOIN ""Person"" ON ""Employee"".""PersonID"" = ""Person"".""ID""
                        INNER JOIN ""Branch"" ON ""Employee"".""BranchID"" = ""Branch"".""ID""
                        INNER JOIN ""Employee_Payout_History"" ON ""Employee"".""ID"" = ""Employee_Payout_History"".""EmployeeID""
                        INNER JOIN ""Employee_Payout_Investment"" ON ""Employee"".""ID"" = ""Employee_Payout_Investment"".""EmployeeID""
                        LEFT JOIN ""BankDetails"" ON ""Employee"".""BankDetailsID"" = ""BankDetails"".""ID""
                        LEFT JOIN ""AccountType"" ON ""AccountType"".""ID"" = ""BankDetails"".""AccountTypeID""
                        LEFT JOIN ""Employee_Comp_Investment"" ON ""Employee"".""ID"" = ""Employee_Comp_Investment"".""EmployeeID""
                    WHERE
                        ""Employee_Payout_History"".""PaidOn"" = '-infinity'
                         AND ""Employee_Payout_History"".""IsReferral"" = 'false' 
 	                     AND ""Employee"".""IsActive""
                    ORDER BY
                        ""Employee"".""ID"")";

                using (var cmd = new NpgsqlCommand(query, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    var dataTable = new DataTable();
                    try
                    {
                        dataTable.Load(reader);
                    }
                    catch (Exception ex)
                    {
                        string errormessage = $"Error loading data into DataTable: {ex.Message} for payuout query";
                        Console.WriteLine(errormessage);
                        LogMessage(errormessage);
                        return;
                    }

                    // Add Flag column and remove "Comp_Investment" column
                    dataTable.Columns.Add("Flag", typeof(string));
                    try
                    {
                        foreach (DataRow row in dataTable.Rows)
                        {
                            decimal payoutInvestment = row["Payout_Investment"] != DBNull.Value ? Convert.ToDecimal(row["Payout_Investment"]) : 0;
                            decimal compInvestment = row["Comp_Investment"] != DBNull.Value ? Convert.ToDecimal(row["Comp_Investment"]) : 0;

                            if (payoutInvestment + compInvestment > 2500000 &&
                                payoutInvestment > 0 &&
                                compInvestment > 0)
                            {
                                row["Flag"] = "Yes";
                            }
                            else
                            {
                                row["Flag"] = "";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"An error occurred while add flag: {ex.Message} for payuout query");
                        string errormessage = $"An error occurred while add flag: {ex.Message} for payuout query";
                        LogMessage(errormessage);
                        // Optionally log the error or rethrow it
                    }
                    try
                    {
                        dataTable.Columns.Remove("Comp_Investment");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error removing 'Comp_Investment' column: {ex.Message} for payuout query");
                        string errormessage = $"Error removing 'Comp_Investment' column: {ex.Message} for payuout query";
                        LogMessage(errormessage);
                    }

                    // Create output folder
                    //string outputFolder = $@"C:\Payout_Sheets_{DateTime.Now.ToString("MMMM")}";
                    string outputFolder = $@"C:\Payout_Sheets_{DateTime.Now.AddMonths(-1).ToString("MMMM")}";

                    try
                    {
                        Directory.CreateDirectory(outputFolder);
                    }
                    catch (Exception ex)
                    {
                        string errormessage = $"Error creating output folder: {ex.Message} for payuout query";
                        Console.WriteLine(errormessage);
                        LogMessage(errormessage);
                        return;
                    }
                    // Group by Branch
                    // Group by Branch
                    var branches = dataTable.AsEnumerable().Select(row => row["BranchName"].ToString()).Distinct();
                    foreach (var branch in branches)
                    {
                        try
                        {
                            var branchDataRows = dataTable.AsEnumerable().Where(row => row["BranchName"].ToString() == branch);
                            if (!branchDataRows.Any())
                            {
                                Console.WriteLine($"No data found for branch {branch}, skipping Excel generation.");
                                continue;
                            }

                            var branchData = branchDataRows.CopyToDataTable();

                            var investorRows = branchData.AsEnumerable().Where(row => row["Emp/Inv"].ToString() == "Investor");
                            var employeeRows = branchData.AsEnumerable().Where(row => row["Emp/Inv"].ToString() == "Employee");

                            DataTable investorData = null;
                            DataTable employeeData = null;

                            if (investorRows.Any())
                            {
                                investorData = investorRows.CopyToDataTable();
                                string[] columnsToRemove = { "BranchName", "Emp/Inv" };
                                RemoveColumns(investorData, columnsToRemove);
                            }

                            if (employeeRows.Any())
                            {
                                employeeData = employeeRows.CopyToDataTable();
                                string[] columnsToRemove = { "BranchName", "Emp/Inv" };
                                RemoveColumns(employeeData, columnsToRemove);
                            }

                            // Skip generating Excel if both investor and employee data are empty
                            if (investorData == null && employeeData == null)
                            {
                                string errormessage = $"No data for investors or employees in branch {branch}, skipping Excel generation.";
                                Console.WriteLine(errormessage);
                                LogMessage(errormessage );
                                continue;
                            }

                            // Create a new workbook
                            var workbook = new XLWorkbook();

                            // Add Investor sheet
                            if (investorData != null && investorData.Rows.Count > 0)
                            {
                                var investorSheet = workbook.Worksheets.Add("Investor");
                                investorSheet.Cell(1, 1).InsertTable(investorData);
                            }

                            // Add Employee sheet
                            if (employeeData != null && employeeData.Rows.Count > 0)
                            {
                                var employeeSheet = workbook.Worksheets.Add("Employee");
                                employeeSheet.Cell(1, 1).InsertTable(employeeData);
                            }

                            // Adjust column widths
                            foreach (var worksheet in workbook.Worksheets)
                            {
                                foreach (var column in worksheet.ColumnsUsed())
                                {
                                    column.AdjustToContents();
                                }
                            }

                            // Save the workbook
                            string filename = Path.Combine(outputFolder, $"{branch}.xlsx");
                            workbook.SaveAs(filename);
                            Console.WriteLine($"Excel file generated successfully for branch {branch}: {filename}");
                        }
                        catch (Exception ex)
                        {
                            string errormessage = $"Error generating Excel file for branch {branch}: {ex.Message}";
                            Console.WriteLine(errormessage);
                            LogMessage(errormessage);
                        }
                    }

                }
            }
        }
        catch (Exception ex)
        {
            string errormessage = $"A general error occurred in GeneratePayoutSheets: {ex.Message}";
            Console.WriteLine(errormessage);
            LogMessage(errormessage);
        }

    }
    public static void RemoveColumns(DataTable table, string[] columnsToRemove)
    {
        foreach (var columnName in columnsToRemove)
        {
            if (table.Columns.Contains(columnName))
            {
                table.Columns.Remove(columnName);
            }
        }
    }
    private static void GetBusinessPartnerDataView()
    {
        string connectionString = "User ID=postgres;Password=Jazz@2702;Host=localhost;Port=5432;Database=InvestorSystem;Pooling=true;";

        var dbContextOptions = new DbContextOptionsBuilder<AppDBContext>()
            .UseNpgsql(connectionString)
            .Options;

        using (var dbContext = new AppDBContext(dbContextOptions))
        {
            var mapperConfig = new MapperConfiguration(cfg =>
            {
                // Configure AutoMapper mappings here, e.g.:
                // cfg.CreateMap<Source, Destination>();
            });
            IMapper mapper = mapperConfig.CreateMapper();

            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                
                try
                {
                    InvestorPayoutService investorpayoutService = new InvestorPayoutService(dbContext, mapper);
                    bool result = investorpayoutService.GetBusinessPartnerDataView();
                    if (result)
                    {
                        string message = "Business partner payout calculation done successfully.";
                        Console.WriteLine(message);  // You can log the message as well
                        LogMessage(message);
                    }
                    else
                    {
                        string message = "Business partner payout calculation failed.";
                        Console.WriteLine(message);
                        LogMessage(message);
                    }
                }
                catch (Exception ex)
                {
                    string message = $"Error in Business partner calculation: {ex.Message}";
                    Console.WriteLine(message);
                    LogMessage(message);

                }
                
            }
        }
    }
    private void AdjustColumnWidths(ExcelWorksheet sheet)
    {
        for (int col = 1; col <= sheet.Dimension.Columns; col++)
        {
            double maxLength = 0;
            for (int row = 1; row <= sheet.Dimension.Rows; row++)
            {
                var cellValue = sheet.Cells[row, col].Text;
                if (cellValue != null)
                {
                    maxLength = Math.Max(maxLength, cellValue.Length);
                }
            }
            sheet.Column(col).Width = maxLength + 2; // Adjust width
        }
    }
    public static async Task SendMessageToWhatsapp(string message)
    {
        string url = "https://backend.aisensy.com/campaign/t1/api/v2";
        string apiKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IjY2ZDZiYTllYWY1ZjU4MGI2ZWZkOGE1ZSIsIm5hbWUiOiJUV0ogSVQgU29sdXRpb25zIiwiYXBwTmFtZSI6IkFpU2Vuc3kiLCJjbGllbnRJZCI6IjY2ZDZiYTlkYWY1ZjU4MGI2ZWZkOGEzOCIsImFjdGl2ZVBsYW4iOiJOT05FIiwiaWF0IjoxNzI1MzQ4NTEwfQ.H4iOmOeJEB0cjXl5hv_iBhpbacAqxxjb9mS4c9yyyro";
        string campaignName = "Investor_system_job_status";
        string destination = "917972239751";
        string userName = "Hitesh";

        var payload = new
        {
            apiKey = apiKey,
            campaignName = campaignName,
            destination = destination,
            userName = userName,
            templateParams = new[] { message }
        };

        try
        {
            using (HttpClient client = new HttpClient())
            {
                var jsonPayload = Newtonsoft.Json.JsonConvert.SerializeObject(payload);
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(url, content);

                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine("Message sent successfully.");
                }
                else
                {
                    string errormessage = $"Failed to send message. Status Code: {response.StatusCode}";
                    Console.WriteLine(errormessage);
                    LogMessage(errormessage);
                }
            }
        }
        catch (Exception ex)
        {
            string errormessage = "Failed to send message}";
            Console.WriteLine(errormessage);
            LogMessage(errormessage);
        }
    }
    private static async System.Threading.Tasks.Task SendmailForPayoutRefSheets(string mailType)
    {
        try
        {
            // Define the message and attachment folder path based on the mailType parameter
            string message = mailType == "payout" ? "Monthly payout sheets" : "Reference payout sheets"; // Message depending on mailType
            string attachmentFolderPath = mailType == "payout"
                ? $@"C:\Payout_Sheets_{DateTime.Now.AddMonths(-1).ToString("MMMM")}\"
                : $@"C:\Reference_Sheets_{DateTime.Now.AddMonths(-1).ToString("MMMM")}\"; // Folder path depending on mailType

            var successfulBranches = new List<string>();
            string excelFilePath = @"C:\Logs\EmailData\BranchMail.xlsx"; // Path to your Excel file

            // Read branch email data from the Excel file
            var branchEmailMap = ReadBranchEmailsFromExcel(excelFilePath);
            if (branchEmailMap.Count == 0)
            {
                string errormessage = "No branch data found in the Excel file.";
                Console.WriteLine(errormessage);
                LogMessage(errormessage);
                //SendMessageToWhatsapp(errormessage);
                return;
            }

            // Loop through each branch and send the email
            foreach (var branch in branchEmailMap)
            {
                string branchName = branch.Key;
                string branchEmail = branch.Value;

                var emailMessage = new MimeMessage();
                emailMessage.From.Add(new MailboxAddress("Hitesh Shirdhankar", "hiteshshirdhankar2000@gmail.com"));
                emailMessage.To.Add(new MailboxAddress(branchName, branchEmail));
                emailMessage.Cc.Add(new MailboxAddress("Deepak Khochare", "deepak.khochare@tradewithjazz.com"));
                emailMessage.Subject = mailType == "payout"
                    ? $"Monthly Payout Sheets - {branchName}"
                    : $"Reference Payout Sheets - {branchName}"; // Subject based on mailType

                var bodyBuilder = new BodyBuilder { HtmlBody = message };
                string attachmentFilePath = Path.Combine(attachmentFolderPath, $"{branchName}.xlsx");

                // Attach the file if it exists
                if (File.Exists(attachmentFilePath))
                {
                    bodyBuilder.Attachments.Add(attachmentFilePath);
                }
                else
                {
                    Console.WriteLine($"The Excel file for branch '{branchName}' was not found at the specified path.");
                    string errorMessage = $"The Excel file for branch '{branchName}' was not found at the specified path.";
                    LogMessage(errorMessage);
                    SendMessageToWhatsapp(errorMessage);
                    continue;
                }

                emailMessage.Body = bodyBuilder.ToMessageBody();

                // Send the email
                using (var smtpClient = new SmtpClient())
                {
                    await smtpClient.ConnectAsync("smtp.gmail.com", 465, true);
                    await smtpClient.AuthenticateAsync("hiteshshirdhankar2000@gmail.com", "ycfl ypmm gust mbag");
                    await smtpClient.SendAsync(emailMessage);
                    await smtpClient.DisconnectAsync(true);
                }

                successfulBranches.Add(branchName);
                Console.WriteLine($"Email sent successfully to {branchName} ({branchEmail}) with attachment.");
            }

            // Send WhatsApp message for successful or failed email delivery
            if (successfulBranches.Count > 0)
            {
                string successMessage = $"Email sent successfully to the following branches with attachment: {string.Join(", ", successfulBranches)}.";
                Console.WriteLine(successMessage);
                SendMessageToWhatsapp(successMessage).GetAwaiter().GetResult();
            }
            else
            {
                string noEmailsMessage = "No emails were sent successfully. Please check the file paths and branch mappings.";
                Console.WriteLine(noEmailsMessage);
                SendMessageToWhatsapp(noEmailsMessage).GetAwaiter().GetResult();
                LogMessage(noEmailsMessage);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending emails: {ex.Message}");
            string errorMessage = $"Error sending emails: {ex.Message}";
            SendMessageToWhatsapp(errorMessage).GetAwaiter().GetResult();
            LogMessage(errorMessage);
        }
    }
    private static Dictionary<string, string> ReadBranchEmailsFromExcel(string excelFilePath)
    {
        var branchEmailMap = new Dictionary<string, string>();

        if (!File.Exists(excelFilePath))
        {
            string message = "Excel file not found at the specified path.";
            Console.WriteLine(message);
            LogMessage(message);
            return branchEmailMap;
        }

        using (var workbook = new XLWorkbook(excelFilePath))
        {
            var worksheet = workbook.Worksheet(1); // Assuming data is in the first worksheet
            var rows = worksheet.RangeUsed().RowsUsed();
            var headerRow = rows.First();
            var branchNameColumnIndex = headerRow.Cells().First(c => c.GetValue<string>().Trim() == "BranchName").Address.ColumnNumber;
            var branchEmailColumnIndex = headerRow.Cells().First(c => c.GetValue<string>().Trim() == "BranchEmail").Address.ColumnNumber;

            foreach (var row in rows.Skip(1)) // Skip the header row
            {
                //string branchName = row.Cell(2).GetValue<string>().Trim();
                //string branchEmail = row.Cell(4).GetValue<string>().Trim();
                string branchName = row.Cell(branchNameColumnIndex).GetValue<string>().Trim();
                string branchEmail = row.Cell(branchEmailColumnIndex).GetValue<string>().Trim();
                Console.WriteLine(branchEmail);

                if (!string.IsNullOrWhiteSpace(branchName) && !string.IsNullOrWhiteSpace(branchEmail))
                {
                    branchEmailMap[branchName] = branchEmail;
                }
            }
        }

        return branchEmailMap;
    }
    public static async Task LogMessage(string message)
    {
        try
        {
            string logDirectory = @"C:\Logs";

            // Ensure the log directory exists
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            // Create a log file with today's date as the name
            string logFilePath = Path.Combine(logDirectory, $"{DateTime.Now:yyyy-MM-dd}.txt");

            // Write the log message to the file
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                await writer.WriteLineAsync($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
            }
        }
        catch (Exception logEx)
        {
            Console.WriteLine($"Failed to log message: {logEx.Message}");
        }
    }
    private static async Task StartIIS()
    {
        try
        {
            string siteName = "InvestorAPI";
            string command = $"Start-Website -Name \"{siteName}\"";
            // Setup to execute PowerShell with the provided command
            ProcessStartInfo processInfo = new ProcessStartInfo("powershell", command)
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            // Start the PowerShell process
            using (Process process = new Process { StartInfo = processInfo })
            {
                process.Start();  // Start the process

                string result = process.StandardOutput.ReadToEnd();
                Console.WriteLine(result);
                Console.WriteLine($"IIS site {siteName} server started sucessfully");
                string message = ($"IIS site {siteName} server started sucessfully");
                //SendMessageToWhatsapp(message).GetAwaiter().GetResult();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            string message = $"Error while starting IIS site: {ex.Message}";
            SendMessageToWhatsapp(message).GetAwaiter().GetResult();
            LogMessage(message);
        }

    }
    private static async Task StopIIS()
    {
        try
        {
            string siteName = "InvestorAPI";
            string command = $"Stop-Website -Name \"{siteName}\"";
            // Setup to execute PowerShell with the provided command
            ProcessStartInfo processInfo = new ProcessStartInfo("powershell", command)
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            // Start the PowerShell process
            using (Process process = new Process { StartInfo = processInfo })
            {
                process.Start();  // Start the process

                string result = process.StandardOutput.ReadToEnd();
                Console.WriteLine(result);
                Console.WriteLine($"IIS site {siteName} server stop sucessfully");
                string message = ($"IIS site {siteName} server stop sucessfully");
                //SendMessageToWhatsapp(message).GetAwaiter().GetResult();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            string message = $"Error while stopping IIS site: {ex.Message}";
            SendMessageToWhatsapp(message).GetAwaiter().GetResult();
            LogMessage(message);
        }

    }

}
