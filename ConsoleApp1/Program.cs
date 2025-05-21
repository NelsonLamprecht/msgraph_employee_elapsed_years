using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

internal class Program
{
    private static async Task Main(string[] args)
    {
        // Configuration - Replace with your values
        const string tenantId = "a guid"; // Your Azure AD Tenant ID
        const string clientId = "a guid "; // Your Azure AD App Registration Client ID
        const string clientSecret = ""; // Your Azure AD App Registration Client Secret
        const string userPrincipalName = "first.last@company.com"; // UPN of the employee whose manager's reports you want to find

        // The scopes depend on the permissions granted to the application in Azure AD.
        // For reading user properties, User.Read.All is common for application permissions.
        var scopes = new[] { "https://graph.microsoft.com/.default" };

        try
        {
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            var graphClient = new GraphServiceClient(credential, scopes);

            Console.WriteLine($"Finding all direct and indirect reports for manager: {userPrincipalName} ...");

            // Get the manager user object
            var managerUser = await graphClient.Users[userPrincipalName].GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "userPrincipalName", "extension_d3f82bf927c2420bb40346c012ce44b7_hireDate" };
            });
            if (managerUser == null)
            {
                Console.WriteLine($"Manager with UPN {userPrincipalName} not found.");
                return;
            }
            var managerId = managerUser.Id;

            // BFS to find all direct and indirect reports
            var allReports = new List<User>();
            var queue = new Queue<string>();
            queue.Enqueue(managerId);
            var seen = new HashSet<string> { managerId };

            while (queue.Count > 0)
            {
                var currentManagerId = queue.Dequeue();
                var directReportsPage = await graphClient.Users[currentManagerId].DirectReports.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "userPrincipalName", "extension_d3f82bf927c2420bb40346c012ce44b7_hireDate" };
                    requestConfiguration.QueryParameters.Top = 999;
                });
                while (directReportsPage != null)
                {
                    var directReports = directReportsPage.Value?.OfType<User>()?.ToList() ?? new List<User>();
                    foreach (var user in directReports)
                    {
                        if (user.Id != null && !seen.Contains(user.Id))
                        {
                            allReports.Add(user);
                            queue.Enqueue(user.Id);
                            seen.Add(user.Id);
                        }
                    }
                    if (directReportsPage.OdataNextLink == null)
                        break;
                    directReportsPage = await graphClient.Users[currentManagerId].DirectReports.GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "userPrincipalName", "extension_d3f82bf927c2420bb40346c012ce44b7_hireDate" };
                        requestConfiguration.QueryParameters.Top = 999;
                        // Use the next link to get the next page of results
                        // Note: The ODataNextLink is not directly usable in the requestConfiguration,
                        //requestConfiguration.QueryParameters.SkipToken = directReportsPage.OdataNextLink;
                    });
                }
            }

            // Include the manager user in the calculation
            allReports.Insert(0, managerUser);

            Console.WriteLine($"Found {allReports.Count} direct and indirect reports (including manager) under {userPrincipalName}.\n");

            int totalWorkingDays = 0;
            double totalElapsedYears = 0.0;

            foreach (var user in allReports)
            {
                if (user.AdditionalData != null && user.AdditionalData.TryGetValue("extension_d3f82bf927c2420bb40346c012ce44b7_hireDate", out object hireDateValue) && hireDateValue != null)
                {
                    DateTime hireDate;
                    if (DateTime.TryParse(hireDateValue.ToString(), out hireDate))
                    {
                        hireDate = DateTime.SpecifyKind(hireDate, DateTimeKind.Local).ToLocalTime();
                        int workingDays = CalculateWorkingDays(hireDate, DateTime.Today);
                        TimeSpan totalTimeSinceHire = DateTime.Today - hireDate.Date;
                        double elapsedYears = totalTimeSinceHire.TotalDays / 365.25;
                        totalWorkingDays += workingDays;
                        totalElapsedYears += elapsedYears;
                        Console.WriteLine($"Employee: {user.DisplayName} ({user.UserPrincipalName})");
                        Console.WriteLine($"  Hire Date: {hireDate:yyyy-MM-dd}");
                        Console.WriteLine($"  Working days since hire: {workingDays}");
                        Console.WriteLine($"  Elapsed years since hire: {elapsedYears:F2}\n");
                    }
                    else
                    {
                        Console.WriteLine($"Employee: {user.DisplayName} ({user.UserPrincipalName}) - Invalid hire date format");
                    }
                }
                else
                {
                    Console.WriteLine($"Employee: {user.DisplayName} ({user.UserPrincipalName}) - No hire date found");
                }
            }
            Console.WriteLine($"Total number of users in allReports: {allReports.Count}");
            Console.WriteLine($"Sum of working days for all users: {totalWorkingDays}");
            Console.WriteLine($"Sum of elapsed years for all users: {totalElapsedYears:F2}");
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: {ex.Message}");
            Console.ResetColor();
            Console.WriteLine("\\\\nEnsure you have configured your Azure AD App Registration correctly:");
            Console.WriteLine("1. Granted appropriate Application Permissions (e.g., User.Read.All).");
            Console.WriteLine("2. Granted Admin Consent for these permissions.");
            Console.WriteLine("3. Correctly set the Tenant ID, Client ID, Client Secret, and User Principal Name.");
        }
    }

    private static int CalculateWorkingDays(DateTime startDate, DateTime endDate)
    {
        int workingDays = 0;
        DateTime currentDate = startDate.Date;

        if (endDate < startDate)
        {
            return 0; // Or throw an exception, depending on desired behavior
        }

        while (currentDate <= endDate.Date)
        {
            if (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday)
            {
                workingDays++;
            }
            currentDate = currentDate.AddDays(1);
        }
        return workingDays;
    }
}