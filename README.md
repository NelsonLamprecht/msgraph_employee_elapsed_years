# msgraph_employee_elapsed_years

This C# console application interacts with Microsoft Graph API to retrieve information about a specified manager and all their direct and indirect reports.

Here's a breakdown of its functionality:

Configuration: It starts by defining constants for Azure AD tenant ID, client ID, client secret, and the User Principal Name (UPN) of the manager to query. 
These need to be filled in with actual values.

Authentication: It authenticates to Microsoft Graph using a client secret credential.

Fetches Manager Details: It retrieves the specified manager's user object, selecting properties like ID, display name, UPN, and a custom hire date extension attribute (extension_d3f82bf927c2420bb40346c012ce44b7_hireDate).

Finds All Reports: It uses a Breadth-First Search (BFS) algorithm to find all direct and indirect reports under the manager. It also fetches the same set of properties for each report. 
It handles pagination when retrieving direct reports.

Calculates Tenure: For each user (manager and reports):
It attempts to parse the hire date from the custom extension attribute.
If successful, it calculates the number of working days (excluding weekends) between the hire date and the current date.
It also calculates the total elapsed years since the hire date.
It then prints the employee's display name, UPN, hire date, calculated working days, and elapsed years.

Prints Summary: After processing all users, it prints the total number of users found, the sum of all working days, and the sum of all elapsed years.

Error Handling: It includes a try-catch block to handle exceptions and provides guidance on common Azure AD App Registration configuration issues if an error occurs.

Helper Method: A private static method CalculateWorkingDays is used to determine the number of weekdays between two dates.
In essence, the application provides a report on the tenure of a manager and their entire reporting chain, based on a custom hire date attribute stored in Azure AD.
