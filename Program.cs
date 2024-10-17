using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.UIA2;
using System.Text.RegularExpressions;
using Microsoft.Playwright;
using System.Threading.Tasks;
using System.Collections.Generic;

public class AccountInfo
{
    // This class holds the account information.
    public string ClientId { get; set; }
    public string AccountNumber { get; set; }
    public string AccountAmount { get; set; }
    public string Wiid { get; set; } // Work Item ID from the UIPath ACME test website
}

class Program
{
    static async Task Main(string[] args)
    {
        // Launch the desktop application (ACME System3) using FlaUI
        var app = Application.Launch("ApplicationPath/ACME-System3.exe");
        var automation = new UIA2Automation();
        var window = app.GetMainWindow(automation);
        ConditionFactory cf = new ConditionFactory(new UIA2PropertyLibrary());
        window.Focus(); // Bring the window into focus
        
        // Login to the desktop application
        window.FindFirstDescendant(cf.ByAutomationId("textBox1")).AsTextBox().Text = "yourEmail"; // Input email for acme windows app
        window.FindFirstDescendant(cf.ByAutomationId("textBox2")).AsTextBox().Text = "YourPassowrd"; // Input password for acme windows app
        window.FindFirstDescendant(cf.ByAutomationId("button1")).AsButton().Click(); // Click login button
        window.WaitUntilEnabled(TimeSpan.FromSeconds(60)); // Wait until login process is complete
        
        // Get the new dashboard window after login
        var dashboard = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));

        // Accessing the menu in the desktop application
        var menu = dashboard.FindFirstDescendant(cf.Menu()).AsMenu();
        menu.Items[2].Invoke().WaitUntilEnabled(TimeSpan.FromSeconds(60)); // Navigate through the menu
        menu.Items[2].Items[1].Invoke().WaitUntilEnabled(TimeSpan.FromSeconds(60));
        menu.Items[2].Items[1].Items[1].Invoke();

        // Start browser automation with Playwright to interact with the ACME test website
        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = true });
        var context = await browser.NewContextAsync();
        var page = await context.NewPageAsync();

        try
        {
            // Navigate to the ACME test website and log in
            await page.GotoAsync("https://acme-test.uipath.com/login");
            await page.GetByLabel("Email:").FillAsync("yourusername"); // Input email for acme web portal
            await page.GetByLabel("Password:").FillAsync("yourpassword"); // Input password for acme web portal
            await page.GetByRole(AriaRole.Button, new() { Name = "Login" }).ClickAsync(); // Click login button
            await page.GetByRole(AriaRole.Button, new() { Name = "Work Items" }).ClickAsync(); // Navigate to "Work Items"

            // Handle pagination on the Work Items page
            var pageCount = await page.Locator("ul.page-numbers li a.page-numbers").CountAsync();
            var search = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));
            System.Console.WriteLine(pageCount); // Log total number of pages

            // Iterate through each page of work items
            for (int i = 1; i <= pageCount - 1; i++)
            {
                // Extract account information for each page
                var accountInfos = await HandlePageAsync(page, i);
                
                // Process each account info
                foreach (var accountInfo in accountInfos)
                {
                    // Search for client in the desktop application
                    search.FindFirstDescendant(cf.ByAutomationId("textBox1")).AsTextBox().Text = accountInfo.ClientId;
                    search.FindFirstDescendant(cf.ByAutomationId("checkBox1")).AsCheckBox().Click();
                    search.FindFirstDescendant(cf.ByAutomationId("button1")).AsButton().Click();
                    search.FindFirstDescendant(cf.ByName(accountInfo.ClientId)).AsButton().DoubleClick(); // Open client details

                    // Process account movements
                    var clientInfo = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));
                    clientInfo.FindFirstDescendant(cf.ByAutomationId("button1")).AsButton().Click(); // Fetch account details
                    var clientAccountInfo = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));

                    // Get the account list in the desktop application
                    var clientAccountList = clientAccountInfo.FindFirstDescendant(cf.ByAutomationId("listView1")).AsGrid().WaitUntilEnabled(TimeSpan.FromSeconds(60));
                    clientAccountInfo.FindFirstDescendant(cf.ByName(accountInfo.AccountNumber.Trim())).AsGridRow().DoubleClick(); // Open specific account

                    // Fetch account movements and calculate total sum
                    var accountMovement = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));
                    accountMovement.FindFirstDescendant(cf.ByName("Show All")).AsButton().Click();
                    var listAccount = accountMovement.FindFirstDescendant(cf.ByAutomationId("listView1")).AsGrid().WaitUntilEnabled(TimeSpan.FromSeconds(60));
                    double sum = listAccount.Rows.Sum(row => double.Parse(row.Cells[2].Name)); // Calculate total amount of transactions

                    // Check if the transaction is successful by comparing amounts
                    string numericPart = Regex.Match(accountInfo.AccountAmount, @"-?\d+").Value;
                    if (double.TryParse(numericPart, out double accountAmount) && accountAmount == sum)
                    {
                        System.Console.WriteLine("Transaction Successful");
                        await UpdatePage(page, true, accountInfo.Wiid, accountInfo.AccountNumber, "Transaction Successful");
                    }
                    else
                    {
                        await UpdatePage(page, false, accountInfo.Wiid, accountInfo.AccountNumber, "Transaction Failed");
                    }

                    // Close open windows
                    clientAccountInfo.FindFirstDescendant(cf.ByAutomationId("ClientAccountMovemnets")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Close();
                    clientInfo.FindFirstDescendant(cf.ByAutomationId("ClientAccounts")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Close();
                    search.FindFirstDescendant(cf.ByAutomationId("ClientDetails")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Close();
                    dashboard.FindFirstDescendant(cf.ByAutomationId("SearchingClient")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60));
                    search.FindFirstDescendant(cf.ByAutomationId("checkBox1")).AsCheckBox().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Click();
                }
            }

            // Complete the process
            System.Console.WriteLine("Process Completed");
            dashboard.FindFirstDescendant(cf.ByAutomationId("SearchingClient")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Close();
            window.Close();
        }
        finally
        {
            // Close browser context and application window
            await context.CloseAsync();
            await browser.CloseAsync();
        }
    }

    // Method to update the status of a work item in the ACME test website
    private static async Task UpdatePage(IPage page, bool isSuccessful, string wiid, string accountNumber, string comment)
    {
        try
        {
            await page.GotoAsync($"https://acme-test.uipath.com/work-items/update/{wiid}");
            await page.Locator("#newComment").FillAsync(comment); // Add a comment for the transaction
            await page.GetByRole(AriaRole.Button, new() { Name = "---" }).ClickAsync();

            // Update the status based on whether the transaction was successful
            string status = isSuccessful ? "Completed" : "Rejected";
            await page.GetByRole(AriaRole.Listbox)
                .GetByRole(AriaRole.Option, new() { Name = status })
                .ClickAsync();

            await page.GetByRole(AriaRole.Button, new() { Name = "Update Work Item" }).ClickAsync();
            Console.WriteLine($"Updated work item {wiid} for account {accountNumber} with status: {status}");
        }
        catch (Exception e)
        {
            Console.Error.WriteLine($"Error updating work item for WIID {wiid}: {e.Message}");
        }
    }

    // Method to handle each page of work items and extract account info
    private static async Task<List<AccountInfo>> HandlePageAsync(IPage page, int pageNumber)
    {
        var accountInfoList = new List<AccountInfo>();

        try
        {
            // Navigate to the specific page number
            await page.GotoAsync($"https://acme-test.uipath.com/work-items?page={pageNumber}");

            // Wait for the work item table to be visible
            await page.WaitForSelectorAsync("table tbody tr");

            // Extract work item details using JavaScript and return as JSON string
            var jsonResult = await page.EvaluateAsync<string>(@"
            () => {
                const rows = document.querySelectorAll('table tbody tr');
                const data = Array.from(rows).map(row => {
                    const cells = row.cells;
                    return {
                        wiid: cells[1]?.textContent?.trim() || '',
                        link: row.querySelector('a')?.href || '',
                        type: cells[3]?.textContent?.trim() || '',
                        clientId: cells[4]?.textContent?.trim() || '',
                        accountNumber: cells[5]?.textContent?.trim() || '',
                        accountAmount: cells[6]?.textContent?.trim() || ''
                    };
                });
                return JSON.stringify(data);
            }");

            // Deserialize the JSON result into a list of AccountInfo objects
            accountInfoList = System.Text.Json.JsonSerializer.Deserialize<List<AccountInfo>>(jsonResult);
        }
        catch (Exception e)
        {
            Console.Error.WriteLine($"Error handling page {pageNumber}: {e.Message}");
        }

        return accountInfoList;
    }
}
