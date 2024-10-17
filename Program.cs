using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.UIA2;
using System.Text.RegularExpressions;
using Microsoft.Playwright;

public class AccountInfo
{
    public string ClientId { get; set; }
    public string AccountNumber { get; set; }
    public string AccountAmount { get; set; }

    public string Wiid { get; set; }
}

class Program
{
    static async Task Main(string[] args)
    {
        var app = Application.Launch("C:\\Users\\Admin\\OneDrive - Xalta Technology Services Pvt Ltd\\Desktop\\ACME-System3-v0.1\\ACME-System3.exe");
        var automation = new UIA2Automation();
        var window = app.GetMainWindow(automation);
        ConditionFactory cf = new ConditionFactory(new UIA2PropertyLibrary());
        window.Focus();

        window.FindFirstDescendant(cf.ByAutomationId("textBox1")).AsTextBox().Text = "anand.kale@xalta.tech";
        window.FindFirstDescendant(cf.ByAutomationId("textBox2")).AsTextBox().Text = "AkAnand@2002";
        window.FindFirstDescendant(cf.ByAutomationId("button1")).AsButton().Click();
        window.WaitUntilEnabled(TimeSpan.FromSeconds(60));

        // Getting New Window
        var dashboard = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));

        // Accessing the menu
        var menu = dashboard.FindFirstDescendant(cf.Menu()).AsMenu();
        menu.Items[2].Invoke().WaitUntilEnabled(TimeSpan.FromSeconds(60));
        menu.Items[2].Items[1].Invoke().WaitUntilEnabled(TimeSpan.FromSeconds(60));
        menu.Items[2].Items[1].Items[1].Invoke();

        // Browser Automation
        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = true });
        var context = await browser.NewContextAsync();
        var page = await context.NewPageAsync();

        try
        {

            await page.GotoAsync("https://acme-test.uipath.com/login");

            await page.GetByLabel("Email:").FillAsync("anand.kale@xalta.tech");
            await page.GetByLabel("Password:").FillAsync("AkAnand@2002");
            await page.GetByRole(AriaRole.Button, new() { Name = "Login" }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "Work Items" }).ClickAsync();

            // Handle pagination
            var pageCount = await page.Locator("ul.page-numbers li a.page-numbers").CountAsync();
            var search = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));
            System.Console.WriteLine(pageCount);
            for (int i = 1; i <= pageCount - 1; i++)
            {
                var accountInfos = await HandlePageAsync(page, i);
                //System.Console.WriteLine("Current Page No."+i+""+accountInfos.Count);
                // Getting New Window

                foreach (var accountInfo in accountInfos)
                {
                    search.FindFirstDescendant(cf.ByAutomationId("textBox1")).AsTextBox().Text = accountInfo.ClientId;
                    search.FindFirstDescendant(cf.ByAutomationId("checkBox1")).AsCheckBox().Click();
                    search.FindFirstDescendant(cf.ByAutomationId("button1")).AsButton().Click();
                    search.FindFirstDescendant(cf.ByName(accountInfo.ClientId)).AsButton().DoubleClick();

                    var clientInfo = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));
                    clientInfo.FindFirstDescendant(cf.ByAutomationId("button1")).AsButton().Click();

                    var clientAccountInfo = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));

                    var clientAccountList = clientAccountInfo.FindFirstDescendant(cf.ByAutomationId("listView1")).AsGrid().WaitUntilEnabled(TimeSpan.FromSeconds(60));


                    clientAccountInfo.FindFirstDescendant(cf.ByName(accountInfo.AccountNumber.Trim())).AsGridRow().DoubleClick();



                    var accountMovement = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));
                    accountMovement.FindFirstDescendant(cf.ByName("Show All")).AsButton().Click();

                    var listAccount = accountMovement.FindFirstDescendant(cf.ByAutomationId("listView1")).AsGrid().WaitUntilEnabled(TimeSpan.FromSeconds(60));

                    double sum = listAccount.Rows.Sum(row => double.Parse(row.Cells[2].Name));



                    // Extract numeric part (including negative sign if present)
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



                    clientAccountInfo.FindFirstDescendant(cf.ByAutomationId("ClientAccountMovemnets")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Close();
                    clientInfo.FindFirstDescendant(cf.ByAutomationId("ClientAccounts")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Close();
                    search.FindFirstDescendant(cf.ByAutomationId("ClientDetails")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Close();
                    dashboard.FindFirstDescendant(cf.ByAutomationId("SearchingClient")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60));
                    search.FindFirstDescendant(cf.ByAutomationId("checkBox1")).AsCheckBox().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Click();
                }


            }

            System.Console.WriteLine("Process Completed");
            dashboard.FindFirstDescendant(cf.ByAutomationId("SearchingClient")).AsWindow().WaitUntilEnabled(TimeSpan.FromSeconds(60)).Close();
            window.Close();


        }
        finally
        {
            //btrowser close
            await context.CloseAsync();
            await browser.CloseAsync();

            //window close
        }
    }


    private static async Task UpdatePage(IPage page, bool isSuccessful, string wiid, string accountNumber, string comment)
    {
        try
        {
            await page.GotoAsync($"https://acme-test.uipath.com/work-items/update/{wiid}");
            await page.Locator("#newComment").FillAsync(comment);
            await page.GetByRole(AriaRole.Button, new() { Name = "---" }).ClickAsync();

            // Set the status based on whether the transaction was successful
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

    private static async Task<List<AccountInfo>> HandlePageAsync(IPage page, int pageNumber)
    {
        var accountInfoList = new List<AccountInfo>();

        try
        {
            // Navigate to the specified page
            await page.GotoAsync($"https://acme-test.uipath.com/work-items?page={pageNumber}");

            // Wait for the table to be visible
            await page.WaitForSelectorAsync("table tbody tr");

            // Get the count of rows
            var rowCount = await page.Locator("table tbody tr").CountAsync();


            // Extract data using JavaScript, but return as JSON string
            var jsonResult = await page.EvaluateAsync<string>(@"
        () => {
            const rows = document.querySelectorAll('table tbody tr');
            const data = Array.from(rows).map(row => {
                const cells = row.cells;
                return {
                    wiid: cells[1]?.textContent?.trim() || '',
                    link: row.querySelector('a')?.href || '',
                    type: cells[3]?.textContent?.trim() || '',
                    status: cells[4]?.textContent?.trim() || ''
                };
            }).filter(item => item.type === 'WI1' && item.status === 'Open');
            return JSON.stringify(data);
        }");

            // Parse the JSON string
            var wiidLinks = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, string>>>(jsonResult);

            // Print the extracted data
            foreach (var item in wiidLinks)
            {
                //Console.WriteLine($"WIID: {item["wiid"]}, Link: {item["link"]}, Type: {item["type"]}, Status: {item["status"]}");

                await page.GotoAsync(item["link"]);
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

                // Wait for the necessary elements to be visible
                await page.WaitForSelectorAsync("div.row div.col-lg-5 p b");

                // Extracting the required information
                string clientId = await page.EvaluateAsync<string>("() => { " +
                    "const b = [...document.querySelectorAll('div.row div.col-lg-5 p b')].find(b => b.textContent.includes('Client ID:')); " +
                    "return b ? b.nextSibling.textContent.trim() : null; " +
                    "}");

                string accountNumber = await page.EvaluateAsync<string>("() => { " +
                    "const b = [...document.querySelectorAll('div.row div.col-lg-5 p b')].find(b => b.textContent.includes('Account Number:')); " +
                    "return b ? b.nextSibling.textContent.trim() : null; " +
                    "}");

                string accountAmount = await page.EvaluateAsync<string>("() => { " +
                    "const b = [...document.querySelectorAll('div.row div.col-lg-5 p b')].find(b => b.textContent.includes('Account Amount:')); " +
                    "return b ? b.nextSibling.textContent.trim() : null; " +
                    "}");

                // Create an AccountInfo object and add it to the list
                var accountInfo = new AccountInfo
                {
                    ClientId = clientId,
                    AccountNumber = accountNumber,
                    AccountAmount = accountAmount,
                    Wiid = item["wiid"]

                };
                accountInfoList.Add(accountInfo);

            }


        }
        catch (Exception e)
        {
            Console.WriteLine($"Error handling page {pageNumber}: {e.Message}");
            Console.WriteLine($"Stack trace: {e.StackTrace}");
        }

        return accountInfoList; // Return the list of account information
    }
}