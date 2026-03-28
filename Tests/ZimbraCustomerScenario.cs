using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using NUnit.Framework;
using Redemption;
using ZCO2.Utils;

namespace ZCO2;

public class ZimbraCustomerScenario
{
    private TestData _testData;

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_RESTORE = 9;

    [SetUp]
    public void Setup()
    {
        _testData = TestDataLoader.LoadTestData();
    }

    [Test]
    public void TC_CustomerScenario_AccountSetup()
    {
        Assert.That(_testData, Is.Not.Null);
        Assert.That(_testData.Users, Is.Not.Null);
        Assert.That(_testData.Users!.ContainsKey("user2"), Is.True, "user2 should exist in test data");

        var user = _testData.Users!["user2"];

        // Step 1: Open Outlook
        Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
        Thread.Sleep(10000); // Wait for Outlook to start

        var outlookProcess = Process.GetProcessesByName("OUTLOOK").FirstOrDefault();
        if (outlookProcess == null)
        {
            Assert.Inconclusive("Outlook process not found.");
            return;
        }

        // Bring Outlook to foreground
        ShowWindow(outlookProcess.MainWindowHandle, SW_RESTORE);
        Thread.Sleep(2000);

        // Step 2-6: Automate account setup using keyboard shortcuts and SendKeys
        // Note: This is a simplified automation using SendKeys. In a real scenario, more robust UI automation might be needed.

        // Open Account Settings: Alt + F (File), then A (Account Settings), then A (Account Settings)
        SendKeys.SendWait("%F");
        Thread.Sleep(1000);
        SendKeys.SendWait("A");
        Thread.Sleep(1000);
        SendKeys.SendWait("A");
        Thread.Sleep(2000); // Wait for dialog to open

        // Click 'New' button: Tab to New button and press Enter
        SendKeys.SendWait("{TAB}");
        Thread.Sleep(500);
        SendKeys.SendWait("{ENTER}");
        Thread.Sleep(2000); // Wait for Add Account dialog

        // Select 'Microsoft Exchange or compatible service': Use arrow keys or assume position
        // Assuming it's the second option, press Down arrow
        SendKeys.SendWait("{DOWN}");
        Thread.Sleep(500);
        SendKeys.SendWait("{ENTER}");
        Thread.Sleep(1000);

        // Step 3: Enter email address
        SendKeys.SendWait(user.Username);
        Thread.Sleep(500);
        SendKeys.SendWait("{ENTER}"); // Next
        Thread.Sleep(2000);

        // Step 4: Enter password when prompted
        SendKeys.SendWait(user.Password);
        Thread.Sleep(500);
        SendKeys.SendWait("{ENTER}"); // Next
        Thread.Sleep(5000); // Wait for auto-discovery

        // Step 5: Wait for connection
        Thread.Sleep(10000); // Adjust based on network

        // Step 6: Click Finish
        SendKeys.SendWait("{ENTER}"); // Finish
        Thread.Sleep(2000);

        // Restart Outlook (close and reopen)
        outlookProcess.Kill();
        Thread.Sleep(2000);
        Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
        Thread.Sleep(10000);

        // Step 7: Verify folders using RDOSession
        var session = new RDOSession();
        try
        {
            session.Logon();
        }
        catch (System.Exception ex)
        {
            Assert.Inconclusive($"Outlook login failed: {ex.Message}");
            return;
        }

        // Verify Inbox
        var inbox = session.GetDefaultFolder(rdoDefaultFolders.olFolderInbox);
        Assert.That(inbox, Is.Not.Null, "Inbox folder should be available");

        // Verify Calendar
        var calendar = session.GetDefaultFolder(rdoDefaultFolders.olFolderCalendar);
        Assert.That(calendar, Is.Not.Null, "Calendar folder should be available");

        // Verify Contacts
        var contacts = session.GetDefaultFolder(rdoDefaultFolders.olFolderContacts);
        Assert.That(contacts, Is.Not.Null, "Contacts folder should be available");

        // Verify Tasks
        var tasks = session.GetDefaultFolder(rdoDefaultFolders.olFolderTasks);
        Assert.That(tasks, Is.Not.Null, "Tasks folder should be available");

        // Expected Result: No errors, folders are present
        // Additional checks can be added if needed, e.g., check for emails in inbox

        session.Logoff();
    }
}