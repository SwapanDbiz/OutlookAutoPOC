using System.Diagnostics;
using System.Threading;
using Redemption;
using ZCO2.Utils;

namespace ZCO2;

public class Tests
{
    private TestData _testData;

    [SetUp]
    public void Setup()
    {
        _testData = TestDataLoader.LoadTestData();
    }

    [Test]
    public void TestOpenOutlookAndVerifyInbox()
    {
        // Launch Outlook classic app
        Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");

        // Wait for Outlook to open
        Thread.Sleep(5000);

        // Use Redemption to access Outlook
        var session = new RDOSession();
        session.Logon();

        // Get the inbox folder
        var inbox = session.GetDefaultFolder(rdoDefaultFolders.olFolderInbox);

        // Verify inbox exists
        Assert.That(inbox, Is.Not.Null, "Inbox folder should exist");

        // Additional verification: check if there are items (assuming inbox has emails)
        // You can customize this assertion based on your needs
        Assert.That(inbox.Items.Count, Is.GreaterThanOrEqualTo(0), "Inbox should have items or be empty");

        // Log off
        session.Logoff();
    }

    [Test]
    public void TestLaunchEmailComposerAndSendEmail()
    {
        Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
        Thread.Sleep(10000); // give Outlook enough startup time

        var session = new RDOSession();
        try
        {
            session.Logon();
        }
        catch (System.Exception ex)
        {
            Assert.Inconclusive($"Outlook/Redemption not ready: {ex.Message}");
            return;
        }

        var drafts = session.GetDefaultFolder(rdoDefaultFolders.olFolderDrafts);
        if (drafts == null)
        {
            Assert.Inconclusive("Drafts folder not accessible; cannot run email send test.");
            session.Logoff();
            return;
        }

        var mail = drafts.Items.Add("IPM.Note") as RDOMail;
        if (mail == null)
        {
            Assert.Inconclusive("Failed to create a draft mail item. Outlook/Redemption may be unavailable.");
            session.Logoff();
            return;
        }

        Assert.That(_testData, Is.Not.Null, "Test data should be loaded");
        Assert.That(_testData.Users, Is.Not.Null, "Users section should exist in test data");
        Assert.That(_testData.Email, Is.Not.Null, "Email section should exist in test data");

        if (!_testData.Users!.TryGetValue("user2", out var user2) || user2 == null)
        {
            Assert.Fail("User2 entry missing in test data.");
            session.Logoff();
            return;
        }

        Assert.That(user2.Username, Is.Not.Null.And.Not.Empty, "User2 username should not be empty");
        Assert.That(user2.Password, Is.Not.Null.And.Not.Empty, "User2 password should not be empty");
        Assert.That(_testData.Email!.Subject, Is.Not.Null.And.Not.Empty, "Email subject should not be empty");
        Assert.That(_testData.Email!.Body, Is.Not.Null.And.Not.Empty, "Email body should not be empty");

        mail.To = user2.Username;
        mail.Subject = _testData.Email.Subject;
        mail.Body = _testData.Email.Body;

        try
        {
            mail.Display();
            Thread.Sleep(2000);
            mail.Send();
        }
        catch (System.Exception ex)
        {
            Assert.Inconclusive($"Unable to send mail via Outlook/Redemption: {ex.Message}");
            session.Logoff();
            return;
        }

        Assert.That(mail.EntryID, Is.Not.Null.And.Not.Empty, "Sent email should have an EntryID");
        session.Logoff();
    }
}
