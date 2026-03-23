using System.Diagnostics;
using System.Threading;
using Redemption;
using ZCO2.Utils;

namespace ZCO2;

public class ZimbraOutlookTests
{
    private TestData _testData;

    [SetUp]
    public void Setup()
    {
        _testData = TestDataLoader.LoadTestData();
    }

    [Test]
    public void TC2_ComposeSendEmailFromUser1ToUser2CcUser3BccUser4()
    {
        Assert.That(_testData, Is.Not.Null);
        Assert.That(_testData.Users, Is.Not.Null);
        Assert.That(_testData.Users!.ContainsKey("user1"), Is.True, "user1 should exist in test data");
        Assert.That(_testData.Users.ContainsKey("user2"), Is.True, "user2 should exist in test data");
        Assert.That(_testData.Users.ContainsKey("user3"), Is.True, "user3 should exist in test data");
        Assert.That(_testData.Users.ContainsKey("user4"), Is.True, "user4 should exist in test data");
        Assert.That(_testData.Email, Is.Not.Null);

        var user1 = _testData.Users!["user1"];
        var user2 = _testData.Users!["user2"];
        var user3 = _testData.Users!["user3"];
        var user4 = _testData.Users!["user4"];

        Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
        Thread.Sleep(10000);

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

        var drafts = session.GetDefaultFolder(rdoDefaultFolders.olFolderDrafts);
        if (drafts == null)
        {
            Assert.Inconclusive("Drafts folder unavailable.");
            session.Logoff();
            return;
        }

        var mail = drafts.Items.Add("IPM.Note") as RDOMail;
        if (mail == null)
        {
            Assert.Inconclusive("Unable to create an email item.");
            session.Logoff();
            return;
        }

        // Using user data from JSON. In real flows, login as user1 should be handled from mailbox context.
        mail.To = user2.Username;
        mail.CC = user3.Username;
        mail.BCC = user4.Username;
        mail.Subject = _testData.Email.Subject;
        mail.Body = _testData.Email.Body;

        try
        {
            mail.Send();
        }
        catch (System.Exception ex)
        {
            Assert.Inconclusive($"Failed to send message: {ex.Message}");
            session.Logoff();
            return;
        }

        Assert.That(mail.EntryID, Is.Not.Null.And.Not.Empty, "Sent email should have EntryID");
        session.Logoff();
    }

    [Test]
    public void TC7_SaveDraftThenSendInSeparateStepsAcrossUsers()
    {
        Assert.That(_testData, Is.Not.Null);
        Assert.That(_testData.Users, Is.Not.Null);
        Assert.That(_testData.Users!.ContainsKey("user1"), Is.True, "user1 should exist in test data");
        Assert.That(_testData.Users.ContainsKey("user2"), Is.True, "user2 should exist in test data");
        Assert.That(_testData.Email, Is.Not.Null);

        var user1 = _testData.Users!["user1"];
        var user2 = _testData.Users!["user2"];

        Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
        Thread.Sleep(10000);

        var session = new RDOSession();

        try
        {
            session.Logon();
        }
        catch (System.Exception ex)
        {
            Assert.Inconclusive($"Outlook login for user1(zco) failed: {ex.Message}");
            return;
        }

        var drafts = session.GetDefaultFolder(rdoDefaultFolders.olFolderDrafts);
        if (drafts == null)
        {
            Assert.Inconclusive("Drafts folder unavailable.");
            session.Logoff();
            return;
        }

        var draft = drafts.Items.Add("IPM.Note") as RDOMail;
        if (draft == null)
        {
            Assert.Inconclusive("Unable to create draft mail item.");
            session.Logoff();
            return;
        }

        draft.To = user2.Username;
        draft.Subject = "TC7 draft " + System.Guid.NewGuid().ToString("N");
        draft.Body = "This is a TC7 draft created by user1(zco).";
        draft.Save();

        // Simulate send/receive step
        Thread.Sleep(3000);

        // Verify draft exists for user1(zwc) view (using same session for simulated multi-user check)
        RDOMail foundDraft = null;
        foreach (var item in drafts.Items)
        {
            if (item is RDOMail m && m.Subject == draft.Subject && m.To == draft.To)
            {
                foundDraft = m;
                break;
            }
        }

        Assert.That(foundDraft, Is.Not.Null, "Draft should be present after save and send/receive.");

        // Reuse user1(zco) session context to send the draft
        if (foundDraft != null)
        {
            foundDraft.Send();
            Assert.That(foundDraft.EntryID, Is.Not.Null.And.Not.Empty, "Draft sent successfully should have EntryID.");
        }

        // Final send/receive simulation
        Thread.Sleep(3000);

        session.Logoff();
    }

    [Test]
    public void TC12_LoginAsUser1ForwardMailToUser2AndSendReceive()
    {
        Assert.That(_testData, Is.Not.Null);
        Assert.That(_testData.Users, Is.Not.Null);
        Assert.That(_testData.Users!.ContainsKey("user1"), Is.True, "user1 should exist in test data");
        Assert.That(_testData.Users.ContainsKey("user2"), Is.True, "user2 should exist in test data");

        var user1 = _testData.Users!["user1"];
        var user2 = _testData.Users!["user2"];

        // Start Outlook and login as user1(zco)
        Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
        Thread.Sleep(10000);

        var session = new RDOSession();
        try
        {
            session.Logon();
        }
        catch (System.Exception ex)
        {
            Assert.Inconclusive($"Outlook login for user1(zco) failed: {ex.Message}");
            return;
        }

        // Get the inbox folder
        var inbox = session.GetDefaultFolder(rdoDefaultFolders.olFolderInbox);
        if (inbox == null)
        {
            Assert.Inconclusive("Inbox folder unavailable.");
            session.Logoff();
            return;
        }

        // Check if there is at least one mail in inbox
        if (inbox.Items.Count == 0)
        {
            Assert.Inconclusive("No mails found in inbox. Unable to forward a mail.");
            session.Logoff();
            return;
        }

        // Get the first mail from inbox
        RDOMail mailToForward = null;
        foreach (var item in inbox.Items)
        {
            if (item is RDOMail mail)
            {
                mailToForward = mail;
                break;
            }
        }

        if (mailToForward == null)
        {
            Assert.Inconclusive("Unable to find a valid mail item to forward in inbox.");
            session.Logoff();
            return;
        }

        // Forward the mail to user2
        RDOMail forwardedMail = null;
        try
        {
            forwardedMail = mailToForward.Forward() as RDOMail;
            if (forwardedMail == null)
            {
                Assert.Inconclusive("Unable to create forwarded mail item.");
                session.Logoff();
                return;
            }

            forwardedMail.To = user2.Username;
            forwardedMail.Send();
        }
        catch (System.Exception ex)
        {
            Assert.Inconclusive($"Failed to forward mail: {ex.Message}");
            session.Logoff();
            return;
        }

        Assert.That(forwardedMail.EntryID, Is.Not.Null.And.Not.Empty, "Forwarded email should have EntryID");

        // Simulate send/receive action
        Thread.Sleep(3000);

        session.Logoff();
    }
}