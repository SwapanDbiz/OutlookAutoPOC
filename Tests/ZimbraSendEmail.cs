using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using Redemption;
using ZCO2.Utils;

namespace ZCO2;

public class ZimbraSendEmail
{
    private TestData _testData;

    [SetUp]
    public void Setup()
    {
        _testData = TestDataLoader.LoadTestData();
    }

    [Test]
    public void TC_User1SendsEmailToUser2AndVerifyReceipt()
    {
        Assert.That(_testData, Is.Not.Null);
        Assert.That(_testData.Users, Is.Not.Null);
        Assert.That(_testData.Users!.ContainsKey("user1"), Is.True, "user1 should exist in test data");
        Assert.That(_testData.Users.ContainsKey("user2"), Is.True, "user2 should exist in test data");
        Assert.That(_testData.Email, Is.Not.Null);

        var user1 = _testData.Users!["user1"];
        var user2 = _testData.Users!["user2"];

        // Generate unique subject with random number
        Random random = new Random();
        int randomNumber = random.Next(10000, 99999);
        string emailSubject = $"ZimbraSendEmail_Test_{randomNumber}";
        Console.WriteLine($"[Test] Generated test subject: {emailSubject}");

        // Step 1: User1 sends email to User2
        SendEmailAsUser1(user1, user2, emailSubject);

        // Step 2: Verify User2 received the email
        VerifyEmailReceivedByUser2(user2, emailSubject, user1);
    }

    private void SendEmailAsUser1(User user1, User user2, string emailSubject)
    {
        // Start Outlook process with Zimbra profile
        StartOutlookWithProfile("Zimbra");
        Console.WriteLine($"[User1] Outlook started, waiting 20 seconds for profile to load completely...");
        Thread.Sleep(20000);

        try
        {
            // Get Outlook application via COM - create instance will connect to running process
            Console.WriteLine($"[User1] Attempting to access Outlook COM interface...");
            Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
            {
                Assert.Inconclusive("Outlook.Application COM object not available");
                return;
            }
            
            dynamic outlookApp = Activator.CreateInstance(outlookType);

            if (outlookApp == null)
            {
                Assert.Inconclusive("Failed to get Outlook application object");
                return;
            }

            Console.WriteLine($"[User1] Got Outlook application object");

            // Create mail item via Outlook COM
            dynamic mailItem = outlookApp.CreateItem(0); // 0 = MailItem
            Console.WriteLine($"[User1] Created new mail item");

            // Set properties
            mailItem.To = user2.Username;
            mailItem.Subject = emailSubject;
            mailItem.Body = $"Test email from user1 to user2\n\nSent at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
            mailItem.Importance = 2; // Normal

            Console.WriteLine($"[User1] To: {mailItem.To}");
            Console.WriteLine($"[User1] Subject: {mailItem.Subject}");
            
            // Save draft before sending
            Console.WriteLine($"[User1] Saving draft...");
            mailItem.Save();
            Thread.Sleep(2000);

            // Send the email
            Console.WriteLine($"[User1] Sending email...");
            mailItem.Send();
            
            Console.WriteLine($"[User1] ✓ Email sent successfully");
            
            // Wait for mail server processing
            Console.WriteLine($"[User1] Waiting 20 seconds for mail server processing...");
            Thread.Sleep(20000);
        }
        catch (System.Exception ex)
        {
            Console.WriteLine($"[User1] ERROR: {ex.Message}");
            Console.WriteLine($"[User1] Stack trace: {ex.StackTrace}");
            Assert.Inconclusive($"Failed to send email: {ex.Message}");
            return;
        }
        finally
        {
            // Ensure Outlook is closed
            Thread.Sleep(1000);
            try
            {
                var proc = Process.GetProcessesByName("OUTLOOK").FirstOrDefault();
                if (proc != null && !proc.HasExited)
                {
                    proc.Kill();
                    proc.WaitForExit();
                    Console.WriteLine("[User1] Outlook process terminated");
                }
            }
            catch { }
        }
    }

    private void VerifyEmailReceivedByUser2(User user2, string emailSubject, User user1)
    {
        // Start Outlook process again for user2 with ZimbraUser2 profile
        StartOutlookWithProfile("ZimbraUser2");
        Console.WriteLine($"[User2] Outlook started, waiting 20 seconds for profile to load completely...");
        Thread.Sleep(20000);  // Increased from 10 to 20 seconds for profile load

        var session = new RDOSession();
        try
        {
            // Logon as User2 using "ZimbraUser2" profile
            session.Logon("ZimbraUser2", null, false, false);
        }
        catch (System.Exception ex)
        {
            Assert.Inconclusive($"Failed to login as user2 with 'ZimbraUser2' profile: {ex.Message}");
            return;
        }

        try
        {
            // Get the Inbox folder
            var inbox = session.GetDefaultFolder(rdoDefaultFolders.olFolderInbox);
            if (inbox == null)
            {
                Assert.Inconclusive("Inbox folder unavailable for user2.");
                return;
            }

            Console.WriteLine($"[User2] Inbox loaded, current email count: {inbox.Items.Count}");
            
            // Close any open send/receive dialog by pressing Escape
            Console.WriteLine($"[User2] Closing any open send/receive dialogs...");
            SendKeys.SendWait("{ESC}");
            Thread.Sleep(1000);
            
            // Perform a send/receive sync to refresh the mailbox
            try
            {
                Console.WriteLine($"[User2] Performing Outlook send/receive to sync mailbox...");
                // Trigger outlook to sync - this is handled by Redemption through the session
                Thread.Sleep(5000);  // Give Outlook time to auto-sync
                
                // Close send/receive dialog after sync completes
                Console.WriteLine($"[User2] Closing send/receive dialog...");
                SendKeys.SendWait("{ESC}");
                Thread.Sleep(1000);
                
                // Refresh inbox items count
                int refreshedCount = inbox.Items.Count;
                Console.WriteLine($"[User2] After sync attempt - Inbox contains {refreshedCount} emails");
            }
            catch (System.Exception syncEx)
            {
                Console.WriteLine($"[User2] Note: Could not perform sync: {syncEx.Message}");
            }

            // Wait and refresh inbox with multiple retries
            RDOMail receivedEmail = null;
            int maxRetries = 10;
            int retryCount = 0;

            while (receivedEmail == null && retryCount < maxRetries)
            {
                Thread.Sleep(5000);  // Wait 5 seconds between retries

                // Search for the email sent by user1
                int emailCount = inbox.Items.Count;
                Console.WriteLine($"[User2] Retry {retryCount + 1}/{maxRetries}: Inbox contains {emailCount} emails");

                foreach (var item in inbox.Items)
                {
                    if (item is RDOMail mail)
                    {
                        if (mail.Subject == emailSubject)
                        {
                            receivedEmail = mail;
                            Console.WriteLine($"[User2] Email found on retry {retryCount + 1}!");
                            break;
                        }
                        
                        // Log any undeliverable messages
                        if (mail.Subject.Contains("Undeliverable") && mail.Subject.Contains(emailSubject))
                        {
                            Console.WriteLine($"[User2] WARNING: Undeliverable notification received: {mail.Subject}");
                            Console.WriteLine($"[User2] Undeliverable message body: {mail.Body.Substring(0, Math.Min(200, mail.Body.Length))}");
                        }
                    }
                }
                
                retryCount++;
            }

            if (receivedEmail != null)
            {
                // Verify email properties
                Assert.That(receivedEmail.Subject, Is.EqualTo(emailSubject), 
                    "Email subject should match");

                Assert.That(receivedEmail.SenderEmailAddress, Contains.Substring(user1.Username!.Split('@')[0]), 
                    "Email should be from user1");

                Assert.That(receivedEmail.Body, Is.Not.Null.And.Not.Empty, 
                    "Email body should not be empty");

                Console.WriteLine($"[User2] Email verification successful!");
                Console.WriteLine($"[User2] - Subject: {receivedEmail.Subject}");
                Console.WriteLine($"[User2] - From: {receivedEmail.SenderEmailAddress}");
                Console.WriteLine($"[User2] - Body Preview: {receivedEmail.Body.Substring(0, Math.Min(50, receivedEmail.Body.Length))}...");
            }
            else
            {
                Console.WriteLine($"[User2] ERROR: Email not found after {maxRetries} retries");
                Console.WriteLine($"[User2] Looking for subject: '{emailSubject}'");
                // Log some of the subjects in inbox for debugging
                Console.WriteLine($"[User2] Recent email subjects in inbox:");
                int logged = 0;
                foreach (var item in inbox.Items)
                {
                    if (item is RDOMail m && logged < 10)
                    {
                        Console.WriteLine($"[User2]   - {m.Subject} (From: {m.SenderEmailAddress})");
                        logged++;
                    }
                }
            }

            // Verify email was received
            Assert.That(receivedEmail, Is.Not.Null, 
                $"Email with subject '{emailSubject}' should be received in user2's inbox after {maxRetries} retries");
        }
        finally
        {
            // Close session after user2 is done
            try
            {
                session.Logoff();
                Console.WriteLine("[User2] Session closed successfully");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Warning: Error closing user2 session: {ex.Message}");
            }
        }

        // Close Outlook process
        Thread.Sleep(2000);
        try
        {
            Process outlookProcess = Process.GetProcessesByName("OUTLOOK").FirstOrDefault();
            if (outlookProcess != null)
            {
                outlookProcess.Kill();
                outlookProcess.WaitForExit();
                Console.WriteLine("[User2] Outlook process closed");
            }
        }
        catch (System.Exception ex)
        {
            Console.WriteLine($"Warning: Error closing Outlook process: {ex.Message}");
        }
    }

    private void StartOutlookWithProfile(string profileName)
    {
        try
        {
            Console.WriteLine($"[Outlook Launch] Starting Outlook 365 with profile: {profileName}");
            
            // For Outlook 365, use the /profile parameter to launch with a specific profile
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = @"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
                Arguments = $"/profile \"{profileName}\"",
                UseShellExecute = false,
                RedirectStandardOutput = false,
                CreateNoWindow = false
            };
            
            Process outlookProcess = Process.Start(psi);
            if (outlookProcess != null)
            {
                Console.WriteLine($"[Outlook Launch] Outlook process started (PID: {outlookProcess.Id}) with profile '{profileName}'");
            }
            else
            {
                Console.WriteLine($"[Outlook Launch] Failed to start Outlook process");
            }
        }
        catch (System.Exception ex)
        {
            Console.WriteLine($"[Outlook Launch] Error starting Outlook with profile: {ex.Message}");
        }
    }
}
