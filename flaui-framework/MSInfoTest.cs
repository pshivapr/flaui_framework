using FlaUI.Core;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using NUnit.Framework;

namespace flaui_framework
{
    [TestFixture]
    public class MSInfoTest
    {
        private UIA3Automation automation;
        private Application? msInfoApp;
        private readonly string valueXpath = "/Text[2]";

        [SetUp]
        public void Setup()
        {
            automation = new UIA3Automation();
        }

        [Test]
        [Description("Launch MSInfo32 via Windows Start menu and Run dialog")]
        public void LaunchMSInfoViaStartMenuAndRun()
        {
            try
            {
                // Step 1: Click Windows Start button using Win key
                Keyboard.Press(VirtualKeyShort.LWIN);
                Thread.Sleep(500);
                Keyboard.Release(VirtualKeyShort.LWIN);
                Thread.Sleep(500);

                // Step 2: Search for "RUN"
                Keyboard.Type("run");
                Thread.Sleep(1000);

                // Step 3: Press Enter to open Run dialog
                Keyboard.Press(VirtualKeyShort.ENTER);
                Thread.Sleep(1000);

                // Step 4: Get the Run dialog window
                var desktop = automation.GetDesktop();
                var runDialog = desktop.FindFirstChild(cf => cf.ByName("Run").And(cf.ByControlType(ControlType.Window)));

                Assert.That(runDialog, Is.Not.Null, "Run dialog should be opened");
                Console.WriteLine($"Run dialog found: {runDialog.Name}");

                // Step 5: Type "msinfo32" in the Run dialog
                Keyboard.Type("msinfo32");
                Thread.Sleep(500);

                // Step 6: Press Enter to launch MSInfo32
                Keyboard.Press(VirtualKeyShort.ENTER);
                Thread.Sleep(2000); // MSInfo32 takes time to load

                // Step 7: Find and verify the System Information window
                var msInfoWindow = desktop.FindFirstChild(cf => cf.ByName("System Information").And(cf.ByControlType(ControlType.Window)));

                Assert.That(msInfoWindow, Is.Not.Null, "System Information window should be opened");
                Assert.That(msInfoWindow.Name, Does.Contain("System Information"));
                Console.WriteLine($"MSInfo32 window title: {msInfoWindow.Name}");

                // Attach to the process for cleanup
                msInfoApp = Application.Attach(msInfoWindow.Properties.ProcessId);

                // Step 8: Assertions on specific properties

                // Verify BIOS Mode is set to UEFI
                var biosMode = msInfoWindow.FindFirstDescendant(cf => cf.ByName("BIOS Mode"));
                var biosModeText = biosMode!.AsLabel();
                var biosModeValue = biosModeText!.FindFirstByXPath(valueXpath).AsLabel();
                Console.WriteLine($"{biosModeText?.Text}: {biosModeValue?.Text}");
                Assert.That(biosModeValue?.Text, Does.Contain("UEFI"));

                // Verify Virtualization-based Security Status is Running
                var vbsStatus = msInfoWindow.FindFirstDescendant(cf => cf.ByName("Virtualisation-based security"));
                var vbsStatusText = vbsStatus!.AsLabel();
                var vbsStatusValue = vbsStatusText!.FindFirstByXPath(valueXpath).AsLabel();
                Console.WriteLine($"{vbsStatusText?.Text}: {vbsStatusValue?.Text}");
                Assert.That(vbsStatusValue?.Text, Does.Contain("Running"));

                // Verify Secure Boot State is On
                var secureBootState = msInfoWindow.FindFirstDescendant(cf => cf.ByName("Secure Boot State"));
                var secureBootStateText = secureBootState!.AsLabel();
                var secureBootStateValue = secureBootStateText!.FindFirstByXPath(valueXpath).AsLabel();
                Console.WriteLine($"{secureBootStateText?.Text}: {secureBootStateValue?.Text}");
                Assert.That(secureBootStateValue?.Text, Does.Contain("On"));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed with error: {ex.Message}");
                throw;
            }
        }

        [TearDown]
        public void TearDown()
        {
            // Close MSInfo32 if it's still open
            if (msInfoApp != null && !msInfoApp.HasExited)
            {
                msInfoApp.Close();
                Thread.Sleep(500);
            }

            automation?.Dispose();
        }
    }
}
