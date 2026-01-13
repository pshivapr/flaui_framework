using FlaUI.Core;
using FlaUI.UIA3;
using NUnit.Framework;

namespace flaui_framework
{
    [TestFixture]
    public class NotepadTest
    {
        private Application application;
        private UIA3Automation uIA3Automation;

        [SetUp]
        public void Setup()
        {
            application = Application.Launch("notepad.exe");
            Thread.Sleep(500);
            application = Application.Attach("notepad.exe");
            uIA3Automation = new UIA3Automation();
        }

        [Test]
        [Description("Verify the main window title is not null")]
        public void LaunchNotepadAndVerifyTitle()
        {
            var mainWindow = application.GetMainWindow(uIA3Automation);
            Assert.That(mainWindow, Is.Not.Null);
            Assert.That(mainWindow.Title, Does.Contain("Untitled"));
            Console.WriteLine(mainWindow.Title);
        }

        [TearDown]
        public void TearDown()
        {
            uIA3Automation?.Dispose();
            application?.Close();
        }
    }
}
