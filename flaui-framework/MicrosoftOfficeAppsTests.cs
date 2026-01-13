using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using NUnit.Framework;

namespace flaui_framework
{
    [TestFixture]
    public class MicrosoftOfficeAppsTests
    {
        private Application? application;
        private UIA3Automation? uIA3Automation;

        [Test]
        [Description("Launch PowerPoint and verify the title")]
        public void PowerPointTest()
        {
            var appName = @"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE";
            application = Application.Launch(appName);
            Thread.Sleep(2000);
            application = Application.Attach(appName);
            uIA3Automation = new UIA3Automation();
            var mainWindow = application.GetMainWindow(uIA3Automation);
            var blankPresentationElement = mainWindow!.FindFirstByXPath("/Pane/Pane/Pane[2]/Group/Group[1]/List/ListItem[1]");
            blankPresentationElement?.Click();
            Assert.That(mainWindow, Is.Not.Null);
            Assert.That(mainWindow.Title, Does.Contain("Presentation1 - PowerPoint"));
            Console.WriteLine($"Title: {mainWindow.Title}");

            // Click File Tab and check info
            var fileTab = mainWindow.FindFirstDescendant(cf => cf.ByAutomationId("FileTabButton")).AsButton();
            fileTab?.Click();
            var infoTab = mainWindow.FindFirstDescendant(cf => cf.ByName("Info")).AsButton();
            infoTab?.Click();
            var showAllPropertiesLink = mainWindow.FindFirstDescendant(cf => cf.ByName("Show All Properties")).AsLabel();
            showAllPropertiesLink?.Click();

            // Find Author property
            var authorProperty = mainWindow.FindFirstDescendant(cf => cf.ByName("Author")).AsLabel();
            var personaName = authorProperty?.Parent.FindFirstDescendant(cf => cf.ByAutomationId("PersonaName")).AsLabel();
            Console.WriteLine($"Author Name: {personaName?.Text}");

            //
            var accountTab = mainWindow.FindFirstDescendant(cf => cf.ByName("Account")).AsButton();
            accountTab?.Click();
            Thread.Sleep(500);

            var aboutOffice = mainWindow.FindFirstDescendant(cf => cf.ByAutomationId("GroupAboutOfficeProducts"));
            var aboutOfficeText = aboutOffice!.AsLabel();
            var versionInfo = aboutOffice!.FindFirstByXPath("/Text[2]").AsLabel();
            Console.WriteLine($"{aboutOfficeText?.Text}: {versionInfo?.Text}");

            // Close PowerPoint
            var closeButton = mainWindow.FindFirstDescendant(cf => cf.ByName("Close")).AsButton();
            closeButton?.Click();
        }

        [Test]
        [Description("Launch Word and verify the title")]
        public void WordTest()
        {
            var appName = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE";
            var application = Application.Launch(appName);
            Thread.Sleep(2000);
            application = Application.Attach(appName);
            uIA3Automation = new UIA3Automation();
            var mainWindow = application.GetMainWindow(uIA3Automation);
            var blankDocumentElement = mainWindow!.FindFirstByXPath("/Pane[1]/Pane/Pane[2]/Group/Group[1]/List/ListItem[1]");
            blankDocumentElement?.Click();
            Assert.That(mainWindow, Is.Not.Null);
            Assert.That(mainWindow.Title, Does.Contain("Document1 - Word"));
            Console.WriteLine(mainWindow.Title);

            // Edit body and close
            var body = mainWindow.FindFirstDescendant(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document)).AsTextBox();
            body!.Enter("This is a desktop automation test for Microsoft Word using FlaUI.");

            var closeButton = mainWindow.FindFirstDescendant(cf => cf.ByName("Close")).AsButton();
            closeButton?.Click();
            Thread.Sleep(500);

            var dontSaveButton = mainWindow.FindFirstDescendant(cf => cf.ByName("Don't Save")).AsButton();
            dontSaveButton?.Click();
            Thread.Sleep(500);
        }

        [TearDown]
        public void TearDown()
        {
            uIA3Automation?.Dispose();
            application?.Close();
        }
    }
}
