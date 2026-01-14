using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using NUnit.Framework;

namespace flaui_framework
{
    [TestFixture]
    public class DeviceManagerTests // Requires admin rights to see all devices and also test needs to be run as administrator
    {
        private UIA3Automation automation;
        private Application? deviceManagerApp;

        [SetUp]
        public void Setup()
        {
            automation = new UIA3Automation();
        }

        [Test]
        [Description("Open Device Manager, expand all items and verify no error devices")]
        public void OpenDeviceManagerExpandAllAndCheckForErrors()
        {
            try
            {
                // Step 1: Open Run dialog using Win+R
                Console.WriteLine("Opening Run dialog...");
                Keyboard.Press(VirtualKeyShort.LWIN);
                Keyboard.Press(VirtualKeyShort.KEY_R);
                Thread.Sleep(500);
                Keyboard.Release(VirtualKeyShort.KEY_R);
                Keyboard.Release(VirtualKeyShort.LWIN);
                Thread.Sleep(500);

                // Step 2: Type devmgmt.msc to open Device Manager
                Console.WriteLine("Typing devmgmt.msc...");
                Keyboard.Type("devmgmt.msc");
                Thread.Sleep(500);

                // Step 3: Press Enter to launch Device Manager
                Keyboard.Press(VirtualKeyShort.ENTER);
                Thread.Sleep(5000); // Wait for Device Manager to load

                // Step 4: Find Device Manager window
                var desktop = automation.GetDesktop();
                var deviceManagerWindow = desktop.FindFirstChild(cf => cf.ByName("Device Manager").And(cf.ByControlType(ControlType.Window)));

                Assert.That(deviceManagerWindow, Is.Not.Null, "Device Manager window should be opened");
                Console.WriteLine($"Device Manager window found: {deviceManagerWindow.Name}");

                // Attach to the process for cleanup
                deviceManagerApp = Application.Attach(deviceManagerWindow.Properties.ProcessId);

                // Step 5: Find the tree view containing device categories
                var treeView = deviceManagerWindow.FindFirstDescendant(cf =>
                    cf.ByControlType(ControlType.Tree));

                Assert.That(treeView, Is.Not.Null, "Device Manager tree view should be found");
                Console.WriteLine("Device Manager tree view found");

                // Step 6: Get all top-level tree items (device categories)
                var treeItems = treeView.FindAllDescendants(cf => cf.ByControlType(ControlType.TreeItem));
                Console.WriteLine($"Found {treeItems.Length} total tree items initially");

                // Step 7: Expand all categories and collect all items
                var expandedItems = new List<AutomationElement>();
                var processedItems = new HashSet<string>();

                // First pass: expand all top-level categories
                ExpandAllTreeItems(treeView, expandedItems, processedItems, 0);

                Console.WriteLine($"Total expanded items: {expandedItems.Count}");

                // Step 8: Check for error indicators
                var errorDevices = new List<string>();
                var warningDevices = new List<string>();

                foreach (var item in expandedItems)
                {
                    try
                    {
                        var itemName = item.Name;

                        // Check for error/warning indicators in the item name or icon
                        // Device Manager typically shows error devices with special markers

                        // Check if item has a help text that indicates an error
                        var helpText = item.Properties.HelpText.ValueOrDefault;
                        if (!string.IsNullOrEmpty(helpText))
                        {
                            if (helpText.Contains("error", StringComparison.OrdinalIgnoreCase) ||
                                helpText.Contains("not working", StringComparison.OrdinalIgnoreCase) ||
                                helpText.Contains("code ", StringComparison.OrdinalIgnoreCase))
                            {
                                errorDevices.Add($"{itemName} - {helpText}");
                            }
                        }

                        // Check for "(Code " in the name which indicates error codes
                        if (itemName.Contains("(Code ", StringComparison.OrdinalIgnoreCase))
                        {
                            errorDevices.Add(itemName);
                        }

                        // Check the item's state/status by looking for child text elements that might contain status
                        var childTexts = item.FindAllChildren(cf => cf.ByControlType(ControlType.Text));
                        foreach (var childText in childTexts)
                        {
                            var text = childText.Name;
                            if (!string.IsNullOrEmpty(text))
                            {
                                if (text.Contains("error", StringComparison.OrdinalIgnoreCase) ||
                                    text.Contains("not working", StringComparison.OrdinalIgnoreCase) ||
                                    text.Contains("disabled", StringComparison.OrdinalIgnoreCase))
                                {
                                    warningDevices.Add($"{itemName} - Status: {text}");
                                }
                            }
                        }

                        // Check for image elements that might indicate error/warning icons
                        // In Device Manager, error devices typically show with yellow exclamation or red X
                        var images = item.FindAllChildren(cf => cf.ByControlType(ControlType.Image));
                        foreach (var image in images)
                        {
                            var imageName = image.Name;
                            if (!string.IsNullOrEmpty(imageName))
                            {
                                if (imageName.Contains("error", StringComparison.OrdinalIgnoreCase) ||
                                    imageName.Contains("warning", StringComparison.OrdinalIgnoreCase) ||
                                    imageName.Contains("exclamation", StringComparison.OrdinalIgnoreCase))
                                {
                                    errorDevices.Add($"{itemName} - Error icon detected");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error checking item: {ex.Message}");
                    }
                }

                // Step 9: Report findings
                Console.WriteLine("\n=== Device Manager Scan Results ===");
                Console.WriteLine($"Total devices scanned: {expandedItems.Count}");
                Console.WriteLine($"Devices with errors: {errorDevices.Count}");
                Console.WriteLine($"Devices with warnings: {warningDevices.Count}");

                if (errorDevices.Count > 0)
                {
                    Console.WriteLine("\nDevices with ERRORS:");
                    foreach (var error in errorDevices)
                    {
                        Console.WriteLine($"  - {error}");
                    }
                }

                if (warningDevices.Count > 0)
                {
                    Console.WriteLine("\nDevices with WARNINGS:");
                    foreach (var warning in warningDevices)
                    {
                        Console.WriteLine($"  - {warning}");
                    }
                }

                // Step 10: Assert no errors found
                Assert.That(errorDevices.Count, Is.EqualTo(0),
                    $"Found {errorDevices.Count} device(s) with errors in Device Manager");

                Console.WriteLine("\nAll devices are working properly - No errors detected!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed with error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Recursively expands all tree items in the Device Manager tree view
        /// </summary>
        private void ExpandAllTreeItems(AutomationElement parent, List<AutomationElement> expandedItems,
            HashSet<string> processedItems, int depth)
        {
            // Limit recursion depth to prevent infinite loops
            if (depth > 10)
            {
                return;
            }

            try
            {
                // Get all direct children tree items
                var treeItems = parent.FindAllChildren(cf => cf.ByControlType(ControlType.TreeItem));

                foreach (var item in treeItems)
                {
                    try
                    {
                        var itemName = item.Name;
                        var itemId = item.Properties.AutomationId.ValueOrDefault ?? itemName;
                        var uniqueKey = $"{itemId}_{depth}_{itemName}";

                        // Skip if already processed
                        if (processedItems.Contains(uniqueKey))
                        {
                            continue;
                        }

                        processedItems.Add(uniqueKey);
                        expandedItems.Add(item);

                        Console.WriteLine($"{new string(' ', depth * 2)}- {itemName}");

                        // Check if the item is expandable
                        var expandCollapsePattern = item.Patterns.ExpandCollapse.PatternOrDefault;
                        if (expandCollapsePattern != null)
                        {
                            var state = expandCollapsePattern.ExpandCollapseState.Value;

                            if (state == ExpandCollapseState.Collapsed || state == ExpandCollapseState.PartiallyExpanded)
                            {
                                try
                                {
                                    // Expand the item
                                    expandCollapsePattern.Expand();
                                    Thread.Sleep(300); // Wait for expansion

                                    // Recursively process children
                                    ExpandAllTreeItems(item, expandedItems, processedItems, depth + 1);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Could not expand item '{itemName}': {ex.Message}");
                                }
                            }
                            else if (state == ExpandCollapseState.Expanded)
                            {
                                // Already expanded, just process children
                                ExpandAllTreeItems(item, expandedItems, processedItems, depth + 1);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing tree item at depth {depth}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ExpandAllTreeItems at depth {depth}: {ex.Message}");
            }
        }

        [TearDown]
        public void TearDown()
        {
            try
            {
                // Close Device Manager
                deviceManagerApp?.Close();
                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during cleanup: {ex.Message}");
            }
            finally
            {
                automation?.Dispose();
            }
        }
    }
}
