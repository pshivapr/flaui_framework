# UI Automation Framework

A Windows UI automation testing framework built with [FlaUI](https://github.com/FlaUI/FlaUI) and NUnit for automating desktop applications.

## Overview

This project demonstrates UI automation testing for Windows desktop applications using:

- **FlaUI 5.0** - UI automation library for .NET
- **UIA3** - UI Automation v3 provider
- **NUnit 4.2** - Testing framework
- **.NET 9.0** - Target framework

## Project Structure

```
ui-automation-framework/
├── flaui-demo/
│   ├── MicrosoftOfficeAppsTests.cs    # Tests for PowerPoint, Word, Excel
│   ├── NotepadTest.cs                 # Basic Notepad automation test
│   └── flaui-demo.csproj             # Project configuration
├── ui-automation-framework.sln        # Solution file
└── README.md                          # This file
```

## Prerequisites

- Windows OS
- .NET 9.0 SDK or later
- Visual Studio 2022 or VS Code with C# extension
- Microsoft Office (for Office app tests)

## Getting Started

### 1. Clone the Repository

```bash
git clone https://github.com/pshivapr/flaui_framework.git
cd ui-automation-framework
```

### 2. Restore Dependencies

```bash
dotnet restore
```

## Building the Project

### Build using .NET CLI

```bash
# Build the entire solution
dotnet build ui-automation-framework.sln

# Or build the specific project
dotnet build flaui-demo/flaui-demo.csproj
```

### Build using Visual Studio

1. Open `ui-automation-framework.sln` in Visual Studio
2. Press `Ctrl+Shift+B` or select **Build > Build Solution**

## Running Tests

### Run All Tests

```bash
# Using dotnet test
dotnet test

# With verbose output
dotnet test --logger "console;verbosity=detailed"
```

### Run Specific Test

```bash
# Run by test name
dotnet test --filter "Name~LaunchNotepadAndVerifyTitle"

# Run specific test class
dotnet test --filter "FullyQualifiedName~MicrosoftOfficeAppsTests"
```

### Run Tests in Visual Studio

1. Open **Test Explorer** (`Ctrl+E, T`)
2. Click **Run All** or right-click individual tests to run

### Run Tests in VS Code

1. Install the **.NET Core Test Explorer** extension
2. Open Test Explorer from the sidebar
3. Click the play button next to tests

## Test Categories

### Notepad Tests

- **VerifyTitle**: Launches Notepad and verifies the window title

### Microsoft Office Tests

- **LaunchPowerPointAndCheckVersion**: Launches PowerPoint, retrieves and validates version information

## Key Features

- ✅ Automated UI testing for Windows desktop applications
- ✅ Support for Microsoft Office applications (PowerPoint, Word)
- ✅ Version detection and validation
- ✅ Proper resource cleanup and disposal
- ✅ NUnit test framework integration

## Inpsecting Window Elements

To inspect elements use FlaUInspect, which allows you to find window elements using AutomationId, Name or XPath.

To install FlaUInspect, either build it yourself or get the zip from the releases page here on GitHub (<https://github.com/FlaUI/FlaUInspect/releases>).

## Dependencies

```xml
<PackageReference Include="FlaUI.Core" Version="5.0.0" />
<PackageReference Include="FlaUI.UIA3" Version="5.0.0" />
<PackageReference Include="NUnit" Version="4.2.2" />
<PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.12.0" />
<PackageReference Include="NUnit3TestAdapter" Version="4.6.0" />
```

## Troubleshooting

### Tests Fail to Find Applications

- Ensure the application is installed on your system
- Check the application path in the test code
- For Office apps, verify the installation path (typically `C:\Program Files\Microsoft Office\root\Office16\`)

### UI Automation Access Denied

- Run Visual Studio or your terminal as Administrator
- Enable UI Automation in Windows accessibility settings

### Process Not Closing After Test

- Tests include cleanup logic in `TearDown` methods
- If processes hang, manually terminate them from Task Manager

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-test`)
3. Commit your changes (`git commit -am 'Add new test'`)
4. Push to the branch (`git push origin feature/new-test`)
5. Create a Pull Request

## License

This project is for demonstration purposes.

## Resources

- [FlaUI Documentation](https://github.com/FlaUI/FlaUI)
- [NUnit Documentation](https://docs.nunit.org/)
- [Windows UI Automation](https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32)
