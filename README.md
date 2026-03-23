# Outlook Test Framework

This is a C# NUnit test framework for automating Outlook using the Redemption library. It provides a base for testing Outlook functionality, specifically for opening the Outlook classic app and verifying the inbox folder.

## Prerequisites

- .NET SDK installed (using "C:\Program Files\dotnet")
- Microsoft Outlook installed at "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
- Redemption library downloaded and registered

## Setup

1. The Redemption DLL has been downloaded and extracted to the project directory.
2. The DLL is registered using `regsvr32.exe`.
3. The project references the Interop.Redemption.dll.

## Running the Tests

To run the tests:

```bash
dotnet test
```

This will launch Outlook, wait for it to open, access the inbox via Redemption, and verify it exists.

## Extending the Framework

This framework can be extended for ZCO (Zero Carbon Outlook?) testing by adding more test methods in the `Tests` class. For example:

- Test sending emails
- Test folder operations
- Test calendar items
- Etc.

Use the Redemption API for advanced Outlook automation.

## Troubleshooting

- If Outlook fails to launch, ensure the path is correct.
- If Redemption fails, ensure the DLL is properly registered.
- For 64-bit systems, Redemption64.dll is used.

## Dependencies

- NUnit 4.3.2
- Redemption library (Interop.Redemption.dll)