# Cabecalho Logo Inserter

## Word Header Automation and Windows Distribution

Cabecalho Logo Inserter is a small Windows desktop application that inserts an image into the primary header of an existing Microsoft Word document.

The application was created to meet a practical requirement: add a logo after another process had already generated a document. The project was also used as a practical study of packaging and distributing that existing Windows Forms application, including self-contained publishing, installer generation with Inno Setup, version-driven packaging, automated Git tags, and GitHub Releases.

It is intentionally a focused document-automation utility rather than a general-purpose document editor.

---

## Features

* Selects an existing `.docx`, `.docm`, or `.doc` document.
* Accepts a logo in PNG, JPG, JPEG, or BMP format.
* Suggests an output filename with the `_com_logo` suffix while allowing a different destination.
* Copies the source document when the output path differs, preserving the original by default.
* Processes the primary header of every document section through Microsoft Word COM automation.
* Removes an image previously inserted by the application before adding the new one.
* Inserts the logo at a fixed width of 90 points, preserves its aspect ratio, aligns it with the left margin, and places it behind the text.
* Can open the processed document after completion.

> Keep a backup of important documents. Selecting the source file itself as the output path modifies that file directly.

---

## Technologies

| Technology | Usage |
| --- | --- |
| C# | Application implementation |
| .NET 8 | Windows application platform and target framework |
| Windows Forms | Desktop user interface and file dialogs |
| Microsoft Word COM automation | Opens, edits, saves, and closes Word documents through `Word.Application` |
| Inno Setup 6 or 7 | Builds the Windows installer |

The project has no `PackageReference`, explicit `COMReference`, or external assembly reference. Microsoft Word is resolved at runtime through late binding, rather than through the `Microsoft.Office.Interop.Word` NuGet package.

---

## Application Flow

```text
MainForm (View)
       |
       | user events and selected paths
       v
MainController
       |
       | InsercaoLogoModel
       v
IWordLogoService
       |
       v
WordLogoService
       |
       | late-bound COM automation
       v
Microsoft Word Desktop
       |
       v
processed document
```

### Flow summary

The `MainForm` presents the document, logo, and output selectors. `MainController` handles the UI events, builds an `InsercaoLogoModel`, updates the application status, and delegates document processing to `WordLogoService`.

The service validates the selected extensions, creates the output directory when needed, copies the source document when appropriate, and starts an invisible Word instance. For each section, it updates the primary header and saves the processed document. COM objects are explicitly released after processing.

---

## Project Structure

```text
|-- .github/
|    -- workflows/
|       -- release.yml
|-- docs/
|    -- TERMS.txt
|-- src/
|   |-- Controllers/
|   |   -- MainController.cs
|   |-- Models/
|   |    -- InsercaoLogoModel.cs
|   |-- Services/
|   |   |-- IWordLogoService.cs
|   |    -- WordLogoService.cs
|   |-- Views/
|   |   |-- IMainView.cs
|   |    -- MainForm.cs
|   |-- Cabecalho.LogoInserter.csproj
|    -- Program.cs
|-- Cabecalho.LogoInserter.sln
|-- setup.iss
```

| Path | Responsibility |
| --- | --- |
| `Controllers/` | Coordinates view events and the document-processing operation |
| `Models/` | Holds the source document, logo, output path, and post-processing option |
| `Services/` | Validates inputs and performs Word COM automation |
| `Views/` | Implements the Windows Forms interface and file dialogs |

---

## Requirements

### To run the published application

* A 64-bit-compatible Windows system. The release is published for `win-x64`, and the installer accepts only x64-compatible systems.
* Microsoft Word Desktop installed and registered as the `Word.Application` COM server.
* Permission to read the source document and logo and to write to the selected output location.

The CI release build is self-contained, so users do not need to install the .NET 8 runtime separately. Microsoft Word is still required: self-contained publishing includes the .NET runtime, not Microsoft Office.

### To build locally

* Windows.
* The .NET 8 SDK.
* Microsoft Word Desktop to exercise the document-processing feature.
* Inno Setup 6 or 7 only when building the installer.

---

## Commands

Run these commands from the repository root:

| Command | Description |
| --- | --- |
| `dotnet restore .\src\Cabecalho.LogoInserter.csproj` | Restores the project; currently there are no third-party NuGet packages |
| `dotnet build .\Cabecalho.LogoInserter.sln` | Builds the solution |
| `dotnet run --project .\src\Cabecalho.LogoInserter.csproj` | Starts the Windows Forms application |
| `dotnet publish .\src\Cabecalho.LogoInserter.csproj -c Release -r win-x64 --self-contained true -o .\publish` | Produces the same type of self-contained Windows publish used by CI |

The application can also be opened through `Cabecalho.LogoInserter.sln` in a Visual Studio version that supports .NET 8 desktop development.

---

## Usage

1. Select the existing Word document.
2. Select the logo image.
3. Review or change the suggested output path.
4. Choose whether the processed document should open automatically.
5. Select **Insert logo**.

Close the document in Word before processing it to avoid file-locking conflicts. By default, the suggested output is created beside the source as `<original-name>_com_logo.<extension>`.

---

## Publishing

The release workflow publishes the project with these options:

| Option | Value | Effect |
| --- | --- | --- |
| Configuration | `Release` | Produces an optimized release build |
| Runtime identifier | `win-x64` | Targets 64-bit Windows |
| Deployment mode | `--self-contained true` | Includes the .NET runtime in the published output |
| Output directory | `publish/` | Supplies the files consumed by Inno Setup |

Local publishing does not create an installer by itself. `setup.iss` packages the contents of `publish/` after the publish step completes.

---

## Windows Installer

The Inno Setup definition creates a per-user installer for **Cabecalho Logo Inserter**:

* Installs under `%LOCALAPPDATA%\Programs\Cabecalho Logo Inserter`.
* Requests the lowest privilege level and therefore does not require an administrative installation.
* Restricts installation to x64-compatible Windows systems and uses 64-bit install mode.
* Displays [`docs/TERMS.txt`](./docs/TERMS.txt) during setup.
* Creates a Start Menu shortcut.
* Offers an optional, unchecked desktop shortcut.
* Registers uninstall support and provides an application icon in the uninstall entry.
* Offers to run `Cabecalho.LogoInserter.exe` after installation.
* Uses LZMA2 solid compression.

CI supplies the application version and creates the installer as:

```text
Cabecalho.LogoInserter-Setup-{Version}.exe
```

---

## License

This project is open source and available under the [MIT License](./LICENSE).

It was created primarily for study and experimentation and may be freely used, modified, and distributed in accordance with the license terms.

The installer also presents the project's [Terms of Use](./docs/TERMS.txt), which provide additional information about usage, warranty disclaimer, and responsibility when processing documents.

---

## What This Project Covers

* Solving the practical requirement of inserting a logo into an existing Word document header.
* Automating Microsoft Word through late-bound COM.
* Publishing a self-contained application for 64-bit Windows.
* Packaging published output as a per-user installer with Inno Setup.

---

## My Links

* GitHub: https://github.com/torugo99
* LinkedIn: https://www.linkedin.com/in/victor-hugo99
* Website: https://victor99dev.website

---

### Credits and Acknowledgments

* .NET and C#: https://learn.microsoft.com/dotnet/
* Windows Forms: https://learn.microsoft.com/dotnet/desktop/winforms/
* Microsoft Word: https://learn.microsoft.com/office/vba/api/overview/word
* Inno Setup: https://jrsoftware.org/isinfo.php
