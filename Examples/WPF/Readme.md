# Document Processing AI Agent Application - Desktop

## Description

A Windows desktop application that brings AI-powered document processing to your desktop. Built with WPF and powered by OpenAI and Syncfusion Agent tools, it integrates with the [Microsoft Agent Framework](https://learn.microsoft.com/en-us/agent-framework/overview/?pivots=programming-language-csharp) to enable autonomous document operations through an intuitive graphical interface.

The application uses an **in-memory document model** where documents are maintained as live objects within document managers and it will automatically cleaned up after 10 minutes (default) of inactivity. This expiration time is customizable.

## Prerequisites

### Requirements
- .NET 8.0 or later with WPF support
- [OpenAI API key](https://platform.openai.com/api-keys)
- [Syncfusion license key](https://www.syncfusion.com/products/communitylicense)

## How to Run

### 1. Configure API Keys and License
Choose one of the following methods to set up your API credentials and license key:

**Option 1: Set Environment Variables**
```bash
# OpenAI Configuration
setx OPENAI_API_KEY "your-openai-api-key"
setx OPENAI_MODEL "gpt-4o"  # Optional, defaults to gpt-4o

# Syncfusion License
setx SYNCFUSION_LICENSE_KEY "your-syncfusion-license-key"
```

**Option 2: Set API key in Code**
Go to [MainWindow.xaml.cs](./AgentChatDesktop/MainWindow.xaml.cs) and replace with your API key:
``` csharp
string? apiKey = "your-openai-api-key";
string? modelId = "gpt-4o";
string? syncfusionKey = "your-syncfusion-license-key";
```

### 2. Setup
```bash
# Navigate to the project directory
cd Examples/WPF/AgentChatDesktop

# Restore NuGet packages
dotnet restore

# Build the project
dotnet build
```

### 3. Run the Application
```bash
dotnet run
```
Once the application starts, you'll see output similar to:

![Desktop Application Startup](./AgentChatDesktop/Assets/Desktop-Application-output-window.png
)
### Usage

The application uses two folders for file management:

- **[Data/Input](./AgentChatDesktop/Data/Input/)**: Place your input documents here before running the application
  - Supported formats: .docx, .xlsx, .pdf, .pptx, .json, .csv, .md, .txt and more
  - Example files: `new_hire_data.json`, `release_notes_v3.2.md`

- **[Data/Output](./AgentChatDesktop/Data/Output/)**: Generated and processed documents are automatically saved here
  - Contains results from document creation, conversion, and processing operations
  - Organized by operation type for easy access

Simply type your [document processing request at the prompt](https://help.syncfusion.com/document-processing/ai-agent-tools/example-prompts) and press Enter to execute it!

## License
Syncfusion .NET Document SDK library requires a commercial license for production use. A [free community license](https://www.syncfusion.com/products/communitylicense) is available for qualifying organizations.

## Related Resources

- [Syncfusion Agent tools](https://help.syncfusion.com/document-processing/ai-agent-tools/overview)
- [Agent Framework](https://github.com/microsoft/agents)

