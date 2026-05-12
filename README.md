
# Syncfusion Document SDK AI Agent Tools

## Overview

**[Syncfusion Document SDK AI Agent Tool](https://www.nuget.org/packages/Syncfusion.DocumentSDK.AI.AgentTools)** is a .NET library offering comprehensive AI toolkit that enables AI models and assistants to autonomously create, manipulate, convert, and extract data from Word, Excel, PDF, PowerPoint, Markdown, HTML, and RTF documents using [Syncfusion Document SDK](https://www.syncfusion.com/document-sdk) libraries.

It exposes a rich set of pre-defined tools and functions that an [AI agent](https://learn.microsoft.com/en-us/agent-framework/get-started/your-first-agent?pivots=programming-language-csharp) can invoke to perform document operations across various file formats - without requiring the host application to implement document-processing logic directly.

You can quickly deploy it to your infrastructure via [NuGet](https://www.nuget.org/packages/Syncfusion.DocumentSDK.AI.AgentTools). If you want to add new functionality or customize any existing functionalities, then you can use our source code available on [GitHub](https://github.com/syncfusion/document-sdk-ai-agent-tools/tree/master/Syncfusion.DocumentSDK.AI.AgentTools).

### Key Capabilities

* **PDF:** PDF processing with support for digital signing, find text, redactions, watermarking, OCR, and security features (encryption, decryption, and permission management). It also supports splitting, merging, compressing, reordering pages, converting documents, extracting text and images, and importing and exporting annotations.

* **Word:** Word document operations with support for bookmarks, form fields, mail merge, find and replace, document merging and splitting, document comparison, import and export operations, security features (encryption and protection), track changes, and convert Word to formats like Markdown, HTML, RTF, Txt and vice-versa. Also, convert Word to PDF and images.

* **Excel:** Excel document operations with support for charts, conditional formatting, data validation, pivot tables, document deletion, security features (encryption and protection), and conversions to CSV, HTML, and JSON formats.

* **PowerPoint:** PowerPoint presentation operations with support for extracting text, retrieving slides, find and replace operations, merging and splitting presentations, and applying security features (encryption, decryption, and protection). Also, convert PowerPoint to PDF and images.

* **Office to PDF Conversion**: Convert Office documents seamlessly by transforming Excel, Word, and PowerPoint files into PDF format.

* **Smart Data Extraction**: Extract structured information efficiently by retrieving data, converting tables to JSON, converting PDF documents and images to Markdown, and recognizing forms with JSON-based output.


## Dependencies

### Agent Library

- [Syncfusion.DocIO.NET](https://www.nuget.org/packages/Syncfusion.DocIO.Net)
- [Syncfusion.Pdf.NET](https://www.nuget.org/packages/Syncfusion.PDF.Net)
- [Syncfusion.XlsIO.NET](https://www.nuget.org/packages/Syncfusion.XlsIO.NET)
- [Syncfusion.Presentation.NET](https://www.nuget.org/packages/Syncfusion.Presentation.NET)
- [Syncfusion.DocIORenderer.NET](https://www.nuget.org/packages/Syncfusion.DocIORenderer.NET)
- [Syncfusion.PresentationRenderer.NET](https://www.nuget.org/packages/Syncfusion.PresentationRenderer.NET)
- [Syncfusion.XlsIORenderer.NET](https://www.nuget.org/packages/Syncfusion.XlsIORenderer.NET)
- [Syncfusion.SmartDataExtractor.NET](https://www.nuget.org/packages/Syncfusion.SmartDataExtractor.NET)
- [Syncfusion.SmartTableExtractor.NET](https://www.nuget.org/packages/Syncfusion.SmartTableExtractor.NET)
- [Syncfusion.SmartFormRecognizer.NET](https://www.nuget.org/packages/Syncfusion.SmartFormRecognizer.NET)
- [Syncfusion.PDF.OCR.NET](https://www.nuget.org/packages/Syncfusion.PDF.OCR.NET)

## Getting Started

[Agent Tools](https://learn.microsoft.com/en-us/agent-framework/get-started/add-tools?pivots=programming-language-csharp) are the callable functions exposed to the AI agent. Each tool class is initialized with the appropriate manager.

Agent tools support two operational modes that determine how documents are handled during AI agent execution. In‑Memory mode enables live, in‑memory document processing, while Document Storage mode supports persistent, storage‑backed document handling. The operational mode is determined by the manager used when initializing the tool.

### Document Manager Modes

- [In-Memory Mode](https://help.syncfusion.com/document-processing/ai-agent-tools/getting-started-in-memory-mode)
- [Storage Mode](https://help.syncfusion.com/document-processing/ai-agent-tools/getting-started-storage-mode)


### Prerequisites

- .NET 8.0 or higher
- [OpenAI API key](https://platform.openai.com/login)
- [Syncfusion license](https://www.syncfusion.com/products/communitylicense) (community or commercial)

### Installation and Configure

1. Clone the repository:
   ```bash
   git clone https://github.com/syncfusion/document-sdk-ai-agent-tools.git
   cd Document-SDK-AI-Agent-Tools
   ```
2. Build the library:
   ```bash
   dotnet build Syncfusion.DocumentSDK.AI.AgentTools\Syncfusion.DocumentSDK.AI.AgentTools.csproj
   ```
3. Launch the example application:

   | Platforms        | Demo Link |
   |------------------|-----------|
   | ASP.NET Core     |  [AgentChatWeb](./Examples/ASP.NET-Core/)        |
   | WPF              | [AgentChatDesktop](./Examples/WPF/)          |
   | Console          |    [AgentChatConsole](./Examples/Console/)       |


## Documentation

- [Overview](https://help.syncfusion.com/document-processing/ai-agent-tools/overview)
- [Getting Started](https://help.syncfusion.com/document-processing/ai-agent-tools/getting-started)
- [Tools](https://help.syncfusion.com/document-processing/ai-agent-tools/tools)
- [Customization](https://help.syncfusion.com/document-processing/ai-agent-tools/customization)
- [Example Prompts](https://help.syncfusion.com/document-processing/ai-agent-tools/example-prompts)

## License

Copyright Syncfusion Inc. 2001 - 2026. All rights reserved.

Syncfusion Document SDK AI Agent Tools are included as a part of Syncfusion Document SDK license. No additional purchase is required, and the same Document SDK license key is enough.

This is a commercial product and requires a paid license for possession or use. Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/eula/es/?utm_source=nuget&utm_medium=listing&utm_campaign=net-document-sdk-ai-agent-tools-nuget). You can purchase a license [here](https://www.syncfusion.com/sales/products?utm_source=nuget&utm_medium=listing&utm_campaign=net-document-sdk-ai-agent-tools-nuget) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=nuget&utm_medium=listing&utm_campaign=net-document-sdk-ai-agent-tools-nuget).
## Support

- **Documentation**: [Syncfusion Help Documentation](https://help.syncfusion.com)
- **Email**: support@syncfusion.com
- **Forum**: [Syncfusion Community Forums](https://www.syncfusion.com/forums)

## Related Resources

- [PDF Library Documentation](https://help.syncfusion.com/document-processing/pdf/pdf-library/overview)
- [Word Library Documentation](https://help.syncfusion.com/document-processing/word/word-library/overview)
- [Excel Library Documentation](https://help.syncfusion.com/document-processing/excel/excel-library/overview)
- [PowerPoint Library Documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/overview)
- [Data Extraction Documentation](https://help.syncfusion.com/document-processing/data-extraction/overview)

---

