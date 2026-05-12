using BlogGenerator.Agent;
using BlogGenerator.Models;
using BlogGenerator.Services;
using Microsoft.Agents.AI;
using Microsoft.Extensions.AI;
using OpenAI;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.AI.AgentTools.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AITool = Syncfusion.AI.AgentTools.Core.AITool;
using ChatMessage = Microsoft.Extensions.AI.ChatMessage;
using ChatRole = Microsoft.Extensions.AI.ChatRole;

// ════════════════════════════════════════════════════════════
//  BLOG GENERATOR – Console App
//  Powered by Microsoft Agent Framework + OpenAI
// ════════════════════════════════════════════════════════════

PrintBanner();

// ── 1. Configuration ────────────────────────────────────────
var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY")
          ?? PromptRequired("OpenAI API key");
var textModel = Environment.GetEnvironmentVariable("OPENAI_TEXT_MODEL") ?? "gpt-4o";
var imageModel = Environment.GetEnvironmentVariable("OPENAI_IMAGE_MODEL") ?? "gpt-image-1.5";

// ── 2. Build the AIAgent (Microsoft Agent Framework) ────────
var openAiClient = new OpenAIClient(new System.ClientModel.ApiKeyCredential(apiKey));
var chatClient   = openAiClient.GetChatClient(textModel);


string outputDir = Path.GetFullPath(Environment.CurrentDirectory + @"..\..\..\..\Data\Output\");
// ── Set up Syncfusion Word Agent Tools ──────────────────────
var wordManager = new WordDocumentManager(TimeSpan.FromMinutes(10));

var allSyncfusionTools = new List<AITool>();
allSyncfusionTools.AddRange(new WordDocumentAgentTools(wordManager, outputDir).GetTools());
allSyncfusionTools.AddRange(new WordImportExportAgentTools(wordManager).GetTools());

var aiTools = allSyncfusionTools
    .Select(t => AIFunctionFactory.Create(
        t.Method,
        t.Instance,
        new AIFunctionFactoryOptions
        {
            Name = t.Name,
            Description = t.Description
        }))
    .Cast<Microsoft.Extensions.AI.AITool>()
    .ToList();

AIAgent aiAgent = chatClient.AsIChatClient().AsAIAgent(
    instructions: """
        You are an expert technical blogger and document designer.
        You always return only valid JSON when asked, with no markdown fences,
        no extra commentary, and no trailing text outside the JSON object.

        You also have access to Syncfusion Word document tools.
        When asked to create a Word document:
        1. Call CreateDocument to create a new Word document (filePath=null). 
        2. Call ImportHtml with the HTML content or file path and the documentId.
        3. Call ExportDocument with the documentId and the output file path (format "Docx").
        Always follow this sequence and wait for each result before proceeding.
        """,
    name: "BlogGenerationAgent",
    tools: aiTools);

var blogAgent      = new BlogGenerationAgent(aiAgent);
var imageGenerator = new ImageGenerator(openAiClient, imageModel);

// ── 3. Get blog topic from user ──────────────────────────────
Console.WriteLine();
Console.Write("Enter a blog topic: ");
var topic = Console.ReadLine()?.Trim() ?? string.Empty;
if (string.IsNullOrEmpty(topic)) { Console.WriteLine("Topic cannot be empty. Exiting."); return; }

// ── 4. Phase 1 – Generate Title & Outline ───────────────────
BlogOutline outline;
while (true)
{
    Console.WriteLine("\n  Phase 1: Generating title and outline...\n");
    outline = await blogAgent.GenerateOutlineAsync(topic);

    Console.WriteLine("╔══════════════════════════════════════════════════╗");
    Console.WriteLine($"║  {Pad(outline.Title, 48)}║");
    Console.WriteLine("╠══════════════════════════════════════════════════╣");
    for (int i = 0; i < outline.Outline.Count; i++)
        Console.WriteLine($"║  {i + 1,2}. {Pad(outline.Outline[i], 44)}║");
    Console.WriteLine("╚══════════════════════════════════════════════════╝");

    Console.Write("\n  Approve this outline? [Y/n/r(egenerate)]: ");
    var choice = Console.ReadLine()?.Trim().ToUpperInvariant() ?? "Y";

    if (choice == "N") { Console.WriteLine("Exiting by user request."); return; }
    if (choice == "Y" || choice == "") break;
    // "R" or anything else → regenerate
    Console.WriteLine("  Regenerating outline...");
}

// ── 5. Phase 2 – Section & Layout Planning ──────────────────
Console.WriteLine("\n  Phase 2: Planning section layout...");
var planList = await blogAgent.PlanSectionsAsync(outline);
var sections = planList.Sections;

Console.WriteLine($"    Planned {sections.Count} sections:");
foreach (var s in sections)
    Console.WriteLine($"      • [{s.SectionType,8}] {s.Section}{(s.NeedsImage ? " [img]" : "")}");

// ── 6. Phase 3 + 4 – HTML Content & Images ──────────────────
Console.WriteLine("\n  Phase 3 & 4: Generating section content and images...\n");

var blogSections = new List<BlogSection>();
int idx = 0;
foreach (var plan in sections)
{
    Console.WriteLine($"  [{idx + 1}/{sections.Count}] Writing: {plan.Section}");

    // Phase 3 – HTML fragment
    var htmlFragment = await blogAgent.GenerateSectionHtmlAsync(
        outline.Title, plan, idx, sections.Count);

    string? imageBase64 = null;
    string? imageCaption = null;

    // Phase 4 + Image generation
    if (plan.NeedsImage)
    {
        Console.WriteLine($"    --> Generating image prompt for: {plan.ImagePurpose}");
        var imagePrompt = await blogAgent.GenerateImagePromptAsync(outline.Title, plan);
        Console.WriteLine($"    --> Prompt: {imagePrompt}");
        imageBase64 = await imageGenerator.GenerateBase64Async(imagePrompt);
        imageCaption = plan.ImagePurpose;
    }

    blogSections.Add(new BlogSection
    {
        Plan          = plan,
        HtmlFragment  = htmlFragment,
        ImageBase64   = imageBase64,
        ImageCaption  = imageCaption
    });

    idx++;
}

// ── 7. Assemble HTML ─────────────────────────────────────────
Console.WriteLine("\n  Assembling final HTML document...");
var html = HtmlAssembler.Assemble(outline.Title, blogSections);

// ── 8. Save file ─────────────────────────────────────────────
var filename  = HtmlAssembler.DeriveFilename(outline.Title) + ".html";


var filePath = Path.Combine(outputDir, filename);
HtmlAssembler.SaveToFile(filePath, html);

Console.WriteLine($"\n  Done! File saved to:\n    {filePath}\n");
Console.WriteLine($"    Sections : {blogSections.Count}");
Console.WriteLine($"    Images   : {blogSections.Count(s => s.ImageBase64 is not null)}");
Console.WriteLine($"    File size: {new FileInfo(filePath).Length / 1024:N0} KB\n");

// ── 9. Convert HTML to Word Document ─────────────────────────
Console.WriteLine("  Phase 5: Converting HTML to Word document...\n");

var wordFilePath = Path.Combine(outputDir, HtmlAssembler.DeriveFilename(outline.Title) + ".docx");

var history = new List<ChatMessage>();
var userPrompt = $"Create a new Word document, import the HTML from the file '{filePath}' into it, and then export/save it as '{wordFilePath}' in Docx format.";
history.Add(new ChatMessage(ChatRole.User, userPrompt));

var response = await aiAgent.RunAsync(history).ConfigureAwait(false);

foreach (var message in response.Messages)
{
    foreach (var content in message.Contents)
    {
        if (content is TextContent text && !string.IsNullOrEmpty(text.Text))
            Console.WriteLine($"    AI: {text.Text}");
        else if (content is FunctionCallContent call)
            Console.WriteLine($"    [Tool call : {call.Name}]");
        else if (content is FunctionResultContent result)
            Console.WriteLine($"    [Tool result: {result.Result}]");
    }
}

if (File.Exists(wordFilePath))
{
    Console.WriteLine($"\n  Word document saved to:\n    {wordFilePath}");
    Console.WriteLine($"    File size: {new FileInfo(wordFilePath).Length / 1024:N0} KB\n");
}
else
{
    Console.WriteLine("\n  Warning: Word document was not created.\n");
}

// ════════════════════════════════════════════════════════════
//  Local helpers
// ════════════════════════════════════════════════════════════

static void PrintBanner()
{
    Console.ForegroundColor = ConsoleColor.Cyan;
    Console.WriteLine("""
        ╔═══════════════════════════════════════════════╗
        ║       Blog Generator — Powered by             ║
        ║   Microsoft Agent Framework + OpenAI          ║
        ╚═══════════════════════════════════════════════╝
        """);
    Console.ResetColor();
}

static string PromptRequired(string label)
{
    Console.Write($"  Enter {label}: ");
    var value = Console.ReadLine()?.Trim() ?? string.Empty;
    if (string.IsNullOrEmpty(value))
        throw new InvalidOperationException($"{label} is required.");
    return value;
}

static string Pad(string text, int width)
{
    if (text.Length >= width) return text[..(width - 1)] + " ";
    return text.PadRight(width);
}

