using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using BlogGenerator.Models;
using Microsoft.Agents.AI;

namespace BlogGenerator.Agent
{
    /// <summary>
    /// Single agent responsible for all four generation phases.
    /// Uses Microsoft.Agents.AI.OpenAI (ChatCompletion backend).
    /// </summary>
    public class BlogGenerationAgent
    {
        private readonly AIAgent _agent;

        private static readonly JsonSerializerOptions _jsonOptions = new()
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        public BlogGenerationAgent(AIAgent agent)
        {
            _agent = agent;
        }

        // ─────────────────────────────────────────────────────────────
        // Phase 1 – Title & Outline
        // ─────────────────────────────────────────────────────────────
        public async Task<BlogOutline> GenerateOutlineAsync(string topic)
        {
            var schema = """
                {
                  "title": "string",
                  "outline": ["section title 1", "section title 2", "..."]
                }
                """;

            var prompt = "You are a professional technical writer and blogger.\n"
                + "Generate a compelling blog title and a detailed outline for the following topic.\n\n"
                + $"TOPIC: {topic}\n\n"
                + "Rules:\n"
                + "- The outline must have 6-10 sections.\n"
                + "- Each section is a short descriptive title (no numbering).\n"
                + "- Always start with an \"Introduction\" section and end with a \"Conclusion\" section.\n"
                + "- Return ONLY valid JSON. No markdown fences, no extra text outside the JSON.\n\n"
                + "Return exactly this JSON shape:\n"
                + schema;

            return await RunWithRetryAsync(prompt, ParseBlogOutline, maxRetries: 3);
        }

        // ─────────────────────────────────────────────────────────────
        // Phase 2 – Section & Layout Planning
        // ─────────────────────────────────────────────────────────────
        public async Task<SectionPlanList> PlanSectionsAsync(BlogOutline outline)
        {
            var outlineJson = JsonSerializer.Serialize(outline, _jsonOptions);

            var jsonSchema = """
                {
                  "sections": [
                    {
                      "section": "string",
                      "sectionType": "string",
                      "headingLevel": "string",
                      "needsImage": bool,
                      "imagePurpose": "string"
                    }
                  ]
                }
                """;

            var prompt = $"""
                You are a document layout designer for a technical blog / ebook.
                Given the blog outline below, plan each section's layout and image requirements.

                OUTLINE JSON:
                {outlineJson}

                Rules for each section:
                - sectionType must be one of: "cover", "chapter", "appendix"
                  * The first section → "cover"
                  * The last section  → "appendix"
                  * All others        → "chapter"
                - headingLevel: "H1" for cover/appendix; "H2" for chapter
                - needsImage: true for 40-60% of chapters (always false for cover)
                - imagePurpose: short description of what the image should visualize (empty string if needsImage=false)
                - Return ONLY valid JSON. No markdown fences, no extra text.

                JSON schema:
                {jsonSchema}
                """;

            return await RunWithRetryAsync(prompt, ParseSectionPlanList, maxRetries: 3);
        }

        // ─────────────────────────────────────────────────────────────
        // Phase 3 – Rich HTML Content Generation (one section at a time)
        // ─────────────────────────────────────────────────────────────
        public async Task<string> GenerateSectionHtmlAsync(
            string blogTitle,
            SectionPlan plan,
            int sectionIndex,
            int totalSections)
        {
            var headingTag = plan.HeadingLevel.ToUpperInvariant() == "H1" ? "h1" : "h2";
            var cssClass = plan.HeadingLevel.ToUpperInvariant() == "H1" ? "Heading_1" : "Heading_2";

            var jsonSchema = """
                {
                  "section": "string (exact section title)",
                  "html": "string (HTML fragment)"
                }
                """;

            var prompt = $"""
                You are a professional HTML content writer for a technical blog ebook.
                Write the HTML content fragment for the section described below.

                BLOG TITLE   : {blogTitle}
                SECTION TITLE: {plan.Section}
                SECTION TYPE : {plan.SectionType}
                HEADING TAG  : <{headingTag} class="{cssClass}">
                POSITION     : Section {sectionIndex + 1} of {totalSections}

                STRICT RULES:
                1. Use ONLY these CSS class names: Heading_1, Heading_2, Heading_3,
                   Body_Text, ChapterLeadP, CalloutBox, CalloutHeadingP, CalloutBodyP,
                   BulletList, NumberList, DataTable, CodeBlock, ExecSummaryBox,
                   ExecSummaryHeadingP, BookTitleP, BookSubtitleP, CoverAuthorP
                2. NO inline CSS, NO <style> tags, NO <img> tags, NO <script> tags.
                3. NO markdown – pure HTML fragment only.
                4. Start with the heading tag. Write 3-6 rich paragraphs / lists / callouts.
                5. For the cover section include a <div class="CoverSection"> wrapper containing
                   BookTitleP, BookSubtitleP, and CoverAuthorP paragraphs.
                6. Return ONLY valid JSON. No markdown fences.

                JSON schema:
                {jsonSchema}
                """;

            var result = await RunWithRetryAsync(prompt, ParseSectionHtml, maxRetries: 3);
            return result.Html;
        }

        // ─────────────────────────────────────────────────────────────
        // Phase 4 – Image Prompt Generation
        // ─────────────────────────────────────────────────────────────
        public async Task<string> GenerateImagePromptAsync(string blogTitle, SectionPlan plan)
        {
            var prompt = $"""
                You are an expert at writing image generation prompts for technical blogs.

                BLOG TITLE   : {blogTitle}
                SECTION TITLE: {plan.Section}
                IMAGE PURPOSE: {plan.ImagePurpose}

                Write a concise, vivid, editorial-style image prompt (max 200 chars).
                Style: clean flat-design infographic, professional color palette (blues, whites, gold accents).
                NO text overlays in the image.
                Return ONLY the plain prompt string – no JSON, no quotes, no extra text.
                """;

            var response = await _agent.RunAsync(prompt);
            return response.Text?.Trim() ?? string.Empty;
        }

        // ─────────────────────────────────────────────────────────────
        // Helpers
        // ─────────────────────────────────────────────────────────────
        private async Task<T> RunWithRetryAsync<T>(
            string prompt,
            Func<string, T?> parser,
            int maxRetries)
        {
            Exception? last = null;
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    var response = await _agent.RunAsync(prompt);
                    var text = response.Text?.Trim() ?? string.Empty;

                    // Strip possible markdown fences the model may include despite instructions
                    text = StripMarkdownFence(text);

                    var result = parser(text);
                    if (result is not null)
                        return result;

                    Console.WriteLine($"  [Agent] Parse returned null on attempt {attempt}, retrying...");
                }
                catch (Exception ex)
                {
                    last = ex;
                    Console.WriteLine($"  [Agent] Attempt {attempt} failed: {ex.Message}");
                    await Task.Delay(1500 * attempt);
                }
            }

            throw new InvalidOperationException(
                $"Agent failed after {maxRetries} attempts.", last);
        }

        private static string StripMarkdownFence(string text)
        {
            if (text.StartsWith("```"))
            {
                var firstNewline = text.IndexOf('\n');
                if (firstNewline >= 0) text = text[(firstNewline + 1)..];
                if (text.EndsWith("```")) text = text[..^3];
                text = text.Trim();
            }
            return text;
        }

        private static BlogOutline? ParseBlogOutline(string json)
        {
            try { return JsonSerializer.Deserialize<BlogOutline>(json, _jsonOptions); }
            catch { return null; }
        }

        private static SectionPlanList? ParseSectionPlanList(string json)
        {
            try { return JsonSerializer.Deserialize<SectionPlanList>(json, _jsonOptions); }
            catch { return null; }
        }

        private static SectionHtml? ParseSectionHtml(string json)
        {
            try { return JsonSerializer.Deserialize<SectionHtml>(json, _jsonOptions); }
            catch { return null; }
        }
    }
}
