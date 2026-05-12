using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using BlogGenerator.Layout;
using BlogGenerator.Models;

namespace BlogGenerator.Services
{
    /// <summary>
    /// Assembles the final self-contained HTML document from all blog sections.
    /// CSS is always sourced from <see cref="CssTemplate"/> – never from the agent.
    /// </summary>
    public static class HtmlAssembler
    {
        /// <summary>
        /// Builds and returns the full HTML string.
        /// </summary>
        public static string Assemble(string blogTitle, List<BlogSection> sections)
        {
            var sb = new StringBuilder();

            // ── Document head ──────────────────────────────────────────
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"en\">");
            sb.AppendLine("<head>");
            sb.AppendLine("  <meta charset=\"UTF-8\" />");
            sb.AppendLine("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />");
            sb.AppendLine($"  <title>{HtmlEncode(blogTitle)}</title>");
            sb.AppendLine("  <style>");
            sb.AppendLine(CssTemplate.Styles);
            sb.AppendLine("  </style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");

            // ── Table of Contents ──────────────────────────────────────
            sb.AppendLine(BuildToc(sections));

            // ── Sections ───────────────────────────────────────────────
            foreach (var section in sections)
            {
                sb.AppendLine($"  <section id=\"{SectionId(section.Plan.Section)}\">");

                // HTML fragment from the agent
                sb.AppendLine(section.HtmlFragment);

                // Inject image (if any) after the first paragraph
                if (section.ImageBase64 is { Length: > 0 })
                {
                    sb.AppendLine("  <div class=\"FigureWrapper\">");
                    sb.AppendLine($"    <img src=\"data:image/png;base64,{section.ImageBase64}\" alt=\"{HtmlEncode(section.Plan.ImagePurpose)}\" width=\"600\" height=\"300\" />");
                    sb.AppendLine($"    <p class=\"FigureCaptionP\">Figure: {HtmlEncode(section.ImageCaption ?? section.Plan.ImagePurpose)}</p>");
                    sb.AppendLine("  </div>");
                }

                sb.AppendLine("  </section>");
                sb.AppendLine("  <hr class=\"ChapterDivider\" />");
            }

            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
        }

        /// <summary>
        /// Derives a safe filename (no extension) from the blog title.
        /// </summary>
        public static string DeriveFilename(string title)
        {
            var safe = Regex.Replace(title, @"[^a-zA-Z0-9\s\-]", "");
            safe = Regex.Replace(safe.Trim(), @"\s+", "-");
            if (safe.Length > 80) safe = safe[..80];
            return safe;
        }

        /// <summary>Saves the HTML content to <paramref name="filePath"/>.</summary>
        public static void SaveToFile(string filePath, string htmlContent)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            File.WriteAllText(filePath, htmlContent, Encoding.UTF8);
        }

        // ── Private helpers ────────────────────────────────────────────
        private static string BuildToc(List<BlogSection> sections)
        {
            var sb = new StringBuilder();
            sb.AppendLine("  <nav class=\"TocSection\">");
            sb.AppendLine("    <p class=\"TocHeadingP\">Table of Contents</p>");
            foreach (var s in sections)
            {
                var id = SectionId(s.Plan.Section);
                sb.AppendLine($"    <p class=\"TocItemP\"><a href=\"#{id}\">{HtmlEncode(s.Plan.Section)}</a></p>");
            }
            sb.AppendLine("  </nav>");
            return sb.ToString();
        }

        private static string SectionId(string title)
            => Regex.Replace(title.ToLowerInvariant(), @"[^a-z0-9]+", "-").Trim('-');

        private static string HtmlEncode(string text)
            => System.Web.HttpUtility.HtmlEncode(text);
    }
}
