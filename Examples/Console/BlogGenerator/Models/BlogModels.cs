using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace BlogGenerator.Models
{
    /// <summary>Phase 1 output: title + ordered outline items.</summary>
    public class BlogOutline
    {
        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        [JsonPropertyName("outline")]
        public List<string> Outline { get; set; } = [];
    }

    /// <summary>Phase 2 output: per-section layout & image decision.</summary>
    public class SectionPlan
    {
        [JsonPropertyName("section")]
        public string Section { get; set; } = string.Empty;

        /// <summary>cover | chapter | appendix</summary>
        [JsonPropertyName("sectionType")]
        public string SectionType { get; set; } = "chapter";

        /// <summary>H1 or H2</summary>
        [JsonPropertyName("headingLevel")]
        public string HeadingLevel { get; set; } = "H2";

        /// <summary>Whether this section needs a generated image.</summary>
        [JsonPropertyName("needsImage")]
        public bool NeedsImage { get; set; }

        /// <summary>Short description of what the image should depict.</summary>
        [JsonPropertyName("imagePurpose")]
        public string ImagePurpose { get; set; } = string.Empty;
    }

    public class SectionPlanList
    {
        [JsonPropertyName("sections")]
        public List<SectionPlan> Sections { get; set; } = [];
    }

    /// <summary>Phase 3 output: raw HTML fragment for one section.</summary>
    public class SectionHtml
    {
        [JsonPropertyName("section")]
        public string Section { get; set; } = string.Empty;

        [JsonPropertyName("html")]
        public string Html { get; set; } = string.Empty;
    }

    /// <summary>Phase 4 output: image generation prompt for one section.</summary>
    public class SectionImagePrompt
    {
        [JsonPropertyName("section")]
        public string Section { get; set; } = string.Empty;

        [JsonPropertyName("imagePrompt")]
        public string ImagePrompt { get; set; } = string.Empty;
    }

    /// <summary>Aggregated result ready for HTML assembly.</summary>
    public class BlogSection
    {
        public SectionPlan Plan { get; set; } = new();
        public string HtmlFragment { get; set; } = string.Empty;
        public string? ImageBase64 { get; set; }
        public string? ImageCaption { get; set; }
    }
}
