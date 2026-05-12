namespace BlogGenerator.Layout
{
    /// <summary>
    /// Provides the static CSS block embedded in every generated HTML file.
    /// The agent is never allowed to modify or generate CSS.
    /// </summary>
    public static class CssTemplate
    {
        public const string Styles = """
            /* ===== BLOG GENERATOR – EBOOK STYLE ===== */
            *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

            body {
                font-family: 'Georgia', 'Times New Roman', serif;
                font-size: 17px;
                line-height: 1.75;
                color: #1a1a2e;
                background: #fafaf8;
                max-width: 900px;
                margin: 0 auto;
                padding: 2rem 2.5rem 4rem;
            }

            /* ── Cover ────────────────────────────── */
            .CoverSection {
                background: linear-gradient(135deg, #0f3460 0%, #16213e 60%, #0a3d62 100%);
                color: #e8e8e8;
                padding: 5rem 4rem 4rem;
                border-radius: 8px;
                margin-bottom: 3rem;
                text-align: center;
            }
            .BookTitleP {
                font-size: 2.8rem;
                font-weight: 700;
                color: #f5c518;
                line-height: 1.2;
                margin-bottom: 1rem;
                letter-spacing: -0.5px;
            }
            .BookSubtitleP {
                font-size: 1.25rem;
                color: #b0c4de;
                font-style: italic;
                margin-bottom: 2rem;
            }
            .CoverAuthorP {
                font-size: 1rem;
                color: #87ceeb;
                text-transform: uppercase;
                letter-spacing: 2px;
            }

            /* ── Table of Contents ─────────────────── */
            .TocSection {
                background: #f0f4f8;
                border-left: 5px solid #0f3460;
                padding: 2rem 2.5rem;
                border-radius: 4px;
                margin-bottom: 3rem;
            }
            .TocHeadingP {
                font-size: 1.4rem;
                font-weight: 700;
                color: #0f3460;
                margin-bottom: 1rem;
                text-transform: uppercase;
                letter-spacing: 1px;
            }
            .TocItemP {
                font-size: 1rem;
                color: #2c3e50;
                padding: 0.2rem 0;
                border-bottom: 1px dotted #c0ccd8;
            }
            .TocItemP a { color: #0f3460; text-decoration: none; }
            .TocItemP a:hover { text-decoration: underline; }

            /* ── Headings ──────────────────────────── */
            .Heading_1 {
                font-size: 2.2rem;
                font-weight: 700;
                color: #0f3460;
                border-bottom: 3px solid #f5c518;
                padding-bottom: 0.4rem;
                margin: 3rem 0 1.2rem;
            }
            .Heading_2 {
                font-size: 1.6rem;
                font-weight: 600;
                color: #16213e;
                margin: 2.5rem 0 0.8rem;
                border-left: 4px solid #0f3460;
                padding-left: 0.75rem;
            }
            .Heading_3 {
                font-size: 1.2rem;
                font-weight: 600;
                color: #2c3e50;
                margin: 2rem 0 0.5rem;
            }

            /* ── Body Text ─────────────────────────── */
            .Body_Text {
                font-family: 'Helvetica Neue', Arial, sans-serif;
                font-size: 1rem;
                color: #2c2c2c;
                margin-bottom: 1rem;
            }
            .ChapterLeadP {
                font-family: 'Helvetica Neue', Arial, sans-serif;
                font-size: 1.2rem;
                color: #34495e;
                font-style: italic;
                margin-bottom: 1.5rem;
                padding: 0.75rem 1.25rem;
                border-left: 4px solid #f5c518;
                background: #fffbea;
                border-radius: 2px;
            }

            /* ── Callout Box ───────────────────────── */
            .CalloutBox {
                background: #e8f4fd;
                border: 1px solid #aed6f1;
                border-left: 5px solid #2980b9;
                padding: 1.25rem 1.5rem;
                border-radius: 4px;
                margin: 1.5rem 0;
            }
            .CalloutHeadingP {
                font-weight: 700;
                color: #1a5276;
                font-size: 1rem;
                margin-bottom: 0.4rem;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            .CalloutBodyP {
                font-family: 'Helvetica Neue', Arial, sans-serif;
                font-size: 0.95rem;
                color: #1a3a4a;
            }

            /* ── Figure / Image ────────────────────── */
            .FigureWrapper {
                text-align: center;
                margin: 2rem auto;
            }
            .FigureWrapper img {
                max-width: 100%;
                border-radius: 6px;
                box-shadow: 0 4px 18px rgba(0,0,0,0.12);
            }
            .FigureCaptionP {
                font-family: 'Helvetica Neue', Arial, sans-serif;
                font-size: 0.85rem;
                color: #666;
                font-style: italic;
                margin-top: 0.5rem;
            }

            /* ── Code Block ────────────────────────── */
            .CodeBlock {
                background: #1e1e2e;
                color: #cdd6f4;
                font-family: 'Cascadia Code', 'Fira Code', Consolas, monospace;
                font-size: 0.88rem;
                padding: 1.25rem 1.5rem;
                border-radius: 6px;
                overflow-x: auto;
                margin: 1.25rem 0;
                white-space: pre;
                line-height: 1.5;
            }

            /* ── Table ─────────────────────────────── */
            .DataTable {
                width: 100%;
                border-collapse: collapse;
                margin: 1.5rem 0;
                font-family: 'Helvetica Neue', Arial, sans-serif;
                font-size: 0.95rem;
            }
            .DataTable th {
                background: #0f3460;
                color: #fff;
                padding: 0.6rem 0.9rem;
                text-align: left;
            }
            .DataTable td {
                padding: 0.55rem 0.9rem;
                border-bottom: 1px solid #dde3ea;
            }
            .DataTable tr:nth-child(even) td { background: #f5f7fa; }

            /* ── Lists ─────────────────────────────── */
            ul.BulletList, ol.NumberList {
                font-family: 'Helvetica Neue', Arial, sans-serif;
                font-size: 1rem;
                color: #2c2c2c;
                margin: 0.5rem 0 1rem 1.5rem;
            }
            ul.BulletList li, ol.NumberList li { margin-bottom: 0.35rem; }

            /* ── Chapter Divider ───────────────────── */
            .ChapterDivider {
                border: none;
                border-top: 2px solid #e0e6ef;
                margin: 3rem 0;
            }

            /* ── Executive Summary ─────────────────── */
            .ExecSummaryBox {
                background: #f9f0ff;
                border: 1px solid #d4a8f0;
                border-left: 5px solid #8e44ad;
                padding: 1.5rem 2rem;
                border-radius: 4px;
                margin-bottom: 3rem;
            }
            .ExecSummaryHeadingP {
                font-size: 1.2rem;
                font-weight: 700;
                color: #6c3483;
                margin-bottom: 0.75rem;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }

            /* ── Print ─────────────────────────────── */
            @media print {
                body { padding: 1rem; background: #fff; }
                .CoverSection { background: #0f3460 !important; -webkit-print-color-adjust: exact; }
            }
            """;
    }
}
