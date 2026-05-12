#region Copyright Syncfusion Inc. 2001 - 2026
//
//  Copyright Syncfusion Inc. 2001 - 2026. All rights reserved.
//
//  Use of this code is subject to the terms of our license.
//  A copy of the current license can be obtained at any time by e-mailing
//  licensing@syncfusion.com. Re-distribution in any form is strictly
//  prohibited. Any infringement will be prosecuted under applicable laws. 
//
#endregion

using System;
using System.Linq;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.XlsIO;

namespace Syncfusion.AI.AgentTools.Excel
{
    /// <summary>
    /// Provides AI agent tools for Excel workbook and worksheet security operations.
    /// Handles encryption, decryption, and protection management.
    /// Supports dual-mode operation: InMemory and DocumentStorage.
    /// </summary>
    public class ExcelSecurityAgentTools : AgentToolBase<IWorkbook>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelSecurityAgentTools"/> class for InMemory mode.
        /// </summary>
        /// <param name="manager">The Excel workbook manager.</param>
        public ExcelSecurityAgentTools(ExcelWorkbookManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelSecurityAgentTools"/> class for DocumentStorage mode.
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public ExcelSecurityAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Encrypts the Excel workbook with a password.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="password">The password to encrypt the workbook.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "EncryptWorkbook", Description = "Encrypts the Excel workbook with the specified password. Supports InMemory and DocumentStorage modes.")]
        public AgentToolResult EncryptWorkbook(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The password to encrypt with")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(password);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Set password for encryption
                workbook.PasswordToOpen = password;
                
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_EncryptWorkbook.xlsx";
                
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Workbook encrypted successfully. Output: {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to encrypt workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes encryption from the Excel workbook.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="password">The current password (for verification).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "DecryptWorkbook", Description = "Removes encryption from the Excel workbook using the provided password. Supports InMemory and DocumentStorage modes.")]
        public AgentToolResult DecryptWorkbook(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The current password")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(password);

                var workbook = OpenDocument(workbookIdOrFilePath, password);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Verify password matches
                if (workbook.PasswordToOpen != password)
                    return AgentToolResult.Fail("Password verification failed");

                // Remove encryption
                workbook.PasswordToOpen = string.Empty;
                
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_DecryptWorkbook.xlsx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Workbook decrypted successfully. Output: {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to decrypt workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Protects the entire workbook structure with a password.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="password">The password to protect with.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "ProtectWorkbook", Description = "Protects the entire workbook structure with a password. Supports InMemory and DocumentStorage modes.")]
        public AgentToolResult ProtectWorkbook(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The password to protect with")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(password);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Protect workbook structure
                workbook.Protect(true, true, password);
                
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_ProtectWorkbook.xlsx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Workbook structure protected successfully. Output: {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to protect workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes workbook structure protection.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="password">The protection password.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "UnprotectWorkbook", Description = "Removes workbook structure protection. Supports InMemory and DocumentStorage modes.")]
        public AgentToolResult UnprotectWorkbook(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The protection password")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(password);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Unprotect workbook
                workbook.Unprotect(password);
                
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_UnprotectWorkbook.xlsx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Workbook unprotected successfully. Output: {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to unprotect workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Protects a worksheet from editing using a password.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet to protect.</param>
        /// <param name="password">The password to protect with.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "ProtectWorksheet", Description = "Protects the worksheet from editing using a password. Supports InMemory and DocumentStorage modes.")]
        public AgentToolResult ProtectWorksheet(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The password to protect with")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(password);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Find worksheet
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Protect worksheet
                worksheet.Protect(password, ExcelSheetProtection.All);
                
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_ProtectWorksheet.xlsx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Worksheet '{worksheetName}' protected successfully. Output: {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to protect worksheet: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes worksheet protection.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="password">The protection password.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "UnprotectWorksheet", Description = "Removes worksheet protection. Supports InMemory and DocumentStorage modes.")]
        public AgentToolResult UnprotectWorksheet(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The protection password")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(password);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Find worksheet
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Unprotect worksheet
                worksheet.Unprotect(password);
                
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_UnprotectWorksheet.xlsx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Worksheet '{worksheetName}' unprotected successfully. Output: {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to unprotect worksheet: {ex.Message}");
            }
        }
    }
}
