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
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Syncfusion.AI.AgentTools.Core
{
    /// <summary>
    /// Base class for all agent tool classes. Provides an abstract foundation for
    /// creating AI-exposable tools within the Syncfusion Document Processing AI tooling ecosystem.
    /// It enables automatic discovery and conversion of specially annotated instance
    /// methods into AITool objects that can be consumed by AI frameworks.
    /// </summary>
    public abstract class AgentToolBase
    {
        /// <summary>
        /// Discovers all methods marked with the <see cref="ToolAttribute"/> in the derived class
        /// and converts them into <see cref="AITool"/> objects.
        /// </summary>
        /// <returns>A list of AI tools ready for registration with an AI agent.</returns>
        public List<AITool> GetTools()
        {
            var tools = new List<AITool>();
            var type = GetType();

            // Derive prefix from the class name
            string prefix = DeriveToolPrefix(type);

            // Find all public instance methods with ToolAttribute
            var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.GetCustomAttribute<ToolAttribute>() != null);

            foreach (var method in methods)
            {
                var toolAttribute = method.GetCustomAttribute<ToolAttribute>();
                if (toolAttribute == null)
                    continue;

                string toolName = toolAttribute.Name ?? method.Name;
                
                // Apply prefix if derived
                if (!string.IsNullOrEmpty(prefix))
                {
                    toolName = $"{prefix}{toolName}";
                }

                var tool = new AITool
                {
                    Name = toolName,
                    Description = toolAttribute.Description ?? string.Empty,
                    Method = method,
                    Instance = this,
                    Parameters = GetMethodParameters(method)
                };

                tools.Add(tool);
            }

            return tools;
        }

        /// <summary>
        /// Derives the tool prefix from the class name.
        /// Examples:
        /// - WordDocumentAgentTools -> "Word_"
        /// - ExcelWorkbookAgentTools -> "Excel_"
        /// - PdfDocumentAgentTools -> "PDF_"
        /// - PresentationDocumentAgentTools -> "PowerPoint_"
        /// </summary>
        private string DeriveToolPrefix(Type type)
        {
            string className = type.Name;

            // Remove "AgentTools" suffix if present
            if (className.EndsWith("AgentTools"))
            {
                className = className.Substring(0, className.Length - "AgentTools".Length);
            }

            // Map specific class name patterns to prefixes
            if (className.StartsWith("Word"))
                return "Word_";
            
            if (className.StartsWith("Excel"))
                return "Excel_";
            
            if (className.StartsWith("Pdf"))
                return "PDF_";
            
            if (className.StartsWith("Presentation"))
                return "PowerPoint_";

            // Default: no prefix for unrecognized patterns
            return string.Empty;
        }

        /// <summary>
        /// Extracts parameter information from a method.
        /// </summary>
        private List<AIToolParameter> GetMethodParameters(MethodInfo method)
        {
            var parameters = new List<AIToolParameter>();

            foreach (var param in method.GetParameters())
            {
                var paramAttribute = param.GetCustomAttribute<ToolParameterAttribute>();

                parameters.Add(new AIToolParameter
                {
                    Name = param.Name ?? string.Empty,
                    Description = paramAttribute?.Description ?? string.Empty,
                    Type = param.ParameterType,
                    IsOptional = param.HasDefaultValue,
                    DefaultValue = param.DefaultValue
                });
            }

            return parameters;
        }
    }

    /// <summary>
    /// Attribute to mark methods that should be exposed as AI tools.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class ToolAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the name of the tool as exposed to the AI agent.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the description of what the tool does.
        /// This is used by the AI to decide when to call the tool.
        /// </summary>
        public string Description { get; set; }
    }

    /// <summary>
    /// Attribute to provide metadata for tool method parameters.
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false)]
    public class ToolParameterAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the description of the parameter.
        /// </summary>
        public string Description { get; set; }
    }

    /// <summary>
    /// Represents an AI tool that can be registered with an AI agent framework.
    /// </summary>
    public class AITool
    {
        /// <summary>
        /// Gets or sets the name of the tool.
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the description of the tool.
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the method to invoke when the tool is called.
        /// </summary>
        public MethodInfo Method { get; set; }

        /// <summary>
        /// Gets or sets the instance to invoke the method on.
        /// </summary>
        public object Instance { get; set; }

        /// <summary>
        /// Gets or sets the parameter definitions for the tool.
        /// </summary>
        internal List<AIToolParameter> Parameters { get; set; } = new List<AIToolParameter>();
    }

    /// <summary>
    /// Represents a parameter definition for an AI tool.
    /// </summary>
    internal class AIToolParameter
    {
        /// <summary>
        /// Gets or sets the name of the parameter.
        /// </summary>
        internal string Name { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the description of the parameter.
        /// </summary>
        internal string Description { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the type of the parameter.
        /// </summary>
        internal Type Type { get; set; }

        /// <summary>
        /// Gets or sets whether the parameter is optional.
        /// </summary>
        internal bool IsOptional { get; set; }

        /// <summary>
        /// Gets or sets the default value if the parameter is optional.
        /// </summary>
        internal object DefaultValue { get; set; }
    }

    /// <summary>
    /// Represents the result of a tool call.
    /// </summary>
    public class AgentToolResult
    {
        /// <summary>
        /// Gets or sets whether the operation succeeded.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Gets or sets the result message.
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets additional data returned by the tool.
        /// </summary>
        public object Data { get; set; }

        /// <summary>
        /// Gets or sets the error message if the operation failed.
        /// </summary>
        public string Error { get; set; }

        /// <summary>
        /// Creates a successful result.
        /// </summary>
        public static AgentToolResult Ok(string message, object data = null)
        {
            return new AgentToolResult
            {
                Success = true,
                Message = message,
                Data = data
            };
        }

        /// <summary>
        /// Creates a failed result.
        /// </summary>
        public static AgentToolResult Fail(string error)
        {
            return new AgentToolResult
            {
                Success = false,
                Error = error,
                Message = $"Operation failed: {error}"
            };
        }
    }

    /// <summary>
    /// Generic base class for tool classes that operate on a specific document type.
    /// Provides dual-mode infrastructure (<see cref="DocumentManagerMode.InMemory"/> />
    /// <see cref="DocumentManagerMode.DocumentStorage"/>) and two helpers —
    /// <see cref="OpenDocument"/> and <see cref="SaveDocument"/> — that eliminate
    /// per-method mode branching in derived tool classes.
    /// </summary>
    /// <typeparam name="TDocument">
    /// The concrete Syncfusion document type this tool class operates on
    /// (e.g., <c>WordDocument</c>, <c>PdfDocument</c>).
    /// </typeparam>
    /// <remarks>
    /// <para>
    /// Derived tool classes inherit one of two constructors to select the operating mode:
    /// </para>
    /// <list type="bullet">
    ///   <item>
    ///     <description>
    ///       <b>Mode 1 (InMemory)</b> — backed by a <see cref="DocumentManagerBase{TDocument}"/>.
    ///       Documents are live, mutable references held in an in-memory dictionary for the
    ///       session lifetime. Mutations are automatically visible without an explicit save.
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       <b>Mode 2 (DocumentStorage)</b> — backed by a <see cref="DocumentStorageManager"/>
    ///       and user-provided <see cref="IDocumentStorage"/>. Documents are deserialized from
    ///       the storage on each <see cref="OpenDocument"/> call and exist only for the duration
    ///       of a single tool invocation. <see cref="SaveDocument"/> must be called after every
    ///       mutation to persist changes.
    ///     </description>
    ///   </item>
    /// </list>
    /// <para>
    /// The typical tool method pattern is: <c>OpenDocument</c> → mutate → <c>SaveDocument</c>.
    /// Read-only methods may omit the <c>SaveDocument</c> call.
    /// </para>
    /// </remarks>    
    public abstract class AgentToolBase<TDocument> : AgentToolBase
        where TDocument : class
    {
        private readonly DocumentManagerBase<TDocument>? _inMemoryManager;
        private readonly DocumentStorageManager? _storageManager;
        private readonly DocumentManagerMode _mode;
        private readonly DocumentType _documentType;

        protected AgentToolBase() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="AgentToolBase{TDocument}"/> class
        /// in <see cref="DocumentManagerMode.InMemory"/> mode, backed by an in-memory
        /// document manager that holds documents for the session lifetime.
        /// </summary>
        /// <param name="manager">
        /// The in-memory document manager that owns the document dictionary.
        /// </param>
        /// <param name="documentType">
        /// The <see cref="DocumentType"/> this tool class operates on (e.g., <see cref="DocumentType.Word"/>).
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="manager"/> is <see langword="null"/>.
        /// </exception>
        protected AgentToolBase(DocumentManagerBase<TDocument> manager, DocumentType documentType)
        {
            ArgumentNullException.ThrowIfNull(manager);
            _inMemoryManager = manager;
            _documentType = documentType;
            _mode = DocumentManagerMode.InMemory;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AgentToolBase{TDocument}"/> class
        /// in <see cref="DocumentManagerMode.DocumentStorage"/> mode, backed by a
        /// <see cref="DocumentStorageManager"/> and user-provided <see cref="IDocumentStorage"/>.
        /// </summary>
        /// <param name="manager">
        /// The storage manager that delegates all I/O to an <see cref="IDocumentStorage"/> implementation.
        /// </param>
        /// <param name="documentType">
        /// The <see cref="DocumentType"/> this tool class operates on (e.g., <see cref="DocumentType.Word"/>).
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="manager"/> is <see langword="null"/>.
        /// </exception>
        protected AgentToolBase(DocumentStorageManager manager, DocumentType documentType)
        {
            ArgumentNullException.ThrowIfNull(manager);
            _storageManager = manager;
            _documentType = documentType;
            _mode = DocumentManagerMode.DocumentStorage;
        }

        /// <summary>
        /// Gets the current operating mode of this tool class.
        /// </summary>
        /// <value>
        /// <see cref="DocumentManagerMode.InMemory"/> when constructed with a
        /// <see cref="DocumentManagerBase{TDocument}"/>, or
        /// <see cref="DocumentManagerMode.DocumentStorage"/> when constructed with a
        /// <see cref="DocumentStorageManager"/>.
        /// </value>
        protected DocumentManagerMode Mode => _mode;

        /// <summary>
        /// Gets the in-memory document manager used in
        /// <see cref="DocumentManagerMode.InMemory"/> mode.
        /// </summary>
        /// <value>
        /// The <see cref="DocumentManagerBase{TDocument}"/> instance, or <see langword="null"/>
        /// when operating in <see cref="DocumentManagerMode.DocumentStorage"/> mode.
        /// </value>
        protected DocumentManagerBase<TDocument>? InMemoryManager => _inMemoryManager;

        /// <summary>
        /// Gets the document storage manager used in
        /// <see cref="DocumentManagerMode.DocumentStorage"/> mode.
        /// </summary>
        /// <value>
        /// The <see cref="DocumentStorageManager"/> instance, or <see langword="null"/>
        /// when operating in <see cref="DocumentManagerMode.InMemory"/> mode.
        /// </value>
        protected DocumentStorageManager? StorageManager => _storageManager;

        /// <summary>
        /// Retrieves a document by ID or storage path, abstracting away the operating mode.
        /// </summary>
        /// <returns>
        /// The strongly-typed document instance, or <see langword="null"/> if the document
        /// is not found or if <paramref name="documentIdOrFilePath"/> is <see langword="null"/>
        /// in <see cref="DocumentManagerMode.DocumentStorage"/> mode.
        /// </returns>
        protected TDocument? OpenDocument(string? documentIdOrFilePath, string? password = null)
        {
            if (_mode == DocumentManagerMode.InMemory)
            {
                if (IsFilePath(documentIdOrFilePath))
                {
                    if (string.IsNullOrEmpty(password))
                        return _inMemoryManager!.ImportDocument(documentIdOrFilePath);
                    else
                        return _inMemoryManager!.ImportDocument(documentIdOrFilePath, password);
                }
                else
                {
                    return _inMemoryManager!.GetDocument(documentIdOrFilePath);
                }
            }

            if (documentIdOrFilePath == null)
                return null;

            return _storageManager!.GetDocumentInstance(documentIdOrFilePath, _documentType, password) as TDocument;
        }

        private bool IsFilePath(string? input)
        {
            if (string.IsNullOrEmpty(input)) return false;

            // Check for path separators or file extension
            return input.Contains('\\') || input.Contains('/') ||
                   Path.HasExtension(input);
        }

        /// <summary>
        /// Saves a document instance into given file path or document id. No operations in <see cref="DocumentManagerMode.InMemory"/> mode, if documentIdOrFilePath is null or empty.
        /// </summary>
        protected void SaveDocument(string documentIdOrFilePath, TDocument document)
        {
            if (_mode == DocumentManagerMode.DocumentStorage)
                _storageManager!.SaveDocument(documentIdOrFilePath, document, _documentType);
            else if(!string.IsNullOrEmpty(documentIdOrFilePath))
                _inMemoryManager!.ExportDocument(documentIdOrFilePath, document);
        }
        /// <summary>
        /// Saves a document stream to respective file name in the storage and close it, abstracting away the operating mode. No operations in <see cref="DocumentManagerMode.InMemory"/> mode.
        /// </summary>
        protected void SaveFile(string documentIdOrFilePath,Stream stream )
        {
            if (_mode == DocumentManagerMode.DocumentStorage)
                _storageManager!.WriteRawStream(documentIdOrFilePath, stream);
        }
    }
}
