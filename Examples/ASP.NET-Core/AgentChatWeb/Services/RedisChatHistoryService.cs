using Microsoft.Extensions.AI;
using Microsoft.Extensions.Caching.Distributed;
using System.Text.Json;
using System.Text.Json.Serialization;

using ChatMessage = Microsoft.Extensions.AI.ChatMessage;
using ChatRole = Microsoft.Extensions.AI.ChatRole;

namespace AgentChatApp.Services;

/// <summary>
/// Persists and retrieves per-session chat history in Redis so that conversation
/// state survives app restarts and is shared across horizontally scaled instances.
/// Key pattern: <c>{InstanceName}{sessionId}</c>
/// </summary>
public sealed class RedisChatHistoryService
{
    private readonly IDistributedCache _cache;
    private readonly TimeSpan _ttl;

    // Serializer options that handle the Microsoft.Extensions.AI polymorphic content types.
    private static readonly JsonSerializerOptions _jsonOptions = BuildJsonOptions();

    public RedisChatHistoryService(IDistributedCache cache, IConfiguration configuration)
    {
        _cache = cache;
        int ttlMinutes = configuration.GetValue<int>("Redis:SessionTtlMinutes", 60);
        _ttl = TimeSpan.FromMinutes(ttlMinutes);
    }

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>Loads the full history for <paramref name="sessionId"/>. Returns an empty list if not found.</summary>
    public async Task<List<ChatMessage>> LoadAsync(string sessionId, CancellationToken ct = default)
    {
        byte[]? data = await _cache.GetAsync(sessionId, ct).ConfigureAwait(false);
        if (data is null or { Length: 0 })
            return [];

        return JsonSerializer.Deserialize<List<SerializableChatMessage>>(data, _jsonOptions)
                   ?.Select(m => m.ToChatMessage())
                   .ToList()
               ?? [];
    }

    /// <summary>Persists the full <paramref name="history"/> for <paramref name="sessionId"/> and slides the TTL.</summary>
    public async Task SaveAsync(string sessionId, List<ChatMessage> history, CancellationToken ct = default)
    {
        var serializable = history.Select(SerializableChatMessage.FromChatMessage).ToList();
        byte[] data = JsonSerializer.SerializeToUtf8Bytes(serializable, _jsonOptions);

        await _cache.SetAsync(sessionId, data,
            new DistributedCacheEntryOptions { SlidingExpiration = _ttl }, ct)
            .ConfigureAwait(false);
    }

    /// <summary>Appends <paramref name="messages"/> to the stored history and slides the TTL.</summary>
    public async Task AppendAsync(string sessionId, IReadOnlyList<ChatMessage> messages, CancellationToken ct = default)
    {
        var history = await LoadAsync(sessionId, ct).ConfigureAwait(false);
        history.AddRange(messages);
        await SaveAsync(sessionId, history, ct).ConfigureAwait(false);
    }

    /// <summary>Deletes the history for <paramref name="sessionId"/>.</summary>
    public async Task DeleteAsync(string sessionId, CancellationToken ct = default)
        => await _cache.RemoveAsync(sessionId, ct).ConfigureAwait(false);

    // ── JSON helpers ──────────────────────────────────────────────────────────

    private static JsonSerializerOptions BuildJsonOptions()
    {
        var opts = new JsonSerializerOptions
        {
            WriteIndented = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };
        return opts;
    }

    // ── Serialisation DTOs ────────────────────────────────────────────────────

    /// <summary>
    /// A flat, JSON-friendly representation of a <see cref="ChatMessage"/>.
    /// Only Text, FunctionCall, and FunctionResult content types are round-tripped
    /// (these are the only types produced by the agent streaming loop).
    /// </summary>
    private sealed class SerializableChatMessage
    {
        public string Role { get; set; } = string.Empty;
        public List<SerializableContent> Contents { get; set; } = [];

        public static SerializableChatMessage FromChatMessage(ChatMessage msg) => new()
        {
            Role     = msg.Role.Value,
            Contents = msg.Contents.Select(SerializableContent.From).ToList()
        };

        public ChatMessage ToChatMessage() => new(
            new ChatRole(Role),
            Contents.Select(c => c.ToAIContent()).ToList());
    }

    private sealed class SerializableContent
    {
        /// <summary>text | functionCall | functionResult</summary>
        public string Type { get; set; } = "text";

        // text
        public string? Text { get; set; }

        // functionCall
        public string? CallId   { get; set; }
        public string? Name     { get; set; }
        public string? ArgsJson { get; set; }

        // functionResult
        public string? ResultJson { get; set; }

        public static SerializableContent From(AIContent c) => c switch
        {
            TextContent t => new() { Type = "text", Text = t.Text },
            FunctionCallContent fc => new()
            {
                Type     = "functionCall",
                CallId   = fc.CallId,
                Name     = fc.Name,
                ArgsJson = fc.Arguments is null ? null
                             : JsonSerializer.Serialize(fc.Arguments, _jsonOptions)
            },
            FunctionResultContent fr => new()
            {
                Type       = "functionResult",
                CallId     = fr.CallId,
                //Name       = fr.Name,
                ResultJson = fr.Result is null ? null
                               : JsonSerializer.Serialize(fr.Result, _jsonOptions)
            },
            _ => new() { Type = "text", Text = c.ToString() }
        };

        public AIContent ToAIContent() => Type switch
        {
            "functionCall" => new FunctionCallContent(
                callId    : CallId ?? string.Empty,
                name      : Name  ?? string.Empty,
                arguments : ArgsJson is null ? null
                              : JsonSerializer.Deserialize<IDictionary<string, object?>>(ArgsJson, _jsonOptions)),
            "functionResult" => new FunctionResultContent(
                callId : CallId ?? string.Empty,
                //name   : Name  ?? string.Empty,
                result : ResultJson is null ? null
                           : JsonSerializer.Deserialize<object>(ResultJson, _jsonOptions)),
            _ => new TextContent(Text ?? string.Empty)
        };
    }
}
