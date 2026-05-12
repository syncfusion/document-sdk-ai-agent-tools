using AgentChatApp.Services;
using Microsoft.AspNetCore.Mvc;

namespace AgentChatApp.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ChatController : ControllerBase
{
    private readonly AgentService _agentService;

    public ChatController(AgentService agentService)
    {
        _agentService = agentService;
    }

    /// <summary>
    /// Sends a message to the AI agent and streams the response as plain text.
    /// </summary>
    [HttpPost("send")]
    public async Task SendMessage([FromBody] ChatRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.Message))
        {
            Response.StatusCode = 400;
            await Response.WriteAsync("Message cannot be empty.", cancellationToken);
            return;
        }

        Response.ContentType = "text/plain; charset=utf-8";
        Response.Headers["Cache-Control"] = "no-cache";
        Response.Headers["X-Accel-Buffering"] = "no";

        try
        {
            await _agentService.StreamResponseAsync(
                request.SessionId,
                request.Message,
                async chunk =>
                {
                    await Response.WriteAsync(chunk, cancellationToken);
                    await Response.Body.FlushAsync(cancellationToken);
                },
                cancellationToken);
        }
        catch (Exception ex) when (!Response.HasStarted)
        {
            Response.StatusCode = 500;
            await Response.WriteAsync($"Agent error: {ex.Message}", cancellationToken);
        }
        catch (Exception ex)
        {
            // Response already started (streaming began) — append error as a final chunk
            await Response.WriteAsync($"\n\n❌ Agent error: {ex.Message}", cancellationToken);
        }
    }

    /// <summary>
    /// Clears the conversation history for the given session.
    /// </summary>
    [HttpDelete("session/{sessionId}")]
    public async Task<IActionResult> ClearSession(string sessionId, CancellationToken cancellationToken)
    {
        await _agentService.ClearSessionAsync(sessionId, cancellationToken);
        return Ok(new { cleared = true });
    }
}

public sealed record ChatRequest(string SessionId, string Message);
