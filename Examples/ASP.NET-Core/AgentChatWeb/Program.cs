using AgentChatApp.Services;
using Microsoft.AspNetCore.Http.Features;

var builder = WebApplication.CreateBuilder(args);

// ── Distributed cache (chat history store) ────────────────────────────────
// In Production: use Redis (required for multi-instance Azure App Service deployments).
// In Development or when no Redis connection string is provided: fall back to the
// in-memory distributed cache so the app works without a local Redis server.
string? redisConn = builder.Configuration["Redis:ConnectionString"]
    ?? Environment.GetEnvironmentVariable("REDIS_CONNECTION_STRING");

string redisInstance = builder.Configuration["Redis:InstanceName"] ?? "AgentChat:";

bool useRedis = !string.IsNullOrWhiteSpace(redisConn);

if (useRedis)
{
    builder.Services.AddStackExchangeRedisCache(options =>
    {
        options.Configuration = redisConn;
        options.InstanceName  = redisInstance;
    });
}
else
{
    // Local dev fallback — single-process in-memory cache (not suitable for scale-out).
    builder.Services.AddDistributedMemoryCache();
}

// ── Services ───────────────────────────────────────────────────────────────
builder.Services.AddControllers();
builder.Services.AddSingleton<RedisChatHistoryService>();    // Redis-backed chat history
builder.Services.AddSingleton<AgentService>();               // shared agent + single DocumentStorageManager

// Surface unhandled exceptions with a plain-text body (useful in deployed environments)
builder.Services.AddProblemDetails();

// Allow uploads up to 100 MB (Azure App Service default limit is 30 MB — raised here)
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 100 * 1024 * 1024; // 100 MB
});
builder.WebHost.ConfigureKestrel(k =>
{
    k.Limits.MaxRequestBodySize = 100 * 1024 * 1024; // 100 MB
});

var app = builder.Build();

// ── Middleware ─────────────────────────────────────────────────────────────
app.UseExceptionHandler(errApp =>
{
    errApp.Run(async ctx =>
    {
        ctx.Response.StatusCode = 500;
        ctx.Response.ContentType = "text/plain; charset=utf-8";
        var feature = ctx.Features.Get<Microsoft.AspNetCore.Diagnostics.IExceptionHandlerFeature>();
        var msg = feature?.Error?.Message ?? "An unexpected error occurred.";
        await ctx.Response.WriteAsync(msg);
    });
});

app.UseDefaultFiles();   // serves index.html for "/"
app.UseStaticFiles();

app.MapControllers();

app.Run();
