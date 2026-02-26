using System.Net.Http;
using System.Text;
using System.Text.Json;

namespace EmailManage.Services;

/// <summary>
/// HTTP client for Ollama's /api/chat endpoint.
/// Sends email metadata, gets structured junk/not-junk classification.
/// </summary>
public class OllamaClient
{
    private readonly DiagnosticLogger _log;
    private readonly HttpClient _http;
    private readonly string _modelName;

    private const string OllamaApiBase = "http://127.0.0.1:11434";

    public OllamaClient(string modelName = "gemma3:4b")
    {
        _log = DiagnosticLogger.Instance;
        _modelName = modelName;
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(120) }; // CPU inference can be slow
    }

    /// <summary>
    /// Classify a single email as JUNK or NOT_JUNK.
    /// </summary>
    public async Task<AiClassification> ClassifyEmailAsync(
        string senderEmail, string subject, string bodyPreview, bool hasUnsubscribe,
        CancellationToken ct = default)
    {
        var result = new AiClassification { SenderEmail = senderEmail, Subject = subject };
        var sw = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            var prompt = BuildPrompt(senderEmail, subject, bodyPreview, hasUnsubscribe);

            var requestBody = new
            {
                model = _modelName,
                messages = new[]
                {
                    new { role = "system", content = "You are an email classifier. Respond ONLY with valid JSON. No markdown, no explanation outside the JSON." },
                    new { role = "user", content = prompt }
                },
                stream = false,
                options = new
                {
                    temperature = 0.1,     // Low temperature for deterministic classification
                    num_predict = 150      // Short response
                }
            };

            var json = JsonSerializer.Serialize(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _http.PostAsync($"{OllamaApiBase}/api/chat", content, ct);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(responseJson);

            // Extract the assistant's message content
            var messageContent = doc.RootElement
                .GetProperty("message")
                .GetProperty("content")
                .GetString() ?? "";

            // Parse the classification JSON from the response
            result = ParseClassification(messageContent, senderEmail, subject);
            result.LatencyMs = sw.ElapsedMilliseconds;

            _log.Debug("AI classified [{Sender}] '{Subject}' as {Class} ({Conf:P0}) in {Ms}ms",
                senderEmail, subject.Length > 40 ? subject[..40] + "..." : subject,
                result.Classification, result.Confidence, result.LatencyMs);
        }
        catch (TaskCanceledException)
        {
            result.Classification = "ERROR";
            result.Reason = "Request timed out";
            result.LatencyMs = sw.ElapsedMilliseconds;
        }
        catch (Exception ex)
        {
            _log.Error(ex, "AI classification failed for [{Sender}]", senderEmail);
            result.Classification = "ERROR";
            result.Reason = ex.Message;
            result.LatencyMs = sw.ElapsedMilliseconds;
        }

        return result;
    }

    private static string BuildPrompt(string senderEmail, string subject, string bodyPreview, bool hasUnsubscribe)
    {
        // Truncate body preview to 200 chars
        if (bodyPreview.Length > 200)
            bodyPreview = bodyPreview[..200] + "...";

        // Clean up body preview (remove excessive whitespace)
        bodyPreview = System.Text.RegularExpressions.Regex.Replace(bodyPreview, @"\s+", " ").Trim();

        var unsubStr = hasUnsubscribe ? "yes" : "no";
        return
            "Classify this email as JUNK or NOT_JUNK.\n\n" +
            "JUNK = marketing, newsletters, promotions, spam, automated notifications you'd typically delete, social media alerts, retail offers, surveys, unsubscribe-able bulk mail.\n" +
            "NOT_JUNK = personal correspondence, important business emails, account security alerts, order confirmations for recent purchases, appointment reminders, direct replies to your emails.\n\n" +
            $"Sender: {senderEmail}\n" +
            $"Subject: {subject}\n" +
            $"Body Preview: {bodyPreview}\n" +
            $"Has Unsubscribe Link: {unsubStr}\n\n" +
            "Reply with JSON only: {\"classification\": \"JUNK\" or \"NOT_JUNK\", \"confidence\": 0.0 to 1.0, \"reason\": \"brief reason\"}";
    }

    private AiClassification ParseClassification(string raw, string senderEmail, string subject)
    {
        var result = new AiClassification
        {
            SenderEmail = senderEmail,
            Subject = subject,
            RawResponse = raw
        };

        try
        {
            // Try to extract JSON from the response (it might have markdown wrapping)
            var jsonStr = raw.Trim();

            // Strip markdown code blocks if present
            if (jsonStr.Contains("```"))
            {
                var start = jsonStr.IndexOf('{');
                var end = jsonStr.LastIndexOf('}');
                if (start >= 0 && end > start)
                    jsonStr = jsonStr[start..(end + 1)];
            }

            // Find the JSON object
            var jsonStart = jsonStr.IndexOf('{');
            var jsonEnd = jsonStr.LastIndexOf('}');
            if (jsonStart >= 0 && jsonEnd > jsonStart)
                jsonStr = jsonStr[jsonStart..(jsonEnd + 1)];

            using var doc = JsonDocument.Parse(jsonStr);
            var root = doc.RootElement;

            if (root.TryGetProperty("classification", out var cls))
                result.Classification = cls.GetString()?.ToUpperInvariant() ?? "UNKNOWN";

            if (root.TryGetProperty("confidence", out var conf))
            {
                if (conf.ValueKind == JsonValueKind.Number)
                    result.Confidence = conf.GetDouble();
                else if (conf.ValueKind == JsonValueKind.String && double.TryParse(conf.GetString(), out var d))
                    result.Confidence = d;
            }

            if (root.TryGetProperty("reason", out var reason))
                result.Reason = reason.GetString() ?? "";
        }
        catch (Exception ex)
        {
            _log.Warn("Failed to parse AI response: {Raw} — {Error}", raw.Length > 100 ? raw[..100] : raw, ex.Message);
            result.Classification = "PARSE_ERROR";
            var preview = raw.Length > 80 ? raw[..80] : raw;
            result.Reason = $"Could not parse: {preview}";
        }

        return result;
    }
}

/// <summary>
/// Result of AI classification for one email.
/// </summary>
public class AiClassification
{
    public string SenderEmail { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Classification { get; set; } = "UNKNOWN"; // JUNK, NOT_JUNK, ERROR, PARSE_ERROR
    public double Confidence { get; set; }
    public string Reason { get; set; } = string.Empty;
    public long LatencyMs { get; set; }
    public string RawResponse { get; set; } = string.Empty;

    public bool IsJunk => Classification == "JUNK";
}
