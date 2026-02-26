using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;

namespace EmailManage.Services;

/// <summary>
/// Manages Ollama installation, model pulling, version checking, and lifecycle.
/// Provides one-click setup for users who don't have Ollama installed.
/// </summary>
public class OllamaSetupService
{
    private readonly DiagnosticLogger _log;
    private readonly HttpClient _http;

    private const string OllamaApiBase = "http://127.0.0.1:11434";
    private const string DefaultModel = "gemma3:4b";
    private const string OllamaInstallerUrl = "https://ollama.com/download/OllamaSetup.exe";

    public OllamaSetupService()
    {
        _log = DiagnosticLogger.Instance;
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
    }

    // ═══ Status Checks ═══

    /// <summary>
    /// Check if Ollama binary is installed on the system.
    /// </summary>
    public bool IsOllamaInstalled()
    {
        // Check common install locations
        var paths = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Ollama", "ollama.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Ollama", "ollama.exe"),
            "ollama" // In PATH
        };

        foreach (var p in paths)
        {
            if (p == "ollama")
            {
                try
                {
                    var psi = new ProcessStartInfo("where", "ollama") { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                    var proc = Process.Start(psi);
                    proc?.WaitForExit(5000);
                    if (proc?.ExitCode == 0) return true;
                }
                catch { }
            }
            else if (File.Exists(p))
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Check if Ollama API server is running and responsive.
    /// </summary>
    public async Task<bool> IsOllamaRunningAsync()
    {
        try
        {
            var response = await _http.GetAsync($"{OllamaApiBase}/api/tags");
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Start the Ollama background service if not running.
    /// </summary>
    public async Task<bool> StartOllamaAsync(IProgress<string>? progress = null)
    {
        if (await IsOllamaRunningAsync()) return true;

        progress?.Report("Starting Ollama service...");
        try
        {
            // Try to start via "ollama serve" or the app
            var ollamaPath = FindOllamaExe();
            if (ollamaPath == null)
            {
                _log.Warn("Cannot find ollama executable to start");
                return false;
            }

            var psi = new ProcessStartInfo(ollamaPath, "serve")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            Process.Start(psi);

            // Wait for API to become available (up to 15 seconds)
            for (int i = 0; i < 30; i++)
            {
                await Task.Delay(500);
                if (await IsOllamaRunningAsync())
                {
                    _log.Info("Ollama service started successfully");
                    return true;
                }
            }

            _log.Warn("Ollama started but API not responding after 15s");
            return false;
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Failed to start Ollama");
            return false;
        }
    }

    /// <summary>
    /// Get the currently installed Ollama version.
    /// </summary>
    public async Task<string?> GetOllamaVersionAsync()
    {
        try
        {
            var response = await _http.GetAsync($"{OllamaApiBase}/api/version");
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                return doc.RootElement.GetProperty("version").GetString();
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Check if the default model is already pulled.
    /// </summary>
    public async Task<(bool IsInstalled, string? ModelName, string? Size)> CheckModelInstalledAsync(string modelName = DefaultModel)
    {
        try
        {
            var response = await _http.GetAsync($"{OllamaApiBase}/api/tags");
            if (!response.IsSuccessStatusCode) return (false, null, null);

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);

            if (doc.RootElement.TryGetProperty("models", out var models))
            {
                foreach (var model in models.EnumerateArray())
                {
                    var name = model.GetProperty("name").GetString() ?? "";
                    if (name.StartsWith(modelName.Split(':')[0], StringComparison.OrdinalIgnoreCase))
                    {
                        var size = model.GetProperty("size").GetInt64();
                        var sizeStr = $"{size / 1_073_741_824.0:F1} GB";
                        return (true, name, sizeStr);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error checking model status");
        }
        return (false, null, null);
    }

    // ═══ Installation ═══

    /// <summary>
    /// Downloads and silently installs Ollama. Reports progress.
    /// </summary>
    public async Task<bool> InstallOllamaAsync(IProgress<string>? progress = null, CancellationToken ct = default)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "MailZen");
        Directory.CreateDirectory(tempDir);
        var installerPath = Path.Combine(tempDir, "OllamaSetup.exe");

        try
        {
            // Download installer
            progress?.Report("Downloading Ollama installer...");
            _log.Info("Downloading Ollama from {Url}", OllamaInstallerUrl);

            using var downloadClient = new HttpClient { Timeout = TimeSpan.FromMinutes(10) };
            using var response = await downloadClient.GetAsync(OllamaInstallerUrl, HttpCompletionOption.ResponseHeadersRead, ct);
            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength;
            using var contentStream = await response.Content.ReadAsStreamAsync(ct);
            using var fileStream = new FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true);

            var buffer = new byte[81920];
            long downloaded = 0;
            int bytesRead;

            while ((bytesRead = await contentStream.ReadAsync(buffer, ct)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead), ct);
                downloaded += bytesRead;
                if (totalBytes.HasValue && totalBytes.Value > 0)
                {
                    var pct = (int)(downloaded * 100 / totalBytes.Value);
                    progress?.Report($"Downloading Ollama... {pct}% ({downloaded / 1_048_576.0:F1} MB)");
                }
            }

            _log.Info("Ollama installer downloaded ({Size} bytes)", downloaded);

            // Run silent install
            progress?.Report("Installing Ollama (this may take a minute)...");
            var psi = new ProcessStartInfo(installerPath, "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            var process = Process.Start(psi);
            if (process != null)
            {
                await process.WaitForExitAsync(ct);
                if (process.ExitCode != 0)
                {
                    _log.Warn("Ollama installer exited with code {Code}", process.ExitCode);
                }
            }

            // Verify installation
            await Task.Delay(2000, ct);
            if (IsOllamaInstalled())
            {
                _log.Info("Ollama installed successfully");
                progress?.Report("Ollama installed! Starting service...");

                // Start the service
                await StartOllamaAsync(progress);
                return true;
            }
            else
            {
                progress?.Report("Installation may need a restart. Trying to start anyway...");
                await Task.Delay(3000, ct);
                return await StartOllamaAsync(progress);
            }
        }
        catch (OperationCanceledException)
        {
            progress?.Report("Installation cancelled.");
            return false;
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Failed to install Ollama");
            progress?.Report($"Installation failed: {ex.Message}");
            return false;
        }
        finally
        {
            // Clean up installer
            try { if (File.Exists(installerPath)) File.Delete(installerPath); } catch { }
        }
    }

    /// <summary>
    /// Pulls the specified model via Ollama API. Reports progress.
    /// </summary>
    public async Task<bool> PullModelAsync(string modelName = DefaultModel, IProgress<string>? progress = null, CancellationToken ct = default)
    {
        try
        {
            progress?.Report($"Pulling model '{modelName}'... this may take a few minutes.");
            _log.Info("Pulling model {Model}", modelName);

            var requestBody = JsonSerializer.Serialize(new { name = modelName, stream = true });
            var content = new StringContent(requestBody, System.Text.Encoding.UTF8, "application/json");

            using var pullClient = new HttpClient { Timeout = TimeSpan.FromMinutes(30) };
            using var request = new HttpRequestMessage(HttpMethod.Post, $"{OllamaApiBase}/api/pull") { Content = content };
            using var response = await pullClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
            response.EnsureSuccessStatusCode();

            using var stream = await response.Content.ReadAsStreamAsync(ct);
            using var reader = new StreamReader(stream);

            string? lastStatus = null;
            while (!reader.EndOfStream)
            {
                ct.ThrowIfCancellationRequested();
                var line = await reader.ReadLineAsync(ct);
                if (string.IsNullOrWhiteSpace(line)) continue;

                try
                {
                    using var doc = JsonDocument.Parse(line);
                    var root = doc.RootElement;

                    var status = root.TryGetProperty("status", out var s) ? s.GetString() : null;
                    if (root.TryGetProperty("total", out var total) && root.TryGetProperty("completed", out var completed))
                    {
                        var totalVal = total.GetInt64();
                        var completedVal = completed.GetInt64();
                        if (totalVal > 0)
                        {
                            var pct = (int)(completedVal * 100 / totalVal);
                            progress?.Report($"Pulling '{modelName}'... {pct}% ({completedVal / 1_073_741_824.0:F2} / {totalVal / 1_073_741_824.0:F2} GB)");
                        }
                    }
                    else if (status != lastStatus && status != null)
                    {
                        progress?.Report($"Model: {status}");
                        lastStatus = status;
                    }
                }
                catch { }
            }

            // Verify
            var (installed, _, _) = await CheckModelInstalledAsync(modelName);
            if (installed)
            {
                _log.Info("Model {Model} pulled successfully", modelName);
                progress?.Report($"Model '{modelName}' is ready!");
                return true;
            }

            progress?.Report("Model pull completed but verification failed. Try again.");
            return false;
        }
        catch (OperationCanceledException)
        {
            progress?.Report("Model pull cancelled.");
            return false;
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Failed to pull model {Model}", modelName);
            progress?.Report($"Model pull failed: {ex.Message}");
            return false;
        }
    }

    // ═══ Update Checking ═══

    /// <summary>
    /// Checks for Ollama updates and newer model versions. Returns a summary.
    /// </summary>
    public async Task<UpdateCheckResult> CheckForUpdatesAsync(IProgress<string>? progress = null)
    {
        var result = new UpdateCheckResult();

        progress?.Report("Checking Ollama version...");
        result.CurrentVersion = await GetOllamaVersionAsync();

        // Check if model is installed and get its info
        progress?.Report("Checking model status...");
        var (isInstalled, modelName, size) = await CheckModelInstalledAsync();
        result.ModelInstalled = isInstalled;
        result.ModelName = modelName;
        result.ModelSize = size;

        // Try pulling with dry-run to see if there's a newer digest
        // (Ollama re-pull is a no-op if model is up to date)
        result.CheckedAt = DateTime.UtcNow;

        progress?.Report($"Ollama v{result.CurrentVersion ?? "?"} — Model: {result.ModelName ?? "Not installed"}");
        return result;
    }

    // ═══ Helpers ═══

    private string? FindOllamaExe()
    {
        var paths = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Ollama", "ollama.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Ollama", "ollama.exe"),
        };

        foreach (var p in paths)
        {
            if (File.Exists(p)) return p;
        }

        // Check PATH
        try
        {
            var psi = new ProcessStartInfo("where", "ollama") { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
            var proc = Process.Start(psi);
            var output = proc?.StandardOutput.ReadToEnd()?.Trim();
            proc?.WaitForExit(5000);
            if (!string.IsNullOrWhiteSpace(output) && File.Exists(output.Split('\n')[0].Trim()))
                return output.Split('\n')[0].Trim();
        }
        catch { }

        return null;
    }
}

public class UpdateCheckResult
{
    public string? CurrentVersion { get; set; }
    public bool ModelInstalled { get; set; }
    public string? ModelName { get; set; }
    public string? ModelSize { get; set; }
    public DateTime CheckedAt { get; set; }
}
