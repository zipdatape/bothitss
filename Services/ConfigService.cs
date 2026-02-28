using System.Text.Json;
using NotificadorBajasHitssApp.Config;

namespace NotificadorBajasHitssApp.Services;

public static class ConfigService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static string GetConfigPath()
    {
        var dir = AppContext.BaseDirectory;
        return Path.Combine(dir, "config.json");
    }

    public static AppConfig Load()
    {
        var path = GetConfigPath();
        if (!File.Exists(path))
            return new AppConfig();
        try
        {
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
        }
        catch
        {
            return new AppConfig();
        }
    }

    public static void Save(AppConfig config)
    {
        var path = GetConfigPath();
        var json = JsonSerializer.Serialize(config, JsonOptions);
        File.WriteAllText(path, json);
    }
}
