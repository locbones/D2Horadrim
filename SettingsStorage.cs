using System.Runtime.InteropServices;
using System.Text.Json;

namespace D2_Horadrim
{
    /// <summary>
    /// Cross-platform settings: Windows uses Registry (HKCU\Software\D2_Horadrim),
    /// Linux (and other non-Windows) use a JSON file in the application directory.
    /// </summary>
    internal static class SettingsStorage
    {
        private const string RegistryKeyPath = @"Software\D2_Horadrim";
        private const string SettingsFileName = "D2Horadrim_settings.json";

        private static Dictionary<string, object>? _jsonSettings;
        private static readonly object _jsonLock = new object();

        public static bool IsWindows => RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

        private static string GetSettingsJsonPath()
        {
            string dir = AppContext.BaseDirectory ?? "";
            return Path.Combine(dir, SettingsFileName);
        }

        private static void EnsureJsonSettingsLoaded()
        {
            if (_jsonSettings != null) return;
            lock (_jsonLock)
            {
                if (_jsonSettings != null) return;
                string path = GetSettingsJsonPath();
                try
                {
                    if (File.Exists(path))
                    {
                        string json = File.ReadAllText(path);
                        _jsonSettings = JsonSerializer.Deserialize<Dictionary<string, object>>(json) ?? new Dictionary<string, object>();
                    }
                    else
                        _jsonSettings = new Dictionary<string, object>();
                }
                catch
                {
                    _jsonSettings = new Dictionary<string, object>();
                }
            }
        }

        private static void SaveJsonSettings()
        {
            if (_jsonSettings == null) return;
            lock (_jsonLock)
            {
                if (_jsonSettings == null) return;
                try
                {
                    string path = GetSettingsJsonPath();
                    var options = new JsonSerializerOptions { WriteIndented = true };
                    string json = JsonSerializer.Serialize(_jsonSettings, options);
                    File.WriteAllText(path, json);
                }
                catch { }
            }
        }

        private static T GetFromJsonValue<T>(object? o, T defaultValue)
        {
            if (o == null) return defaultValue;
            if (o is JsonElement je)
            {
                try
                {
                    if (typeof(T) == typeof(int)) return (T)(object)je.GetInt32();
                    if (typeof(T) == typeof(long)) return (T)(object)je.GetInt64();
                    if (typeof(T) == typeof(string)) return (T)(object)(je.GetString() ?? "");
                    if (typeof(T) == typeof(float)) return (T)(object)(float)je.GetDouble();
                    if (typeof(T) == typeof(double)) return (T)(object)je.GetDouble();
                }
                catch { }
            }
            try { return (T)Convert.ChangeType(o, typeof(T)); }
            catch { return defaultValue; }
        }

        public static T GetSetting<T>(string key, T defaultValue)
        {
            if (IsWindows)
            {
                try
                {
                    using (var keyObj = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                    {
                        object? val = keyObj?.GetValue(key, defaultValue);
                        if (val == null) return defaultValue;
                        if (val is T t) return t;
                        return (T)Convert.ChangeType(val, typeof(T));
                    }
                }
                catch { return defaultValue; }
            }
            EnsureJsonSettingsLoaded();
            lock (_jsonLock)
            {
                if (_jsonSettings!.TryGetValue(key, out object? o))
                    return GetFromJsonValue(o, defaultValue);
            }
            return defaultValue;
        }

        public static string[]? GetSettingStringArray(string key)
        {
            if (IsWindows)
            {
                try
                {
                    using (var keyObj = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                    {
                        object? val = keyObj?.GetValue(key);
                        if (val is string[] arr) return arr;
                        return null;
                    }
                }
                catch { return null; }
            }
            EnsureJsonSettingsLoaded();
            lock (_jsonLock)
            {
                if (!_jsonSettings!.TryGetValue(key, out object? o) || o == null) return null;
                if (o is JsonElement je && je.ValueKind == JsonValueKind.Array)
                {
                    var list = new List<string>();
                    foreach (var e in je.EnumerateArray())
                        list.Add(e.GetString() ?? "");
                    return list.ToArray();
                }
            }
            return null;
        }

        public static void SetSetting(string key, object? value)
        {
            if (IsWindows)
            {
                try
                {
                    using (var keyObj = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                    {
                        if (keyObj == null) return;
                        if (value == null)
                            keyObj.DeleteValue(key, false);
                        else
                            keyObj.SetValue(key, value);
                    }
                }
                catch { }
                return;
            }
            lock (_jsonLock)
            {
                EnsureJsonSettingsLoaded();
                if (value == null)
                    _jsonSettings!.Remove(key);
                else
                    _jsonSettings![key] = value;
                SaveJsonSettings();
            }
        }

        public static void SetSettingStringArray(string key, string[] value)
        {
            if (IsWindows)
            {
                try
                {
                    using (var keyObj = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                        keyObj?.SetValue(key, value, Microsoft.Win32.RegistryValueKind.MultiString);
                }
                catch { }
                return;
            }
            lock (_jsonLock)
            {
                EnsureJsonSettingsLoaded();
                _jsonSettings![key] = value;
                SaveJsonSettings();
            }
        }

        public static int? GetSettingOptionalInt(string key)
        {
            if (IsWindows)
            {
                try
                {
                    using (var keyObj = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                    {
                        object? val = keyObj?.GetValue(key);
                        if (val == null) return null;
                        return Convert.ToInt32(val);
                    }
                }
                catch { return null; }
            }
            EnsureJsonSettingsLoaded();
            lock (_jsonLock)
            {
                if (!_jsonSettings!.TryGetValue(key, out object? o) || o == null) return null;
                if (o is JsonElement je)
                {
                    try { return je.GetInt32(); }
                    catch { return null; }
                }
                try { return Convert.ToInt32(o); }
                catch { return null; }
            }
        }

        public static void DeleteValue(string key)
        {
            if (IsWindows)
            {
                try
                {
                    using (var keyObj = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                        keyObj?.DeleteValue(key, false);
                }
                catch { }
                return;
            }
            SetSetting(key, null);
        }
    }
}
