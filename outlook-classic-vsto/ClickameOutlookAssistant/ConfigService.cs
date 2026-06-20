using System;
using System.IO;
using ClickameOutlookAssistant.Models;
using Newtonsoft.Json;

namespace ClickameOutlookAssistant
{
    /// <summary>
    /// Llegeix i desa la configuració local a
    /// %APPDATA%\ClickameOutlookAssistant\config.json
    /// </summary>
    public class ConfigService
    {
        private static readonly Lazy<ConfigService> _instance = new Lazy<ConfigService>(() => new ConfigService());
        public static ConfigService Instance => _instance.Value;

        private Config _current;

        public string ConfigFolder =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ClickameOutlookAssistant");

        public string ConfigFilePath => Path.Combine(ConfigFolder, "config.json");

        /// <summary>Configuració carregada en memòria (es carrega de manera mandrosa).</summary>
        public Config Current
        {
            get
            {
                if (_current == null) Load();
                return _current;
            }
        }

        /// <summary>(Re)carrega la configuració des de disc. Si no existeix, en crea una per defecte.</summary>
        public Config Load()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var json = File.ReadAllText(ConfigFilePath);
                    _current = JsonConvert.DeserializeObject<Config>(json) ?? Config.CreateDefault();
                    Logger.Info($"Configuració carregada de {ConfigFilePath}");
                }
                else
                {
                    _current = Config.CreateDefault();
                    Save(_current);
                    Logger.Info("No s'ha trobat config.json; s'ha creat una configuració per defecte.");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error carregant la configuració; s'usa la per defecte.", ex);
                _current = Config.CreateDefault();
            }

            // Normalitza el mode per evitar valors invàlids.
            if (_current.ModeBcc != BccMode.OnNewMail && _current.ModeBcc != BccMode.OnSend)
                _current.ModeBcc = BccMode.OnSend;

            return _current;
        }

        /// <summary>Desa la configuració a disc i l'estableix com a actual.</summary>
        public void Save(Config config)
        {
            try
            {
                Directory.CreateDirectory(ConfigFolder);
                var json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(ConfigFilePath, json);
                _current = config;
                Logger.Info($"Configuració desada a {ConfigFilePath}");
            }
            catch (Exception ex)
            {
                Logger.Error("Error desant la configuració.", ex);
                throw;
            }
        }

        /// <summary>Desa la configuració actualment en memòria.</summary>
        public void SaveCurrent() => Save(Current);
    }
}
