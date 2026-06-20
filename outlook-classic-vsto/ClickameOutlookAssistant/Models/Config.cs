using System.Collections.Generic;
using Newtonsoft.Json;

namespace ClickameOutlookAssistant.Models
{
    /// <summary>
    /// Modes en què s'afegeix el BCC automàtic.
    /// </summary>
    public static class BccMode
    {
        /// <summary>Afegir el BCC en obrir/crear un correu nou.</summary>
        public const string OnNewMail = "on_new_mail";

        /// <summary>Afegir el BCC just abans d'enviar (ItemSend).</summary>
        public const string OnSend = "on_send";
    }

    /// <summary>
    /// Configuració local de l'add-in. Es serialitza a
    /// %APPDATA%\ClickameOutlookAssistant\config.json
    /// </summary>
    public class Config
    {
        /// <summary>Adreça que s'afegirà com a CCO/BCC.</summary>
        [JsonProperty("email_bcc")]
        public string EmailBcc { get; set; } = "";

        /// <summary>Activa o desactiva el BCC automàtic.</summary>
        [JsonProperty("activar_bcc_auto")]
        public bool ActivarBccAuto { get; set; } = false;

        /// <summary>Mode d'aplicació del BCC: <see cref="BccMode.OnNewMail"/> o <see cref="BccMode.OnSend"/>.</summary>
        [JsonProperty("mode_bcc")]
        public string ModeBcc { get; set; } = BccMode.OnSend;

        /// <summary>Llista de plantilles disponibles.</summary>
        [JsonProperty("plantilles")]
        public List<EmailTemplate> Plantilles { get; set; } = new List<EmailTemplate>();

        public static Config CreateDefault()
        {
            return new Config
            {
                EmailBcc = "",
                ActivarBccAuto = false,
                ModeBcc = BccMode.OnSend,
                Plantilles = new List<EmailTemplate>
                {
                    new EmailTemplate
                    {
                        Nom = "Salutació inicial",
                        Assumpte = "Gràcies pel teu interès",
                        EsHtml = true,
                        Cos = "<p>Hola,</p><p>Gràcies per contactar amb nosaltres. " +
                              "En breu et donarem resposta.</p><p>Salutacions,<br/>L'equip de Clickame</p>"
                    }
                }
            };
        }
    }
}
