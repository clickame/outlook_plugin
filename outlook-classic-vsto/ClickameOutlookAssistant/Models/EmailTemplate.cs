using System;
using Newtonsoft.Json;

namespace ClickameOutlookAssistant.Models
{
    /// <summary>
    /// Una plantilla de correu predefinida.
    /// El cos pot ser HTML o text pla; s'aplica segons el format del MailItem actiu.
    /// </summary>
    public class EmailTemplate
    {
        /// <summary>Identificador estable per editar/eliminar (es genera si falta).</summary>
        [JsonProperty("id")]
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        /// <summary>Nom visible a la llista de plantilles.</summary>
        [JsonProperty("nom")]
        public string Nom { get; set; } = "";

        /// <summary>Assumpte opcional. Si està buit, no es toca l'assumpte del correu.</summary>
        [JsonProperty("assumpte")]
        public string Assumpte { get; set; } = "";

        /// <summary>Cos de la plantilla. Pot ser HTML o text pla.</summary>
        [JsonProperty("cos")]
        public string Cos { get; set; } = "";

        /// <summary>Indica si <see cref="Cos"/> conté HTML.</summary>
        [JsonProperty("es_html")]
        public bool EsHtml { get; set; } = true;

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Nom) ? "(sense nom)" : Nom;
        }
    }
}
