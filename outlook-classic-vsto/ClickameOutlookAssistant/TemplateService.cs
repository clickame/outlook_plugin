using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using ClickameOutlookAssistant.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClickameOutlookAssistant
{
    /// <summary>
    /// Gestió de plantilles (CRUD sobre la configuració) i inserció dins d'un MailItem.
    /// </summary>
    public class TemplateService
    {
        private static readonly Lazy<TemplateService> _instance = new Lazy<TemplateService>(() => new TemplateService());
        public static TemplateService Instance => _instance.Value;

        private Config Cfg => ConfigService.Instance.Current;

        public IReadOnlyList<EmailTemplate> GetAll() => Cfg.Plantilles;

        public EmailTemplate GetById(string id) => Cfg.Plantilles.FirstOrDefault(t => t.Id == id);

        public void AddOrUpdate(EmailTemplate template)
        {
            if (template == null) throw new ArgumentNullException(nameof(template));
            if (string.IsNullOrEmpty(template.Id)) template.Id = Guid.NewGuid().ToString("N");

            var existing = GetById(template.Id);
            if (existing == null)
                Cfg.Plantilles.Add(template);
            else
            {
                existing.Nom = template.Nom;
                existing.Assumpte = template.Assumpte;
                existing.Cos = template.Cos;
                existing.EsHtml = template.EsHtml;
            }
            ConfigService.Instance.SaveCurrent();
            Logger.Info($"Plantilla desada: {template.Nom}");
        }

        public void Delete(string id)
        {
            var t = GetById(id);
            if (t != null)
            {
                Cfg.Plantilles.Remove(t);
                ConfigService.Instance.SaveCurrent();
                Logger.Info($"Plantilla eliminada: {t.Nom}");
            }
        }

        /// <summary>
        /// Insereix una plantilla al MailItem actiu.
        /// - Si el correu és HTML, insereix a HTMLBody (al cursor si es pot, si no al final).
        /// - Si és text pla, insereix a Body.
        /// - Si la plantilla té assumpte i el correu encara no en té, l'omple.
        /// </summary>
        public void InsertIntoMailItem(Outlook.MailItem mail, EmailTemplate template)
        {
            if (mail == null) throw new ArgumentNullException(nameof(mail));
            if (template == null) throw new ArgumentNullException(nameof(template));

            try
            {
                // Assumpte opcional: només si la plantilla en porta i el correu no en té.
                if (!string.IsNullOrWhiteSpace(template.Assumpte) &&
                    string.IsNullOrWhiteSpace(mail.Subject))
                {
                    mail.Subject = template.Assumpte;
                }

                bool insertedAtCursor = TryInsertAtCursor(mail, template);
                if (!insertedAtCursor)
                {
                    AppendToBody(mail, template);
                }

                Logger.Info($"Plantilla inserida: {template.Nom}");
            }
            catch (Exception ex)
            {
                Logger.Error("Error inserint la plantilla; es prova append simple.", ex);
                try { AppendToBody(mail, template); } catch (Exception ex2) { Logger.Error("Append també ha fallat.", ex2); throw; }
            }
        }

        /// <summary>
        /// Intenta inserir al punt del cursor mitjançant la selecció de l'inspector Word.
        /// Retorna false si no hi ha selecció accessible (s'haurà de fer append).
        /// </summary>
        private bool TryInsertAtCursor(Outlook.MailItem mail, EmailTemplate template)
        {
            try
            {
                var inspector = mail.GetInspector;
                // L'editor de Word està disponible quan el cos és HTML/RTF i WordEditor és actiu.
                if (inspector?.EditorType == Outlook.OlEditorType.olEditorWord &&
                    inspector.WordEditor is Microsoft.Office.Interop.Word.Document doc)
                {
                    var selection = doc.Application?.Selection;
                    if (selection == null) return false;

                    if (template.EsHtml)
                    {
                        // Word importa HTML al punt del cursor preservant el format via InsertFile,
                        // a partir d'un fitxer .html temporal local (no surt res del PC).
                        string tmp = Path.Combine(
                            Path.GetTempPath(),
                            "clickame_tpl_" + Guid.NewGuid().ToString("N") + ".html");
                        try
                        {
                            File.WriteAllText(
                                tmp,
                                "<html><head><meta charset=\"utf-8\"></head><body>" +
                                    template.Cos + "</body></html>",
                                Encoding.UTF8);
                            selection.InsertFile(tmp);
                        }
                        finally
                        {
                            try { if (File.Exists(tmp)) File.Delete(tmp); }
                            catch { /* el fitxer temporal es netejarà igualment pel sistema */ }
                        }
                    }
                    else
                    {
                        selection.TypeText(template.Cos);
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn("No s'ha pogut inserir al cursor: " + ex.Message);
            }
            return false;
        }

        /// <summary>Treu etiquetes HTML per al fallback de text dins WordML.</summary>
        private static string StripTags(string html)
        {
            if (string.IsNullOrEmpty(html)) return "";
            return System.Text.RegularExpressions.Regex.Replace(html, "<.*?>", string.Empty);
        }

        /// <summary>Afegeix la plantilla al final del cos respectant HTML vs text pla.</summary>
        private void AppendToBody(Outlook.MailItem mail, EmailTemplate template)
        {
            bool mailIsHtml = mail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML;

            if (mailIsHtml)
            {
                var existing = mail.HTMLBody ?? "";
                var fragment = template.EsHtml ? template.Cos : "<pre>" + HttpUtility.HtmlEncode(template.Cos) + "</pre>";

                if (existing.IndexOf("</body>", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    int idx = existing.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                    mail.HTMLBody = existing.Substring(0, idx) + fragment + existing.Substring(idx);
                }
                else
                {
                    mail.HTMLBody = existing + fragment;
                }
            }
            else
            {
                // Cos en text pla: afegim text. Si la plantilla és HTML, en treiem etiquetes.
                var text = template.EsHtml ? StripTags(template.Cos) : template.Cos;
                mail.Body = (mail.Body ?? "") + Environment.NewLine + text;
            }
        }
    }
}
