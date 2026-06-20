using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using ClickameOutlookAssistant.Forms;
using ClickameOutlookAssistant.Models;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClickameOutlookAssistant
{
    /// <summary>
    /// Ribbon personalitzada (IRibbonExtensibility) amb un botó de Configuració
    /// i un menú desplegable de Plantilles. Es mostra a la pestanya de composició de correu.
    /// </summary>
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            // Mostrem la pestanya només al compositor de correu (TabNewMailMessage).
            return GetResourceText("ClickameOutlookAssistant.Ribbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        /// <summary>Força el refresc dels controls dinàmics (la llista de plantilles).</summary>
        public void Invalidate()
        {
            try { _ribbon?.Invalidate(); } catch { /* ignore */ }
        }

        // ---- Callbacks de la Ribbon ----

        public void OnSettingsClick(Office.IRibbonControl control)
        {
            try
            {
                using (var form = new SettingsForm())
                {
                    form.ShowDialog();
                }
                Invalidate(); // Per si han canviat les plantilles.
            }
            catch (Exception ex)
            {
                Logger.Error("Error obrint Configuració.", ex);
                MessageBox.Show("No s'ha pogut obrir la configuració. Mira el log per detalls.",
                    "Clickame", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnManageTemplatesClick(Office.IRibbonControl control)
        {
            try
            {
                using (var form = new TemplatesForm())
                {
                    form.ShowDialog();
                }
                Invalidate();
            }
            catch (Exception ex)
            {
                Logger.Error("Error obrint Plantilles.", ex);
            }
        }

        // ---- Menú dinàmic de plantilles ----

        public string GetTemplatesMenuContent(Office.IRibbonControl control)
        {
            var ns = "http://schemas.microsoft.com/office/2009/07/customui";
            var menu = new XElement(XName.Get("menu", ns));

            var plantilles = ConfigService.Instance.Current.Plantilles ?? new List<EmailTemplate>();
            if (!plantilles.Any())
            {
                menu.Add(new XElement(XName.Get("button", ns),
                    new XAttribute("id", "noTemplates"),
                    new XAttribute("label", "(Cap plantilla. Obre Configuració)"),
                    new XAttribute("enabled", "false")));
            }
            else
            {
                int i = 0;
                foreach (var t in plantilles)
                {
                    menu.Add(new XElement(XName.Get("button", ns),
                        new XAttribute("id", "tpl_" + i),
                        new XAttribute("tag", t.Id),
                        new XAttribute("label", string.IsNullOrWhiteSpace(t.Nom) ? "(sense nom)" : t.Nom),
                        new XAttribute("onAction", "OnInsertTemplate")));
                    i++;
                }
            }

            menu.Add(new XElement(XName.Get("menuSeparator", ns), new XAttribute("id", "sep1")));
            menu.Add(new XElement(XName.Get("button", ns),
                new XAttribute("id", "manageTpl"),
                new XAttribute("label", "Gestionar plantilles…"),
                new XAttribute("onAction", "OnManageTemplatesClick")));

            return menu.ToString();
        }

        public void OnInsertTemplate(Office.IRibbonControl control)
        {
            try
            {
                var templateId = control.Tag;
                var template = TemplateService.Instance.GetById(templateId);
                if (template == null)
                {
                    Logger.Warn("Plantilla no trobada per id: " + templateId);
                    return;
                }

                Outlook.MailItem mail = Globals.ThisAddIn.GetActiveComposeMail();
                if (mail == null)
                {
                    MessageBox.Show("Obre o crea un correu en mode composició per inserir una plantilla.",
                        "Clickame", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                TemplateService.Instance.InsertIntoMailItem(mail, template);
            }
            catch (Exception ex)
            {
                Logger.Error("Error inserint plantilla des de la Ribbon.", ex);
            }
        }

        // ---- Utilitats ----

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var names = asm.GetManifestResourceNames();
            var match = names.FirstOrDefault(n => string.Equals(n, resourceName, StringComparison.OrdinalIgnoreCase));
            if (match == null) return null;

            using (var stream = asm.GetManifestResourceStream(match))
            using (var reader = new System.IO.StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
