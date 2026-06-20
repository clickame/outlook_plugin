using System;
using System.Collections.Generic;
using ClickameOutlookAssistant.Models;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClickameOutlookAssistant
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors _inspectors;

        // Mantenim referències vives als inspectors per evitar que el GC alliberi els handlers.
        private readonly List<Outlook.Inspector> _trackedInspectors = new List<Outlook.Inspector>();

        // Instància de la Ribbon perquè els callbacks puguin accedir a l'add-in.
        internal Ribbon RibbonInstance { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                Logger.Info("=== Clickame Outlook Assistant iniciant ===");
                ConfigService.Instance.Load();

                // BCC en obrir un correu nou (mode on_new_mail).
                _inspectors = this.Application.Inspectors;
                _inspectors.NewInspector += Inspectors_NewInspector;

                // BCC garantit just abans d'enviar (mode on_send).
                this.Application.ItemSend += Application_ItemSend;

                Logger.Info("Esdeveniments registrats correctament.");
            }
            catch (Exception ex)
            {
                Logger.Error("Error a l'arrencada de l'add-in.", ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (_inspectors != null) _inspectors.NewInspector -= Inspectors_NewInspector;
                this.Application.ItemSend -= Application_ItemSend;
                _trackedInspectors.Clear();
                Logger.Info("=== Add-in aturat ===");
            }
            catch (Exception ex)
            {
                Logger.Error("Error a l'aturada de l'add-in.", ex);
            }
        }

        /// <summary>Connecta la Ribbon personalitzada amb VSTO.</summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            RibbonInstance = new Ribbon();
            return RibbonInstance;
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            try
            {
                if (!(inspector.CurrentItem is Outlook.MailItem mail)) return;

                _trackedInspectors.Add(inspector);

                var cfg = ConfigService.Instance.Current;
                if (cfg.ActivarBccAuto && cfg.ModeBcc == BccMode.OnNewMail)
                {
                    // Només per a correus en composició (no rebuts).
                    if (mail.Sent == false)
                    {
                        BccService.Instance.EnsureBcc(mail);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error a NewInspector.", ex);
            }
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            try
            {
                if (!(item is Outlook.MailItem mail)) return;

                var cfg = ConfigService.Instance.Current;
                if (!cfg.ActivarBccAuto) return;

                // En mode on_send sempre garantim; en on_new_mail també revalidem per si l'usuari
                // ha esborrat el BCC manualment després d'obrir el correu.
                bool ok = BccService.Instance.EnsureBcc(mail);

                if (!ok && string.IsNullOrWhiteSpace((cfg.EmailBcc ?? "").Trim()))
                {
                    Logger.Warn("ItemSend: no s'ha pogut garantir el BCC (sense adreça configurada).");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error a ItemSend.", ex);
            }
        }

        /// <summary>Retorna el MailItem en composició a l'inspector actiu, o null.</summary>
        internal Outlook.MailItem GetActiveComposeMail()
        {
            try
            {
                var inspector = this.Application.ActiveInspector();
                if (inspector?.CurrentItem is Outlook.MailItem mail && mail.Sent == false)
                    return mail;
            }
            catch (Exception ex)
            {
                Logger.Warn("No s'ha pogut obtenir el correu actiu: " + ex.Message);
            }
            return null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
