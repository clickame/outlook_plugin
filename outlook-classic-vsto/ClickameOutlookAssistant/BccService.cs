using System;
using System.Linq;
using ClickameOutlookAssistant.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClickameOutlookAssistant
{
    /// <summary>
    /// Lògica per afegir el BCC/CCO automàtic a un MailItem evitant duplicats.
    /// </summary>
    public class BccService
    {
        private static readonly Lazy<BccService> _instance = new Lazy<BccService>(() => new BccService());
        public static BccService Instance => _instance.Value;

        /// <summary>
        /// Afegeix el BCC configurat al correu si està activat i encara no hi és.
        /// Retorna true si s'ha afegit (o ja hi era correctament), false si està desactivat o sense adreça.
        /// </summary>
        public bool EnsureBcc(Outlook.MailItem mail)
        {
            if (mail == null) return false;

            var cfg = ConfigService.Instance.Current;
            if (!cfg.ActivarBccAuto) return false;

            var target = (cfg.EmailBcc ?? "").Trim();
            if (string.IsNullOrWhiteSpace(target))
            {
                Logger.Warn("BCC automàtic activat però sense adreça configurada.");
                return false;
            }

            try
            {
                if (BccAlreadyPresent(mail, target))
                {
                    Logger.Info("El BCC ja existeix; no es duplica.");
                    return true;
                }

                // Afegir destinatari de tipus BCC.
                var recipient = mail.Recipients.Add(target);
                recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;
                bool resolved = recipient.Resolve();
                Logger.Info($"BCC afegit: {target} (resolt={resolved})");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("No s'ha pogut afegir el BCC.", ex);
                return false;
            }
        }

        /// <summary>Comprova si l'adreça ja és present com a destinatari BCC.</summary>
        private bool BccAlreadyPresent(Outlook.MailItem mail, string target)
        {
            try
            {
                foreach (Outlook.Recipient r in mail.Recipients)
                {
                    if (r.Type != (int)Outlook.OlMailRecipientType.olBCC) continue;

                    string addr = GetSmtpAddress(r);
                    if (!string.IsNullOrEmpty(addr) &&
                        addr.Equals(target, StringComparison.OrdinalIgnoreCase))
                        return true;

                    // Comparació també pel nom mostrat per cobrir adreces no resoltes.
                    if (!string.IsNullOrEmpty(r.Name) &&
                        r.Name.Equals(target, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn("Error comprovant duplicats de BCC: " + ex.Message);
            }
            return false;
        }

        /// <summary>Obté l'adreça SMTP d'un destinatari, gestionant comptes Exchange.</summary>
        private string GetSmtpAddress(Outlook.Recipient recipient)
        {
            try
            {
                var ae = recipient.AddressEntry;
                if (ae == null) return recipient.Address;

                if (ae.Type == "EX")
                {
                    var exUser = ae.GetExchangeUser();
                    if (exUser != null && !string.IsNullOrEmpty(exUser.PrimarySmtpAddress))
                        return exUser.PrimarySmtpAddress;
                }
                return recipient.Address;
            }
            catch
            {
                return recipient.Address;
            }
        }
    }
}
