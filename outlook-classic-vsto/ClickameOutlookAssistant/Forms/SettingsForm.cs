using System;
using System.Drawing;
using System.Windows.Forms;
using ClickameOutlookAssistant.Models;

namespace ClickameOutlookAssistant.Forms
{
    /// <summary>
    /// Formulari de configuració: adreça CCO, activar/desactivar CCO automàtic i mode.
    /// També dona accés a la gestió de plantilles.
    /// </summary>
    public class SettingsForm : Form
    {
        private TextBox _txtBcc;
        private CheckBox _chkAuto;
        private ComboBox _cmbMode;
        private Button _btnTemplates;
        private Button _btnOk;
        private Button _btnCancel;
        private LinkLabel _lnkConfig;

        public SettingsForm()
        {
            BuildUi();
            LoadValues();
        }

        private void BuildUi()
        {
            Text = "Clickame Outlook Assistant — Configuració";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(460, 250);
            Font = new Font("Segoe UI", 9F);

            var lblBcc = new Label { Text = "Adreça CCO (BCC):", Left = 16, Top = 20, Width = 140 };
            _txtBcc = new TextBox { Left = 160, Top = 17, Width = 280 };

            _chkAuto = new CheckBox { Text = "Activar CCO automàtic", Left = 160, Top = 52, Width = 280 };

            var lblMode = new Label { Text = "Mode:", Left = 16, Top = 84, Width = 140 };
            _cmbMode = new ComboBox { Left = 160, Top = 81, Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbMode.Items.Add(new ModeItem(BccMode.OnSend, "En enviar (recomanat) — on_send"));
            _cmbMode.Items.Add(new ModeItem(BccMode.OnNewMail, "En obrir correu nou — on_new_mail"));

            _btnTemplates = new Button { Text = "Gestionar plantilles…", Left = 160, Top = 120, Width = 180, Height = 30 };
            _btnTemplates.Click += (s, e) =>
            {
                using (var f = new TemplatesForm()) f.ShowDialog(this);
            };

            _lnkConfig = new LinkLabel
            {
                Text = "Obrir carpeta de configuració/logs",
                Left = 16,
                Top = 165,
                Width = 280
            };
            _lnkConfig.LinkClicked += (s, e) =>
            {
                try { System.Diagnostics.Process.Start(ConfigService.Instance.ConfigFolder); }
                catch (Exception ex) { Logger.Warn("No s'ha pogut obrir la carpeta: " + ex.Message); }
            };

            _btnOk = new Button { Text = "Desar", Left = 270, Top = 205, Width = 80, DialogResult = DialogResult.OK };
            _btnOk.Click += BtnOk_Click;
            _btnCancel = new Button { Text = "Cancel·lar", Left = 360, Top = 205, Width = 80, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[]
            {
                lblBcc, _txtBcc, _chkAuto, lblMode, _cmbMode,
                _btnTemplates, _lnkConfig, _btnOk, _btnCancel
            });

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }

        private void LoadValues()
        {
            var cfg = ConfigService.Instance.Current;
            _txtBcc.Text = cfg.EmailBcc;
            _chkAuto.Checked = cfg.ActivarBccAuto;

            foreach (ModeItem item in _cmbMode.Items)
            {
                if (item.Value == cfg.ModeBcc) { _cmbMode.SelectedItem = item; break; }
            }
            if (_cmbMode.SelectedItem == null && _cmbMode.Items.Count > 0)
                _cmbMode.SelectedIndex = 0;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            var bcc = (_txtBcc.Text ?? "").Trim();
            if (_chkAuto.Checked && string.IsNullOrWhiteSpace(bcc))
            {
                MessageBox.Show("Has d'indicar una adreça CCO per activar el CCO automàtic.",
                    "Clickame", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            var cfg = ConfigService.Instance.Current;
            cfg.EmailBcc = bcc;
            cfg.ActivarBccAuto = _chkAuto.Checked;
            cfg.ModeBcc = (_cmbMode.SelectedItem as ModeItem)?.Value ?? BccMode.OnSend;

            try
            {
                ConfigService.Instance.Save(cfg);
            }
            catch
            {
                MessageBox.Show("No s'ha pogut desar la configuració. Mira el log.",
                    "Clickame", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
            }
        }

        private class ModeItem
        {
            public string Value { get; }
            private readonly string _label;
            public ModeItem(string value, string label) { Value = value; _label = label; }
            public override string ToString() => _label;
        }
    }
}
