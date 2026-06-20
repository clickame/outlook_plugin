using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClickameOutlookAssistant.Models;

namespace ClickameOutlookAssistant.Forms
{
    /// <summary>
    /// Gestió de plantilles: crear, editar i eliminar.
    /// La llista de l'esquerra mostra les plantilles; el panell dret n'edita els camps.
    /// </summary>
    public class TemplatesForm : Form
    {
        private ListBox _list;
        private TextBox _txtNom;
        private TextBox _txtAssumpte;
        private TextBox _txtCos;
        private CheckBox _chkHtml;
        private Button _btnNew;
        private Button _btnDelete;
        private Button _btnSave;
        private Button _btnClose;

        private EmailTemplate _selected;

        public TemplatesForm()
        {
            BuildUi();
            RefreshList();
        }

        private void BuildUi()
        {
            Text = "Clickame Outlook Assistant — Plantilles";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimumSize = new Size(720, 460);
            ClientSize = new Size(760, 480);
            Font = new Font("Segoe UI", 9F);

            _list = new ListBox { Left = 12, Top = 12, Width = 220, Height = 380, IntegralHeight = false };
            _list.SelectedIndexChanged += List_SelectedIndexChanged;

            _btnNew = new Button { Text = "Nova", Left = 12, Top = 400, Width = 105 };
            _btnNew.Click += (s, e) => SelectTemplate(null);
            _btnDelete = new Button { Text = "Eliminar", Left = 127, Top = 400, Width = 105 };
            _btnDelete.Click += BtnDelete_Click;

            var lblNom = new Label { Text = "Nom:", Left = 248, Top = 16, Width = 90 };
            _txtNom = new TextBox { Left = 344, Top = 13, Width = 400, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };

            var lblAssumpte = new Label { Text = "Assumpte (opc.):", Left = 248, Top = 48, Width = 90 };
            _txtAssumpte = new TextBox { Left = 344, Top = 45, Width = 400, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };

            _chkHtml = new CheckBox { Text = "El cos és HTML", Left = 344, Top = 78, Width = 200, Checked = true };

            var lblCos = new Label { Text = "Cos:", Left = 248, Top = 108, Width = 90 };
            _txtCos = new TextBox
            {
                Left = 344,
                Top = 108,
                Width = 400,
                Height = 284,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            _btnSave = new Button { Text = "Desar plantilla", Left = 524, Top = 400, Width = 110, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            _btnSave.Click += BtnSave_Click;
            _btnClose = new Button { Text = "Tancar", Left = 644, Top = 400, Width = 100, Anchor = AnchorStyles.Bottom | AnchorStyles.Right, DialogResult = DialogResult.OK };

            Controls.AddRange(new Control[]
            {
                _list, _btnNew, _btnDelete,
                lblNom, _txtNom, lblAssumpte, _txtAssumpte, _chkHtml, lblCos, _txtCos,
                _btnSave, _btnClose
            });

            CancelButton = _btnClose;
        }

        private void RefreshList()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (var t in TemplateService.Instance.GetAll())
                _list.Items.Add(t);
            _list.EndUpdate();
        }

        private void List_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_list.SelectedItem is EmailTemplate t)
                SelectTemplate(t);
        }

        private void SelectTemplate(EmailTemplate t)
        {
            _selected = t;
            _txtNom.Text = t?.Nom ?? "";
            _txtAssumpte.Text = t?.Assumpte ?? "";
            _txtCos.Text = t?.Cos ?? "";
            _chkHtml.Checked = t?.EsHtml ?? true;
            if (t == null) _list.ClearSelected();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            var nom = (_txtNom.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(nom))
            {
                MessageBox.Show("La plantilla necessita un nom.", "Clickame",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var template = _selected ?? new EmailTemplate();
            template.Nom = nom;
            template.Assumpte = (_txtAssumpte.Text ?? "").Trim();
            template.Cos = _txtCos.Text ?? "";
            template.EsHtml = _chkHtml.Checked;

            TemplateService.Instance.AddOrUpdate(template);
            _selected = template;
            RefreshList();

            // Reselecciona la plantilla desada.
            var idx = TemplateService.Instance.GetAll().ToList().FindIndex(x => x.Id == template.Id);
            if (idx >= 0) _list.SelectedIndex = idx;
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_selected == null)
            {
                MessageBox.Show("Selecciona una plantilla per eliminar.", "Clickame",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirm = MessageBox.Show($"Eliminar la plantilla \"{_selected.Nom}\"?",
                "Clickame", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            TemplateService.Instance.Delete(_selected.Id);
            SelectTemplate(null);
            RefreshList();
        }
    }
}
