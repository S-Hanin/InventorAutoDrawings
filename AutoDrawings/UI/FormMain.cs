using InvAddIn.Drawings;
using System;
using System.Windows.Forms;

namespace InvAddIn.UI {
    internal partial class FormMain : Form {
        public Drawings.DrawingMaker drawingMaker;
        private BindingSource binding = new BindingSource();

        public FormMain() {
            InitializeComponent();
        }

        public FormMain(Drawings.DrawingMaker maker) {
            InitializeComponent();
            drawingMaker = maker;
            binding.DataSource = drawingMaker.Parts;
            dgvParts.DataSource = binding;
        }

        private void BtnOk_Click(object sender, EventArgs e) {
            drawingMaker.MakeDrawing();
            this.Close();
        }

        private void cbFormat_SelectedIndexChanged(object sender, EventArgs e) {
            var cb = sender as ComboBox;
            foreach (DataGridViewRow item in dgvParts.SelectedRows) {
                ((ModelInfo)item.DataBoundItem).SheetFormat = cb.SelectedItem as string;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}
