using System.Drawing;
using System.Windows.Forms;

namespace AdAgencyManager
{
    public partial class ClientsForm : Form
    {
        private DataGridView dataGridViewClients;

        public ClientsForm()
        {
            InitializeComponent();
            SetupForm();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(600, 450);
            this.Text = "Управление клиентами";
            this.ResumeLayout(false);
        }

        private void SetupForm()
        {
            dataGridViewClients = new DataGridView
            {
                Location = new Point(20, 20),
                Size = new Size(560, 400),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            dataGridViewClients.Columns.Add("ClientID", "ID");
            dataGridViewClients.Columns.Add("Name", "Название");
            dataGridViewClients.Columns.Add("Email", "Email");
            dataGridViewClients.Columns.Add("Phone", "Телефон");

            this.Controls.Add(dataGridViewClients);
        }
    }
}