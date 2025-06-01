using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace AdAgencyManager
{
    public partial class MediaPlatformsForm : Form
    {
        private DataGridView dataGridViewMedia;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnDelete;
        private Button btnExportExcel;
        private TextBox txtSearch;
        private Button btnSearch;

        private DataTable mediaTable;
        private SqlDataAdapter adapter;
        private SqlCommandBuilder commandBuilder;

        public MediaPlatformsForm()
        {
            InitializeComponent();
            SetupForm();
            LoadMediaPlatforms();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(700, 500);
            this.Text = "Управление медиаплатформами";
            this.ResumeLayout(false);
        }

        private void SetupForm()
        {
            // Поле поиска
            txtSearch = new TextBox
            {
                Location = new Point(20, 20),
                Size = new Size(200, 20),
                PlaceholderText = "Поиск по названию"
            };

            btnSearch = new Button
            {
                Text = "Найти",
                Location = new Point(230, 20),
                Size = new Size(80, 23)
            };
            btnSearch.Click += btnSearch_Click;

            // DataGridView
            dataGridViewMedia = new DataGridView
            {
                Location = new Point(20, 60),
                Size = new Size(660, 350),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            // Кнопки
            btnAdd = new Button
            {
                Text = "Добавить",
                Location = new Point(20, 420),
                Size = new Size(100, 30)
            };
            btnAdd.Click += btnAdd_Click;

            btnEdit = new Button
            {
                Text = "Редактировать",
                Location = new Point(130, 420),
                Size = new Size(100, 30)
            };
            btnEdit.Click += btnEdit_Click;

            btnDelete = new Button
            {
                Text = "Удалить",
                Location = new Point(240, 420),
                Size = new Size(100, 30)
            };
            btnDelete.Click += btnDelete_Click;

            btnExportExcel = new Button
            {
                Text = "Экспорт в Excel",
                Location = new Point(350, 420),
                Size = new Size(120, 30)
            };
            btnExportExcel.Click += btnExportExcel_Click;

            // Добавление элементов
            this.Controls.Add(txtSearch);
            this.Controls.Add(btnSearch);
            this.Controls.Add(dataGridViewMedia);
            this.Controls.Add(btnAdd);
            this.Controls.Add(btnEdit);
            this.Controls.Add(btnDelete);
            this.Controls.Add(btnExportExcel);
        }

        private void LoadMediaPlatforms(string searchTerm = "")
        {
            try
            {
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    connection.Open();

                    string query = "SELECT PlatformID, Name, CostPerImpression, AudienceReach FROM MediaPlatform";
                    if (!string.IsNullOrWhiteSpace(searchTerm))
                    {
                        query += " WHERE Name LIKE @SearchTerm";
                    }

                    adapter = new SqlDataAdapter(query, connection);

                    if (!string.IsNullOrWhiteSpace(searchTerm))
                    {
                        adapter.SelectCommand.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");
                    }

                    commandBuilder = new SqlCommandBuilder(adapter);
                    mediaTable = new DataTable();
                    adapter.Fill(mediaTable);

                    dataGridViewMedia.DataSource = mediaTable;

                    // Настройка колонок
                    dataGridViewMedia.Columns["PlatformID"].HeaderText = "ID";
                    dataGridViewMedia.Columns["PlatformID"].Width = 50;
                    dataGridViewMedia.Columns["Name"].HeaderText = "Название платформы";
                    dataGridViewMedia.Columns["Name"].Width = 200;
                    dataGridViewMedia.Columns["CostPerImpression"].HeaderText = "Стоимость за 1000 показов";
                    dataGridViewMedia.Columns["CostPerImpression"].Width = 150;
                    dataGridViewMedia.Columns["CostPerImpression"].DefaultCellStyle.Format = "N2";
                    dataGridViewMedia.Columns["AudienceReach"].HeaderText = "Охват аудитории";
                    dataGridViewMedia.Columns["AudienceReach"].Width = 120;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            LoadMediaPlatforms(txtSearch.Text);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            using (var form = new MediaPlatformEditForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (SqlConnection connection = DatabaseHelper.GetConnection())
                        {
                            connection.Open();

                            string query = @"INSERT INTO MediaPlatform (Name, CostPerImpression, AudienceReach) 
                                             VALUES (@Name, @CostPerImpression, @AudienceReach)";
                            SqlCommand cmd = new SqlCommand(query, connection);
                            cmd.Parameters.AddWithValue("@Name", form.PlatformName);
                            cmd.Parameters.AddWithValue("@CostPerImpression", form.CostPerImpression);
                            cmd.Parameters.AddWithValue("@AudienceReach", form.AudienceReach);

                            cmd.ExecuteNonQuery();
                        }

                        LoadMediaPlatforms(txtSearch.Text);
                        MessageBox.Show("Медиаплатформа успешно добавлена", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка добавления платформы: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (dataGridViewMedia.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите медиаплатформу для редактирования", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = dataGridViewMedia.SelectedRows[0];
            int platformId = (int)row.Cells["PlatformID"].Value;
            string name = row.Cells["Name"].Value.ToString();
            decimal cost = Convert.ToDecimal(row.Cells["CostPerImpression"].Value);
            int audience = Convert.ToInt32(row.Cells["AudienceReach"].Value);

            using (var form = new MediaPlatformEditForm(platformId, name, cost, audience))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (SqlConnection connection = DatabaseHelper.GetConnection())
                        {
                            connection.Open();

                            string query = @"UPDATE MediaPlatform 
                                            SET Name = @Name, 
                                                CostPerImpression = @CostPerImpression, 
                                                AudienceReach = @AudienceReach 
                                            WHERE PlatformID = @PlatformID";
                            SqlCommand cmd = new SqlCommand(query, connection);
                            cmd.Parameters.AddWithValue("@PlatformID", platformId);
                            cmd.Parameters.AddWithValue("@Name", form.PlatformName);
                            cmd.Parameters.AddWithValue("@CostPerImpression", form.CostPerImpression);
                            cmd.Parameters.AddWithValue("@AudienceReach", form.AudienceReach);

                            cmd.ExecuteNonQuery();
                        }

                        LoadMediaPlatforms(txtSearch.Text);
                        MessageBox.Show("Данные медиаплатформы успешно обновлены", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка обновления данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dataGridViewMedia.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите медиаплатформу для удаления", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = dataGridViewMedia.SelectedRows[0];
            int platformId = (int)row.Cells["PlatformID"].Value;
            string platformName = row.Cells["Name"].Value.ToString();

            if (MessageBox.Show($"Вы действительно хотите удалить медиаплатформу {platformName}?",
                "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection connection = DatabaseHelper.GetConnection())
                    {
                        connection.Open();

                        // Проверка связанных кампаний
                        string checkQuery = "SELECT COUNT(*) FROM Campaign WHERE PlatformID = @PlatformID";
                        SqlCommand checkCmd = new SqlCommand(checkQuery, connection);
                        checkCmd.Parameters.AddWithValue("@PlatformID", platformId);
                        int campaignCount = (int)checkCmd.ExecuteScalar();

                        if (campaignCount > 0)
                        {
                            MessageBox.Show("Невозможно удалить платформу, так как она используется в кампаниях", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Удаление платформы
                        string deleteQuery = "DELETE FROM MediaPlatform WHERE PlatformID = @PlatformID";
                        SqlCommand deleteCmd = new SqlCommand(deleteQuery, connection);
                        deleteCmd.Parameters.AddWithValue("@PlatformID", platformId);

                        int rowsAffected = deleteCmd.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            LoadMediaPlatforms(txtSearch.Text);
                            MessageBox.Show("Медиаплатформа успешно удалена", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления платформы: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Сохранить как Excel файл";
                    saveFileDialog.FileName = "Медиаплатформы_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Медиаплатформы");

                            // Заголовки
                            worksheet.Cells[1, 1].Value = "ID";
                            worksheet.Cells[1, 2].Value = "Название платформы";
                            worksheet.Cells[1, 3].Value = "Стоимость за 1000 показов";
                            worksheet.Cells[1, 4].Value = "Охват аудитории";

                            // Данные
                            int row = 2;
                            foreach (DataGridViewRow dgvRow in dataGridViewMedia.Rows)
                            {
                                worksheet.Cells[row, 1].Value = dgvRow.Cells["PlatformID"].Value;
                                worksheet.Cells[row, 2].Value = dgvRow.Cells["Name"].Value;
                                worksheet.Cells[row, 3].Value = dgvRow.Cells["CostPerImpression"].Value;
                                worksheet.Cells[row, 4].Value = dgvRow.Cells["AudienceReach"].Value;
                                row++;
                            }

                            // Форматирование стоимости
                            worksheet.Cells[2, 3, row, 3].Style.Numberformat.Format = "#,##0.00";

                            // Автоширина колонок
                            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                            // Сохранение
                            FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                            excel.SaveAs(excelFile);

                            MessageBox.Show("Экспорт в Excel выполнен успешно", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта в Excel: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    public class MediaPlatformEditForm : Form
    {
        private TextBox txtName;
        private TextBox txtCost;
        private TextBox txtAudience;
        private Button btnSave;
        private Button btnCancel;

        public string PlatformName => txtName.Text;
        public decimal CostPerImpression => decimal.Parse(txtCost.Text);
        public int AudienceReach => int.Parse(txtAudience.Text);

        public MediaPlatformEditForm(int platformId = 0, string name = "", decimal cost = 0, int audience = 0)
        {
            InitializeForm(platformId, name, cost, audience);
        }

        private void InitializeForm(int platformId, string name, decimal cost, int audience)
        {
            this.Size = new Size(400, 200);
            this.Text = platformId == 0 ? "Добавить медиаплатформу" : "Редактировать медиаплатформу";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Название
            Label lblName = new Label
            {
                Text = "Название платформы:",
                Location = new Point(20, 20),
                AutoSize = true
            };

            txtName = new TextBox
            {
                Location = new Point(150, 20),
                Size = new Size(200, 20),
                Text = name
            };

            // Стоимость
            Label lblCost = new Label
            {
                Text = "Стоимость за 1000 показов:",
                Location = new Point(20, 50),
                AutoSize = true
            };

            txtCost = new TextBox
            {
                Location = new Point(150, 50),
                Size = new Size(200, 20),
                Text = cost.ToString("N2")
            };

            // Охват аудитории
            Label lblAudience = new Label
            {
                Text = "Охват аудитории:",
                Location = new Point(20, 80),
                AutoSize = true
            };

            txtAudience = new TextBox
            {
                Location = new Point(150, 80),
                Size = new Size(200, 20),
                Text = audience.ToString()
            };

            // Кнопки
            btnSave = new Button
            {
                Text = "Сохранить",
                DialogResult = DialogResult.OK,
                Location = new Point(150, 120),
                Size = new Size(80, 30)
            };

            btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(250, 120),
                Size = new Size(80, 30)
            };

            // Добавление элементов
            this.Controls.Add(lblName);
            this.Controls.Add(txtName);
            this.Controls.Add(lblCost);
            this.Controls.Add(txtCost);
            this.Controls.Add(lblAudience);
            this.Controls.Add(txtAudience);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);

            // Валидация
            btnSave.Click += (s, e) =>
            {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Введите название платформы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus();
                this.DialogResult = DialogResult.None;
                return;
            }

            if (!decimal.TryParse(txtCost.Text, out decimal costValue) || costValue <= 0)
            {
                MessageBox.Show("Введите корректную стоимость (положительное число)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtCost.Focus();
                this.DialogResult = DialogResult.None;
                return;
            }

            if (!int.TryParse(txtAudience.Text, out int audienceValue) || audienceValue <= 0)
            {
                MessageBox.Show("Введите корректный охват аудитории (целое положительное число)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAudience.Focus();
                this.DialogResult = DialogResult.None;
                    return;
                }
            };
        }
    }
}

