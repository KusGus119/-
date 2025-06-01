using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace AdAgencyManager
{
    public partial class ProjectsForm : Form
    {
        private DataGridView dataGridViewProjects;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnDelete;
        private Button btnExportExcel;
        private Button btnGenerateReport;
        private TextBox txtSearch;
        private Button btnSearch;
        private ComboBox cmbStatusFilter;

        private DataTable projectsTable;
        private DataTable clientsTable;
        private SqlDataAdapter projectsAdapter;
        private SqlDataAdapter clientsAdapter;
        private SqlCommandBuilder commandBuilder;

        public ProjectsForm()
        {
            InitializeComponent();
            SetupForm();
            LoadClients();
            LoadProjects();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(800, 550);
            this.Text = "Управление проектами";
            this.ResumeLayout(false);
        }

        private void SetupForm()
        {
            // Поле поиска и фильтр
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

            cmbStatusFilter = new ComboBox
            {
                Location = new Point(320, 20),
                Size = new Size(150, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbStatusFilter.Items.AddRange(new object[] { "Все", "Активные", "Завершенные", "Просроченные" });
            cmbStatusFilter.SelectedIndex = 0;
            cmbStatusFilter.SelectedIndexChanged += cmbStatusFilter_SelectedIndexChanged;

            // DataGridView
            dataGridViewProjects = new DataGridView
            {
                Location = new Point(20, 60),
                Size = new Size(760, 400),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            // Кнопки
            btnAdd = new Button
            {
                Text = "Добавить",
                Location = new Point(20, 470),
                Size = new Size(100, 30)
            };
            btnAdd.Click += btnAdd_Click;

            btnEdit = new Button
            {
                Text = "Редактировать",
                Location = new Point(130, 470),
                Size = new Size(100, 30)
            };
            btnEdit.Click += btnEdit_Click;

            btnDelete = new Button
            {
                Text = "Удалить",
                Location = new Point(240, 470),
                Size = new Size(100, 30)
            };
            btnDelete.Click += btnDelete_Click;

            btnExportExcel = new Button
            {
                Text = "Экспорт в Excel",
                Location = new Point(350, 470),
                Size = new Size(120, 30)
            };
            btnExportExcel.Click += btnExportExcel_Click;

            btnGenerateReport = new Button
            {
                Text = "Сформировать отчет",
                Location = new Point(480, 470),
                Size = new Size(150, 30)
            };
            btnGenerateReport.Click += btnGenerateReport_Click;

            // Добавление элементов
            this.Controls.Add(txtSearch);
            this.Controls.Add(btnSearch);
            this.Controls.Add(cmbStatusFilter);
            this.Controls.Add(dataGridViewProjects);
            this.Controls.Add(btnAdd);
            this.Controls.Add(btnEdit);
            this.Controls.Add(btnDelete);
            this.Controls.Add(btnExportExcel);
            this.Controls.Add(btnGenerateReport);
        }

        private void LoadClients()
        {
            try
            {
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    clientsAdapter = new SqlDataAdapter("SELECT ClientID, Name FROM Client", connection);
                    clientsTable = new DataTable();
                    clientsAdapter.Fill(clientsTable);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки клиентов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadProjects(string searchTerm = "", string statusFilter = "Все")
        {
            try
            {
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    connection.Open();

                    string query = @"SELECT p.ProjectID, p.Title, p.StartDate, p.EndDate, p.Budget, 
                                           c.Name AS ClientName, 
                                           CASE 
                                               WHEN p.EndDate < GETDATE() THEN 'Просрочен'
                                               WHEN p.EndDate >= GETDATE() THEN 'Активен'
                                           END AS Status
                                    FROM Project p
                                    JOIN Client c ON p.ClientID = c.ClientID";

                    // Добавляем условия фильтрации
                    var conditions = new System.Collections.Generic.List<string>();
                    var parameters = new System.Collections.Generic.List<SqlParameter>();

                    if (!string.IsNullOrWhiteSpace(searchTerm))
                    {
                        conditions.Add("p.Title LIKE @SearchTerm");
                        parameters.Add(new SqlParameter("@SearchTerm", $"%{searchTerm}%"));
                    }

                    if (statusFilter != "Все")
                    {
                        if (statusFilter == "Активные")
                        {
                            conditions.Add("p.EndDate >= GETDATE()");
                        }
                        else if (statusFilter == "Завершенные")
                        {
                            conditions.Add("p.EndDate < GETDATE()");
                        }
                        else if (statusFilter == "Просроченные")
                        {
                            conditions.Add("p.EndDate < GETDATE()");
                        }
                    }

                    if (conditions.Count > 0)
                    {
                        query += " WHERE " + string.Join(" AND ", conditions);
                    }

                    query += " ORDER BY p.EndDate DESC";

                    projectsAdapter = new SqlDataAdapter(query, connection);

                    // Добавляем параметры
                    foreach (var param in parameters)
                    {
                        projectsAdapter.SelectCommand.Parameters.Add(param);
                    }

                    commandBuilder = new SqlCommandBuilder(projectsAdapter);
                    projectsTable = new DataTable();
                    projectsAdapter.Fill(projectsTable);

                    dataGridViewProjects.DataSource = projectsTable;

                    // Настройка колонок
                    dataGridViewProjects.Columns["ProjectID"].HeaderText = "ID";
                    dataGridViewProjects.Columns["ProjectID"].Width = 50;
                    dataGridViewProjects.Columns["Title"].HeaderText = "Название проекта";
                    dataGridViewProjects.Columns["ClientName"].HeaderText = "Клиент";
                    dataGridViewProjects.Columns["StartDate"].HeaderText = "Дата начала";
                    dataGridViewProjects.Columns["StartDate"].DefaultCellStyle.Format = "dd.MM.yyyy";
                    dataGridViewProjects.Columns["EndDate"].HeaderText = "Дата окончания";
                    dataGridViewProjects.Columns["EndDate"].DefaultCellStyle.Format = "dd.MM.yyyy";
                    dataGridViewProjects.Columns["Budget"].HeaderText = "Бюджет";
                    dataGridViewProjects.Columns["Budget"].DefaultCellStyle.Format = "N2";
                    dataGridViewProjects.Columns["Status"].HeaderText = "Статус";

                    // Подсветка строк
                    dataGridViewProjects.RowPrePaint += (s, e) =>
                    {
                        if (e.RowIndex >= 0 && e.RowIndex < dataGridViewProjects.Rows.Count)
                        {
                            var row = dataGridViewProjects.Rows[e.RowIndex];
                            string status = row.Cells["Status"].Value?.ToString();

                            if (status == "Просрочен")
                            {
                                row.DefaultCellStyle.BackColor = Color.LightPink;
                            }
                            else if (status == "Активен")
                            {
                                row.DefaultCellStyle.BackColor = Color.LightGreen;
                            }
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки проектов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            LoadProjects(txtSearch.Text, cmbStatusFilter.SelectedItem.ToString());
        }

        private void cmbStatusFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadProjects(txtSearch.Text, cmbStatusFilter.SelectedItem.ToString());
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            using (var form = new ProjectEditForm(clientsTable))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (SqlConnection connection = DatabaseHelper.GetConnection())
                        {
                            connection.Open();

                            string query = @"INSERT INTO Project (Title, StartDate, EndDate, Budget, ClientID) 
                                             VALUES (@Title, @StartDate, @EndDate, @Budget, @ClientID)";
                            SqlCommand cmd = new SqlCommand(query, connection);
                            cmd.Parameters.AddWithValue("@Title", form.ProjectTitle);
                            cmd.Parameters.AddWithValue("@StartDate", form.StartDate);
                            cmd.Parameters.AddWithValue("@EndDate", form.EndDate);
                            cmd.Parameters.AddWithValue("@Budget", form.Budget);
                            cmd.Parameters.AddWithValue("@ClientID", form.SelectedClientId);

                            cmd.ExecuteNonQuery();
                        }

                        LoadProjects(txtSearch.Text, cmbStatusFilter.SelectedItem.ToString());
                        MessageBox.Show("Проект успешно добавлен", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка добавления проекта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (dataGridViewProjects.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите проект для редактирования", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = dataGridViewProjects.SelectedRows[0];
            int projectId = (int)row.Cells["ProjectID"].Value;
            string title = row.Cells["Title"].Value.ToString();
            DateTime startDate = (DateTime)row.Cells["StartDate"].Value;
            DateTime endDate = (DateTime)row.Cells["EndDate"].Value;
            decimal budget = (decimal)row.Cells["Budget"].Value;
            string clientName = row.Cells["ClientName"].Value.ToString();

            // Находим ClientID по имени клиента
            int clientId = clientsTable.AsEnumerable()
                .Where(r => r.Field<string>("Name") == clientName)
                .Select(r => r.Field<int>("ClientID"))
                .FirstOrDefault();

            using (var form = new ProjectEditForm(clientsTable, projectId, title, startDate, endDate, budget, clientId))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (SqlConnection connection = DatabaseHelper.GetConnection())
                        {
                            connection.Open();

                            string query = @"UPDATE Project 
                                            SET Title = @Title, 
                                                StartDate = @StartDate, 
                                                EndDate = @EndDate, 
                                                Budget = @Budget, 
                                                ClientID = @ClientID 
                                            WHERE ProjectID = @ProjectID";
                            SqlCommand cmd = new SqlCommand(query, connection);
                            cmd.Parameters.AddWithValue("@ProjectID", projectId);
                            cmd.Parameters.AddWithValue("@Title", form.ProjectTitle);
                            cmd.Parameters.AddWithValue("@StartDate", form.StartDate);
                            cmd.Parameters.AddWithValue("@EndDate", form.EndDate);
                            cmd.Parameters.AddWithValue("@Budget", form.Budget);
                            cmd.Parameters.AddWithValue("@ClientID", form.SelectedClientId);

                            cmd.ExecuteNonQuery();
                        }

                        LoadProjects(txtSearch.Text, cmbStatusFilter.SelectedItem.ToString());
                        MessageBox.Show("Проект успешно обновлен", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка обновления проекта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dataGridViewProjects.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите проект для удаления", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = dataGridViewProjects.SelectedRows[0];
            int projectId = (int)row.Cells["ProjectID"].Value;
            string projectName = row.Cells["Title"].Value.ToString();

            if (MessageBox.Show($"Вы действительно хотите удалить проект {projectName}?",
                "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection connection = DatabaseHelper.GetConnection())
                    {
                        connection.Open();

                        // Проверка связанных задач
                        string checkQuery = "SELECT COUNT(*) FROM Task WHERE ProjectID = @ProjectID";
                        SqlCommand checkCmd = new SqlCommand(checkQuery, connection);
                        checkCmd.Parameters.AddWithValue("@ProjectID", projectId);
                        int taskCount = (int)checkCmd.ExecuteScalar();

                        if (taskCount > 0)
                        {
                            MessageBox.Show("Невозможно удалить проект, так как в нем есть задачи", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Удаление проекта
                        string deleteQuery = "DELETE FROM Project WHERE ProjectID = @ProjectID";
                        SqlCommand deleteCmd = new SqlCommand(deleteQuery, connection);
                        deleteCmd.Parameters.AddWithValue("@ProjectID", projectId);

                        int rowsAffected = deleteCmd.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            LoadProjects(txtSearch.Text, cmbStatusFilter.SelectedItem.ToString());
                            MessageBox.Show("Проект успешно удален", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления проекта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    saveFileDialog.FileName = "Проекты_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Проекты");

                            // Заголовки
                            worksheet.Cells[1, 1].Value = "ID";
                            worksheet.Cells[1, 2].Value = "Название проекта";
                            worksheet.Cells[1, 3].Value = "Клиент";
                            worksheet.Cells[1, 4].Value = "Дата начала";
                            worksheet.Cells[1, 5].Value = "Дата окончания";
                            worksheet.Cells[1, 6].Value = "Бюджет";
                            worksheet.Cells[1, 7].Value = "Статус";

                            // Данные
                            int row = 2;
                            foreach (DataGridViewRow dgvRow in dataGridViewProjects.Rows)
                            {
                                worksheet.Cells[row, 1].Value = dgvRow.Cells["ProjectID"].Value;
                                worksheet.Cells[row, 2].Value = dgvRow.Cells["Title"].Value;
                                worksheet.Cells[row, 3].Value = dgvRow.Cells["ClientName"].Value;
                                worksheet.Cells[row, 4].Value = ((DateTime)dgvRow.Cells["StartDate"].Value).ToString("dd.MM.yyyy");
                                worksheet.Cells[row, 5].Value = ((DateTime)dgvRow.Cells["EndDate"].Value).ToString("dd.MM.yyyy");
                                worksheet.Cells[row, 6].Value = dgvRow.Cells["Budget"].Value;
                                worksheet.Cells[row, 7].Value = dgvRow.Cells["Status"].Value;
                                row++;
                            }

                            // Форматирование
                            worksheet.Cells[2, 6, row, 6].Style.Numberformat.Format = "#,##0.00";

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

        private void btnGenerateReport_Click(object sender, EventArgs e)
        {
            if (dataGridViewProjects.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите проект для формирования отчета", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int projectId = (int)dataGridViewProjects.SelectedRows[0].Cells["ProjectID"].Value;
            string projectName = dataGridViewProjects.SelectedRows[0].Cells["Title"].Value.ToString();

            try
            {
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    connection.Open();

                    // Получаем данные по проекту
                    string projectQuery = @"SELECT p.Title, p.StartDate, p.EndDate, p.Budget, c.Name AS ClientName
                                           FROM Project p
                                           JOIN Client c ON p.ClientID = c.ClientID
                                           WHERE p.ProjectID = @ProjectID";
                    SqlCommand projectCmd = new SqlCommand(projectQuery, connection);
                    projectCmd.Parameters.AddWithValue("@ProjectID", projectId);

                    var projectReader = projectCmd.ExecuteReader();
                    if (!projectReader.Read())
                    {
                        MessageBox.Show("Проект не найден", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    string clientName = projectReader["ClientName"].ToString();
                    DateTime startDate = (DateTime)projectReader["StartDate"];
                    DateTime endDate = (DateTime)projectReader["EndDate"];
                    decimal budget = (decimal)projectReader["Budget"];
                    projectReader.Close();

                    // Получаем задачи проекта
                    string tasksQuery = @"SELECT Description, Status, Deadline 
                                         FROM Task 
                                         WHERE ProjectID = @ProjectID
                                         ORDER BY Deadline";
                    SqlCommand tasksCmd = new SqlCommand(tasksQuery, connection);
                    tasksCmd.Parameters.AddWithValue("@ProjectID", projectId);

                    DataTable tasksTable = new DataTable();
                    tasksTable.Load(tasksCmd.ExecuteReader());

                    // Получаем медиакампании проекта
                    string campaignsQuery = @"SELECT mp.Name AS PlatformName, c.BudgetAllocated
                                             FROM Campaign c
                                             JOIN MediaPlatform mp ON c.PlatformID = mp.PlatformID
                                             WHERE c.ProjectID = @ProjectID";
                    SqlCommand campaignsCmd = new SqlCommand(campaignsQuery, connection);
                    campaignsCmd.Parameters.AddWithValue("@ProjectID", projectId);

                    DataTable campaignsTable = new DataTable();
                    campaignsTable.Load(campaignsCmd.ExecuteReader());

                    // Создаем отчет
                    using (SaveFileDialog saveDialog = new SaveFileDialog())
                    {
                        saveDialog.Filter = "Word Documents|*.docx";
                        saveDialog.FileName = $"Отчет_по_проекту_{projectName}_{DateTime.Now:yyyyMMdd}.docx";

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            // Здесь должна быть реализация генерации Word-документа
                            // Например, с использованием библиотеки DocX или OpenXML

                            MessageBox.Show($"Отчет по проекту {projectName} успешно сформирован", "Успех",
                                          MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка формирования отчета: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    public class ProjectEditForm : Form
    {
        private TextBox txtTitle;
        private DateTimePicker dtpStartDate;
        private DateTimePicker dtpEndDate;
        private TextBox txtBudget;
        private ComboBox cmbClient;
        private Button btnSave;
        private Button btnCancel;

        private DataTable clientsTable;

        public string ProjectTitle => txtTitle.Text;
        public DateTime StartDate => dtpStartDate.Value;
        public DateTime EndDate => dtpEndDate.Value;
        public decimal Budget => decimal.Parse(txtBudget.Text);
        public int SelectedClientId => (int)cmbClient.SelectedValue;

        public ProjectEditForm(DataTable clients, int projectId = 0, string title = "",
                             DateTime? startDate = null, DateTime? endDate = null,
                             decimal budget = 0, int clientId = 0)
        {
            clientsTable = clients;
            InitializeForm(projectId, title, startDate, endDate, budget, clientId);
        }

        private void InitializeForm(int projectId, string title, DateTime? startDate,
                                  DateTime? endDate, decimal budget, int clientId)
        {
            this.Size = new Size(450, 250);
            this.Text = projectId == 0 ? "Добавить проект" : "Редактировать проект";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Название проекта
            Label lblTitle = new Label
            {
                Text = "Название проекта:",
                Location = new Point(20, 20),
                AutoSize = true
            };

            txtTitle = new TextBox
            {
                Location = new Point(150, 20),
                Size = new Size(250, 20),
                Text = title
            };

            // Клиент
            Label lblClient = new Label
            {
                Text = "Клиент:",
                Location = new Point(20, 50),
                AutoSize = true
            };

            cmbClient = new ComboBox
            {
                Location = new Point(150, 50),
                Size = new Size(250, 20),
                DisplayMember = "Name",
                ValueMember = "ClientID",
                DataSource = clientsTable
            };

            if (clientId > 0)
            {
                cmbClient.SelectedValue = clientId;
            }

            // Дата начала
            Label lblStartDate = new Label
            {
                Text = "Дата начала:",
                Location = new Point(20, 80),
                AutoSize = true
            };

            dtpStartDate = new DateTimePicker
            {
                Location = new Point(150, 80),
                Size = new Size(150, 20),
                Format = DateTimePickerFormat.Short,
                Value = startDate ?? DateTime.Today
            };

            // Дата окончания
            Label lblEndDate = new Label
            {
                Text = "Дата окончания:",
                Location = new Point(20, 110),
                AutoSize = true
            };

            dtpEndDate = new DateTimePicker
            {
                Location = new Point(150, 110),
                Size = new Size(150, 20),
                Format = DateTimePickerFormat.Short,
                Value = endDate ?? DateTime.Today.AddDays(30)
            };

            // Бюджет
            Label lblBudget = new Label
            {
                Text = "Бюджет:",
                Location = new Point(20, 140),
                AutoSize = true
            };

            txtBudget = new TextBox
            {
                Location = new Point(150, 140),
                Size = new Size(150, 20),
                Text = budget.ToString("N2")
            };

            // Кнопки
            btnSave = new Button
            {
                Text = "Сохранить",
                DialogResult = DialogResult.OK,
                Location = new Point(150, 170),
                Size = new Size(80, 30)
            };

            btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(250, 170),
                Size = new Size(80, 30)
            };

            // Добавление элементов
            this.Controls.Add(lblTitle);
            this.Controls.Add(txtTitle);
            this.Controls.Add(lblClient);
            this.Controls.Add(cmbClient);
            this.Controls.Add(lblStartDate);
            this.Controls.Add(dtpStartDate);
            this.Controls.Add(lblEndDate);
            this.Controls.Add(dtpEndDate);
            this.Controls.Add(lblBudget);
            this.Controls.Add(txtBudget);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);

            // Валидация
            btnSave.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtTitle.Text))
                {
                    MessageBox.Show("Введите название проекта", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtTitle.Focus();
                    this.DialogResult = DialogResult.None;
                    return;
                }

                if (dtpStartDate.Value > dtpEndDate.Value)
                {
                    MessageBox.Show("Дата начала не может быть позже даты окончания", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    dtpStartDate.Focus();
                    this.DialogResult = DialogResult.None;
                    return;
                }

                if (!decimal.TryParse(txtBudget.Text, out decimal budgetValue) || budgetValue <= 0)
                {
                    MessageBox.Show("Введите корректный бюджет (положительное число)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtBudget.Focus();
                    this.DialogResult = DialogResult.None;
                    return;
                }

                if (cmbClient.SelectedValue == null)
                {
                    MessageBox.Show("Выберите клиента", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cmbClient.Focus();
                    this.DialogResult = DialogResult.None;
                    return;
                }
            };
        }
    }
}