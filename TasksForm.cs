using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace AdAgencyManager
{
    public partial class TasksForm : Form
    {
        private DataGridView dataGridViewTasks;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnDelete;
        private Button btnChangeStatus;
        private Button btnExportExcel;
        private ComboBox cmbProjectFilter;
        private ComboBox cmbStatusFilter;
        private Button btnApplyFilters;

        private DataTable tasksTable;
        private DataTable projectsTable;
        private SqlDataAdapter tasksAdapter;
        private SqlDataAdapter projectsAdapter;
        private SqlCommandBuilder commandBuilder;

        public TasksForm()
        {
            InitializeComponent();
            SetupForm();
            LoadProjects();
            LoadTasks();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(800, 550);
            this.Text = "Управление задачами";
            this.ResumeLayout(false);
        }

        private void SetupForm()
        {
            // Фильтры
            Label lblProjectFilter = new Label
            {
                Text = "Проект:",
                Location = new Point(20, 20),
                AutoSize = true
            };

            cmbProjectFilter = new ComboBox
            {
                Location = new Point(70, 20),
                Size = new Size(200, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            Label lblStatusFilter = new Label
            {
                Text = "Статус:",
                Location = new Point(280, 20),
                AutoSize = true
            };

            cmbStatusFilter = new ComboBox
            {
                Location = new Point(330, 20),
                Size = new Size(150, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbStatusFilter.Items.AddRange(new object[] { "Все", "В работе", "Завершено" });
            cmbStatusFilter.SelectedIndex = 0;

            btnApplyFilters = new Button
            {
                Text = "Применить",
                Location = new Point(490, 20),
                Size = new Size(80, 23)
            };
            btnApplyFilters.Click += btnApplyFilters_Click;

            // DataGridView
            dataGridViewTasks = new DataGridView
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

            btnChangeStatus = new Button
            {
                Text = "Изменить статус",
                Location = new Point(350, 470),
                Size = new Size(120, 30)
            };
            btnChangeStatus.Click += btnChangeStatus_Click;

            btnExportExcel = new Button
            {
                Text = "Экспорт в Excel",
                Location = new Point(480, 470),
                Size = new Size(120, 30)
            };
            btnExportExcel.Click += btnExportExcel_Click;

            // Добавление элементов
            this.Controls.Add(lblProjectFilter);
            this.Controls.Add(cmbProjectFilter);
            this.Controls.Add(lblStatusFilter);
            this.Controls.Add(cmbStatusFilter);
            this.Controls.Add(btnApplyFilters);
            this.Controls.Add(dataGridViewTasks);
            this.Controls.Add(btnAdd);
            this.Controls.Add(btnEdit);
            this.Controls.Add(btnDelete);
            this.Controls.Add(btnChangeStatus);
            this.Controls.Add(btnExportExcel);
        }

        private void LoadProjects()
        {
            try
            {
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    projectsAdapter = new SqlDataAdapter(
                        "SELECT ProjectID, Title FROM Project ORDER BY Title",
                        connection);

                    projectsTable = new DataTable();
                    projectsAdapter.Fill(projectsTable);

                    cmbProjectFilter.DataSource = new DataView(projectsTable);
                    cmbProjectFilter.DisplayMember = "Title";
                    cmbProjectFilter.ValueMember = "ProjectID";

                    // Добавляем элемент "Все проекты"
                    DataRow allProjectsRow = projectsTable.NewRow();
                    allProjectsRow["ProjectID"] = 0;
                    allProjectsRow["Title"] = "Все проекты";
                    projectsTable.Rows.InsertAt(allProjectsRow, 0);

                    cmbProjectFilter.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки проектов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadTasks(int projectId = 0, string statusFilter = "Все")
        {
            try
            {
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    connection.Open();

                    string query = @"SELECT 
                                        t.TaskID,
                                        t.Description,
                                        t.Status,
                                        t.Deadline,
                                        p.Title AS ProjectTitle,
                                        CASE 
                                            WHEN t.Deadline < GETDATE() AND t.Status != 'Завершено' THEN 'Просрочена'
                                            ELSE ''
                                        END AS Overdue
                                    FROM Task t
                                    JOIN Project p ON t.ProjectID = p.ProjectID
                                    WHERE 1=1";

                    var parameters = new System.Collections.Generic.List<SqlParameter>();

                    if (projectId > 0)
                    {
                        query += " AND t.ProjectID = @ProjectID";
                        parameters.Add(new SqlParameter("@ProjectID", projectId));
                    }

                    if (statusFilter != "Все")
                    {
                        query += " AND t.Status = @Status";
                        parameters.Add(new SqlParameter("@Status", statusFilter));
                    }

                    query += " ORDER BY t.Deadline";

                    tasksAdapter = new SqlDataAdapter(query, connection);

                    foreach (var param in parameters)
                    {
                        tasksAdapter.SelectCommand.Parameters.Add(param);
                    }

                    commandBuilder = new SqlCommandBuilder(tasksAdapter);
                    tasksTable = new DataTable();
                    tasksAdapter.Fill(tasksTable);

                    dataGridViewTasks.DataSource = tasksTable;

                    // Настройка колонок
                    dataGridViewTasks.Columns["TaskID"].HeaderText = "ID";
                    dataGridViewTasks.Columns["TaskID"].Width = 40;
                    dataGridViewTasks.Columns["Description"].HeaderText = "Описание";
                    dataGridViewTasks.Columns["Status"].HeaderText = "Статус";
                    dataGridViewTasks.Columns["Status"].Width = 80;
                    dataGridViewTasks.Columns["Deadline"].HeaderText = "Срок выполнения";
                    dataGridViewTasks.Columns["Deadline"].DefaultCellStyle.Format = "dd.MM.yyyy";
                    dataGridViewTasks.Columns["ProjectTitle"].HeaderText = "Проект";
                    dataGridViewTasks.Columns["Overdue"].HeaderText = " ";
                    dataGridViewTasks.Columns["Overdue"].Width = 30;

                    // Подсветка просроченных задач
                    dataGridViewTasks.RowPrePaint += (s, e) =>
                    {
                        if (e.RowIndex >= 0 && e.RowIndex < dataGridViewTasks.Rows.Count)
                        {
                            var row = dataGridViewTasks.Rows[e.RowIndex];
                            if (row.Cells["Overdue"].Value?.ToString() == "Просрочена")
                            {
                                row.DefaultCellStyle.BackColor = Color.LightPink;
                            }
                            else if (row.Cells["Status"].Value?.ToString() == "Завершено")
                            {
                                row.DefaultCellStyle.BackColor = Color.LightGreen;
                            }
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки задач: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnApplyFilters_Click(object sender, EventArgs e)
        {
            int projectId = 0;
            if (cmbProjectFilter.SelectedValue != null && cmbProjectFilter.SelectedIndex > 0)
            {
                projectId = (int)cmbProjectFilter.SelectedValue;
            }

            LoadTasks(projectId, cmbStatusFilter.SelectedItem.ToString());
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            using (var form = new TaskEditForm(projectsTable))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (SqlConnection connection = DatabaseHelper.GetConnection())
                        {
                            connection.Open();

                            string query = @"INSERT INTO Task (Description, Status, Deadline, ProjectID) 
                                             VALUES (@Description, @Status, @Deadline, @ProjectID)";
                            SqlCommand cmd = new SqlCommand(query, connection);
                            cmd.Parameters.AddWithValue("@Description", form.TaskDescription);
                            cmd.Parameters.AddWithValue("@Status", "В работе");
                            cmd.Parameters.AddWithValue("@Deadline", form.Deadline);
                            cmd.Parameters.AddWithValue("@ProjectID", form.SelectedProjectId);

                            cmd.ExecuteNonQuery();
                        }

                        btnApplyFilters_Click(null, null);
                        MessageBox.Show("Задача успешно добавлена", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка добавления задачи: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (dataGridViewTasks.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите задачу для редактирования", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = dataGridViewTasks.SelectedRows[0];
            int taskId = (int)row.Cells["TaskID"].Value;
            string description = row.Cells["Description"].Value.ToString();
            string status = row.Cells["Status"].Value.ToString();
            DateTime deadline = (DateTime)row.Cells["Deadline"].Value;
            string projectTitle = row.Cells["ProjectTitle"].Value.ToString();

            // Находим ProjectID по названию проекта
            int projectId = projectsTable.AsEnumerable()
                .Where(r => r.Field<string>("Title") == projectTitle)
                .Select(r => r.Field<int>("ProjectID"))
                .FirstOrDefault();

            using (var form = new TaskEditForm(projectsTable, taskId, description, status, deadline, projectId))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (SqlConnection connection = DatabaseHelper.GetConnection())
                        {
                            connection.Open();

                            string query = @"UPDATE Task 
                                            SET Description = @Description, 
                                                Deadline = @Deadline, 
                                                ProjectID = @ProjectID 
                                            WHERE TaskID = @TaskID";
                            SqlCommand cmd = new SqlCommand(query, connection);
                            cmd.Parameters.AddWithValue("@TaskID", taskId);
                            cmd.Parameters.AddWithValue("@Description", form.TaskDescription);
                            cmd.Parameters.AddWithValue("@Deadline", form.Deadline);
                            cmd.Parameters.AddWithValue("@ProjectID", form.SelectedProjectId);

                            cmd.ExecuteNonQuery();
                        }

                        btnApplyFilters_Click(null, null);
                        MessageBox.Show("Задача успешно обновлена", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка обновления задачи: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dataGridViewTasks.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите задачу для удаления", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = dataGridViewTasks.SelectedRows[0];
            int taskId = (int)row.Cells["TaskID"].Value;
            string description = row.Cells["Description"].Value.ToString();

            if (MessageBox.Show($"Вы действительно хотите удалить задачу: \"{description}\"?",
                "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection connection = DatabaseHelper.GetConnection())
                    {
                        connection.Open();

                        string deleteQuery = "DELETE FROM Task WHERE TaskID = @TaskID";
                        SqlCommand deleteCmd = new SqlCommand(deleteQuery, connection);
                        deleteCmd.Parameters.AddWithValue("@TaskID", taskId);

                        int rowsAffected = deleteCmd.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            btnApplyFilters_Click(null, null);
                            MessageBox.Show("Задача успешно удалена", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления задачи: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnChangeStatus_Click(object sender, EventArgs e)
        {
            if (dataGridViewTasks.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите задачу для изменения статуса", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = dataGridViewTasks.SelectedRows[0];
            int taskId = (int)row.Cells["TaskID"].Value;
            string currentStatus = row.Cells["Status"].Value.ToString();
            string newStatus = currentStatus == "В работе" ? "Завершено" : "В работе";

            try
            {
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    connection.Open();

                    string query = "UPDATE Task SET Status = @Status WHERE TaskID = @TaskID";
                    SqlCommand cmd = new SqlCommand(query, connection);
                    cmd.Parameters.AddWithValue("@Status", newStatus);
                    cmd.Parameters.AddWithValue("@TaskID", taskId);

                    int rowsAffected = cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        btnApplyFilters_Click(null, null);
                        MessageBox.Show($"Статус задачи изменен на \"{newStatus}\"", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка изменения статуса: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (tasksTable == null || tasksTable.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Экспорт задач в Excel";
                    saveFileDialog.FileName = $"Задачи_{DateTime.Now:yyyyMMdd}.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Задачи");

                            // Заголовки
                            worksheet.Cells[1, 1].Value = "ID";
                            worksheet.Cells[1, 2].Value = "Описание";
                            worksheet.Cells[1, 3].Value = "Статус";
                            worksheet.Cells[1, 4].Value = "Срок выполнения";
                            worksheet.Cells[1, 5].Value = "Проект";
                            worksheet.Cells[1, 6].Value = "Просрочена";

                            // Данные
                            for (int i = 0; i < tasksTable.Rows.Count; i++)
                            {
                                DataRow row = tasksTable.Rows[i];
                                worksheet.Cells[i + 2, 1].Value = row["TaskID"];
                                worksheet.Cells[i + 2, 2].Value = row["Description"];
                                worksheet.Cells[i + 2, 3].Value = row["Status"];
                                worksheet.Cells[i + 2, 4].Value = ((DateTime)row["Deadline"]).ToString("dd.MM.yyyy");
                                worksheet.Cells[i + 2, 5].Value = row["ProjectTitle"];
                                worksheet.Cells[i + 2, 6].Value = row["Overdue"];

                                // Подсветка просроченных задач
                                if (row["Overdue"].ToString() == "Просрочена")
                                {
                                    worksheet.Cells[i + 2, 1, i + 2, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    worksheet.Cells[i + 2, 1, i + 2, 6].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                                }
                                else if (row["Status"].ToString() == "Завершено")
                                {
                                    worksheet.Cells[i + 2, 1, i + 2, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    worksheet.Cells[i + 2, 1, i + 2, 6].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                }
                            }

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

    public class TaskEditForm : Form
    {
        private TextBox txtDescription;
        private DateTimePicker dtpDeadline;
        private ComboBox cmbProject;
        private Button btnSave;
        private Button btnCancel;

        private DataTable projectsTable;

        public string TaskDescription => txtDescription.Text;
        public DateTime Deadline => dtpDeadline.Value;
        public int SelectedProjectId => (int)cmbProject.SelectedValue;

        public TaskEditForm(DataTable projects, int taskId = 0, string description = "",
                          string status = "", DateTime? deadline = null, int projectId = 0)
        {
            projectsTable = projects;
            InitializeForm(taskId, description, status, deadline, projectId);
        }

        private void InitializeForm(int taskId, string description, string status,
                                  DateTime? deadline, int projectId)
        {
            this.Size = new Size(500, 250);
            this.Text = taskId == 0 ? "Добавить задачу" : "Редактировать задачу";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Описание задачи
            Label lblDescription = new Label
            {
                Text = "Описание:",
                Location = new Point(20, 20),
                AutoSize = true
            };

            txtDescription = new TextBox
            {
                Location = new Point(150, 20),
                Size = new Size(300, 20),
                Text = description,
                Multiline = true,
                Height = 60
            };

            // Проект
            Label lblProject = new Label
            {
                Text = "Проект:",
                Location = new Point(20, 90),
                AutoSize = true
            };

            cmbProject = new ComboBox
            {
                Location = new Point(150, 90),
                Size = new Size(300, 20),
                DisplayMember = "Title",
                ValueMember = "ProjectID",
                DataSource = projectsTable.DefaultView
            };

            if (projectId > 0)
            {
                cmbProject.SelectedValue = projectId;
            }

            // Срок выполнения
            Label lblDeadline = new Label
            {
                Text = "Срок выполнения:",
                Location = new Point(20, 120),
                AutoSize = true
            };

            dtpDeadline = new DateTimePicker
            {
                Location = new Point(150, 120),
                Size = new Size(150, 20),
                Format = DateTimePickerFormat.Short,
                Value = deadline ?? DateTime.Today.AddDays(7)
            };

            // Кнопки
            btnSave = new Button
            {
                Text = "Сохранить",
                DialogResult = DialogResult.OK,
                Location = new Point(150, 160),
                Size = new Size(80, 30)
            };

            btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(250, 160),
                Size = new Size(80, 30)
            };

            // Добавление элементов
            this.Controls.Add(lblDescription);
            this.Controls.Add(txtDescription);
            this.Controls.Add(lblProject);
            this.Controls.Add(cmbProject);
            this.Controls.Add(lblDeadline);
            this.Controls.Add(dtpDeadline);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);

            // Валидация
            btnSave.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtDescription.Text))
                {
                    MessageBox.Show("Введите описание задачи", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtDescription.Focus();
                    this.DialogResult = DialogResult.None;
                    return;
                }

                if (cmbProject.SelectedValue == null || (int)cmbProject.SelectedValue == 0)
                {
                    MessageBox.Show("Выберите проект", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cmbProject.Focus();
                    this.DialogResult = DialogResult.None;
                    return;
                }
            };
        }
    }
}