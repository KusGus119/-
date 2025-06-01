using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Diagnostics;

namespace AdAgencyManager
{
    public partial class ReportsForm : Form
    {
        private ComboBox cmbReportType;
        private DateTimePicker dtpStartDate;
        private DateTimePicker dtpEndDate;
        private Button btnGenerate;
        private Button btnExport;
        private DataGridView dataGridViewReport;
        private ProgressBar progressBar;

        private DataTable reportData;

        public ReportsForm()
        {
            InitializeComponent();
            SetupForm();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(800, 600);
            this.Text = "Генерация отчетов";
            this.ResumeLayout(false);
        }

        private void SetupForm()
        {
            // Выбор типа отчета
            Label lblReportType = new Label
            {
                Text = "Тип отчета:",
                Location = new Point(20, 20),
                AutoSize = true
            };

            cmbReportType = new ComboBox
            {
                Location = new Point(150, 20),
                Size = new Size(200, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbReportType.Items.AddRange(new object[] {
                "Отчет по проектам",
                "Отчет по клиентам",
                "Отчет по медиаплатформам",
                "Финансовый отчет",
                "Отчет по выполнению задач"
            });
            cmbReportType.SelectedIndex = 0;
            cmbReportType.SelectedIndexChanged += cmbReportType_SelectedIndexChanged;

            // Период отчета
            Label lblPeriod = new Label
            {
                Text = "Период:",
                Location = new Point(20, 50),
                AutoSize = true
            };

            dtpStartDate = new DateTimePicker
            {
                Location = new Point(150, 50),
                Size = new Size(120, 20),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today.AddMonths(-1)
            };

            Label lblTo = new Label
            {
                Text = "по",
                Location = new Point(280, 50),
                AutoSize = true
            };

            dtpEndDate = new DateTimePicker
            {
                Location = new Point(310, 50),
                Size = new Size(120, 20),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today
            };

            // Кнопки
            btnGenerate = new Button
            {
                Text = "Сформировать",
                Location = new Point(20, 80),
                Size = new Size(120, 30)
            };
            btnGenerate.Click += btnGenerate_Click;

            btnExport = new Button
            {
                Text = "Экспорт в Excel",
                Location = new Point(150, 80),
                Size = new Size(120, 30),
                Enabled = false
            };
            btnExport.Click += btnExport_Click;

            // DataGridView
            dataGridViewReport = new DataGridView
            {
                Location = new Point(20, 120),
                Size = new Size(760, 400),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false
            };

            // ProgressBar
            progressBar = new ProgressBar
            {
                Location = new Point(20, 530),
                Size = new Size(760, 20),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Visible = false
            };

            // Добавление элементов
            this.Controls.Add(lblReportType);
            this.Controls.Add(cmbReportType);
            this.Controls.Add(lblPeriod);
            this.Controls.Add(dtpStartDate);
            this.Controls.Add(lblTo);
            this.Controls.Add(dtpEndDate);
            this.Controls.Add(btnGenerate);
            this.Controls.Add(btnExport);
            this.Controls.Add(dataGridViewReport);
            this.Controls.Add(progressBar);
        }

        private void cmbReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Сбрасываем данные при изменении типа отчета
            reportData = null;
            dataGridViewReport.DataSource = null;
            btnExport.Enabled = false;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (dtpStartDate.Value > dtpEndDate.Value)
            {
                MessageBox.Show("Дата начала не может быть позже даты окончания", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            progressBar.Visible = true;
            btnGenerate.Enabled = false;

            // Запускаем в отдельном потоке, чтобы не блокировать UI
            System.Threading.Tasks.Task.Run(() => GenerateReport());
        }

        private void GenerateReport()
        {
            try
            {
                string reportType = cmbReportType.SelectedItem.ToString();
                DateTime startDate = dtpStartDate.Value;
                DateTime endDate = dtpEndDate.Value;

                this.Invoke((MethodInvoker)delegate {
                    progressBar.Style = ProgressBarStyle.Marquee;
                });

                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    connection.Open();
                    string query = "";

                    switch (reportType)
                    {
                        case "Отчет по проектам":
                            query = @"SELECT 
                                        p.ProjectID AS 'ID',
                                        p.Title AS 'Название проекта',
                                        c.Name AS 'Клиент',
                                        p.StartDate AS 'Дата начала',
                                        p.EndDate AS 'Дата окончания',
                                        p.Budget AS 'Бюджет',
                                        CASE 
                                            WHEN p.EndDate < GETDATE() THEN 'Завершен'
                                            ELSE 'Активен'
                                        END AS 'Статус',
                                        (SELECT COUNT(*) FROM Task t WHERE t.ProjectID = p.ProjectID) AS 'Кол-во задач',
                                        (SELECT COUNT(*) FROM Task t WHERE t.ProjectID = p.ProjectID AND t.Status = 'Завершено') AS 'Завершено задач'
                                    FROM Project p
                                    JOIN Client c ON p.ClientID = c.ClientID
                                    WHERE p.StartDate BETWEEN @StartDate AND @EndDate
                                    ORDER BY p.EndDate";
                            break;

                        case "Отчет по клиентам":
                            query = @"SELECT 
                                        c.ClientID AS 'ID',
                                        c.Name AS 'Клиент',
                                        COUNT(p.ProjectID) AS 'Кол-во проектов',
                                        SUM(p.Budget) AS 'Общий бюджет',
                                        MIN(p.StartDate) AS 'Первый проект',
                                        MAX(p.EndDate) AS 'Последний проект'
                                    FROM Client c
                                    LEFT JOIN Project p ON c.ClientID = p.ClientID
                                    WHERE (p.StartDate BETWEEN @StartDate AND @EndDate OR p.StartDate IS NULL)
                                    GROUP BY c.ClientID, c.Name
                                    ORDER BY SUM(p.Budget) DESC";
                            break;

                        case "Отчет по медиаплатформам":
                            query = @"SELECT 
                                        mp.PlatformID AS 'ID',
                                        mp.Name AS 'Платформа',
                                        COUNT(c.CampaignID) AS 'Кол-во кампаний',
                                        SUM(c.BudgetAllocated) AS 'Общий бюджет',
                                        mp.CostPerImpression AS 'Стоимость за 1000 показов',
                                        mp.AudienceReach AS 'Охват аудитории'
                                    FROM MediaPlatform mp
                                    LEFT JOIN Campaign c ON mp.PlatformID = c.PlatformID
                                    LEFT JOIN Project p ON c.ProjectID = p.ProjectID
                                    WHERE (p.StartDate BETWEEN @StartDate AND @EndDate OR p.StartDate IS NULL)
                                    GROUP BY mp.PlatformID, mp.Name, mp.CostPerImpression, mp.AudienceReach
                                    ORDER BY SUM(c.BudgetAllocated) DESC";
                            break;

                        case "Финансовый отчет":
                            query = @"SELECT 
                                        c.Name AS 'Клиент',
                                        p.Title AS 'Проект',
                                        p.Budget AS 'Бюджет проекта',
                                        (SELECT SUM(BudgetAllocated) FROM Campaign WHERE ProjectID = p.ProjectID) AS 'Распределенный бюджет',
                                        p.Budget - (SELECT ISNULL(SUM(BudgetAllocated), 0) FROM Campaign WHERE ProjectID = p.ProjectID) AS 'Остаток бюджета',
                                        (SELECT COUNT(*) FROM Task WHERE ProjectID = p.ProjectID) AS 'Всего задач',
                                        (SELECT COUNT(*) FROM Task WHERE ProjectID = p.ProjectID AND Status = 'Завершено') AS 'Завершено задач',
                                        CAST((SELECT COUNT(*) FROM Task WHERE ProjectID = p.ProjectID AND Status = 'Завершено') AS FLOAT) / 
                                            NULLIF((SELECT COUNT(*) FROM Task WHERE ProjectID = p.ProjectID), 0) * 100 AS 'Процент выполнения'
                                    FROM Project p
                                    JOIN Client c ON p.ClientID = c.ClientID
                                    WHERE p.StartDate BETWEEN @StartDate AND @EndDate
                                    ORDER BY c.Name, p.Title";
                            break;

                        case "Отчет по выполнению задач":
                            query = @"SELECT 
                                        p.Title AS 'Проект',
                                        t.Description AS 'Задача',
                                        t.Deadline AS 'Срок выполнения',
                                        t.Status AS 'Статус',
                                        DATEDIFF(day, GETDATE(), t.Deadline) AS 'Дней до дедлайна',
                                        CASE 
                                            WHEN t.Status = 'Завершено' THEN 'Выполнено вовремя'
                                            WHEN t.Deadline < GETDATE() THEN 'Просрочена'
                                            ELSE 'В работе'
                                        END AS 'Состояние'
                                    FROM Task t
                                    JOIN Project p ON t.ProjectID = p.ProjectID
                                    WHERE t.Deadline BETWEEN @StartDate AND @EndDate
                                    ORDER BY t.Deadline";
                            break;
                    }

                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    adapter.SelectCommand.Parameters.AddWithValue("@StartDate", startDate);
                    adapter.SelectCommand.Parameters.AddWithValue("@EndDate", endDate.AddDays(1)); // Чтобы включить весь день окончания

                    reportData = new DataTable();
                    adapter.Fill(reportData);

                    this.Invoke((MethodInvoker)delegate {
                        dataGridViewReport.DataSource = reportData;
                        FormatDataGridView(reportType);
                        btnExport.Enabled = true;
                        progressBar.Visible = false;
                        btnGenerate.Enabled = true;
                    });
                }
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate {
                    MessageBox.Show($"Ошибка генерации отчета: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBar.Visible = false;
                    btnGenerate.Enabled = true;
                });
            }
        }

        private void FormatDataGridView(string reportType)
        {
            // Очищаем предыдущее форматирование
            dataGridViewReport.DefaultCellStyle.BackColor = Color.White;
            dataGridViewReport.DefaultCellStyle.ForeColor = Color.Black;
            dataGridViewReport.RowsDefaultCellStyle = null;

            // Применяем форматирование в зависимости от типа отчета
            switch (reportType)
            {
                case "Отчет по выполнению задач":
                    foreach (DataGridViewRow row in dataGridViewReport.Rows)
                    {
                        if (row.Cells["Состояние"].Value?.ToString() == "Просрочена")
                        {
                            row.DefaultCellStyle.BackColor = Color.LightPink;
                        }
                        else if (row.Cells["Состояние"].Value?.ToString() == "Выполнено вовремя")
                        {
                            row.DefaultCellStyle.BackColor = Color.LightGreen;
                        }
                    }
                    break;

                case "Финансовый отчет":
                    dataGridViewReport.Columns["Бюджет проекта"].DefaultCellStyle.Format = "N2";
                    dataGridViewReport.Columns["Распределенный бюджет"].DefaultCellStyle.Format = "N2";
                    dataGridViewReport.Columns["Остаток бюджета"].DefaultCellStyle.Format = "N2";
                    dataGridViewReport.Columns["Процент выполнения"].DefaultCellStyle.Format = "N1";
                    break;
            }

            // Автоматическая ширина столбцов
            dataGridViewReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (reportData == null || reportData.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    string reportType = cmbReportType.SelectedItem.ToString();
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Экспорт отчета в Excel";
                    saveFileDialog.FileName = $"{reportType}_{DateTime.Now:yyyyMMdd}.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Continuous;
                        progressBar.Value = 0;

                        System.Threading.Tasks.Task.Run(() => ExportToExcel(saveFileDialog.FileName));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при подготовке экспорта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToExcel(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excel = new ExcelPackage())
                {
                    string reportType = cmbReportType.SelectedItem.ToString();
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(reportType);

                    // Заголовки
                    for (int i = 0; i < reportData.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = reportData.Columns[i].ColumnName;
                        worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    // Данные
                    int totalRows = reportData.Rows.Count;
                    for (int row = 0; row < totalRows; row++)
                    {
                        for (int col = 0; col < reportData.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = reportData.Rows[row][col];
                        }

                        // Обновление прогресса
                        int progress = (int)((row + 1) / (double)totalRows * 100);
                        this.Invoke((MethodInvoker)delegate {
                            progressBar.Value = progress;
                        });
                    }

                    // Форматирование
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Специфическое форматирование для числовых полей
                    if (reportType == "Финансовый отчет")
                    {
                        int budgetCol = GetColumnIndex(reportData, "Бюджет проекта");
                        int allocatedCol = GetColumnIndex(reportData, "Распределенный бюджет");
                        int remainingCol = GetColumnIndex(reportData, "Остаток бюджета");

                        if (budgetCol >= 0)
                            worksheet.Cells[2, budgetCol + 1, totalRows + 1, budgetCol + 1].Style.Numberformat.Format = "#,##0.00";
                        if (allocatedCol >= 0)
                            worksheet.Cells[2, allocatedCol + 1, totalRows + 1, allocatedCol + 1].Style.Numberformat.Format = "#,##0.00";
                        if (remainingCol >= 0)
                            worksheet.Cells[2, remainingCol + 1, totalRows + 1, remainingCol + 1].Style.Numberformat.Format = "#,##0.00";
                    }

                    // Сохранение файла
                    FileInfo excelFile = new FileInfo(filePath);
                    excel.SaveAs(excelFile);

                    this.Invoke((MethodInvoker)delegate {
                        progressBar.Visible = false;
                        if (MessageBox.Show("Отчет успешно экспортирован. Открыть файл?", "Успех",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate {
                    MessageBox.Show($"Ошибка экспорта в Excel: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBar.Visible = false;
                });
            }
        }

        private int GetColumnIndex(DataTable table, string columnName)
        {
            for (int i = 0; i < table.Columns.Count; i++)
            {
                if (table.Columns[i].ColumnName == columnName)
                    return i;
            }
            return -1;
        }
    }
}