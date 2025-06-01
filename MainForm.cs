using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace AdAgencyManager
{
    public partial class MainForm : Form
    {
        private Button btnProjects;
        private Button btnTasks;
        private Button btnClients;
        private Button btnMediaPlatforms;
        private Button btnReports;
        private Button btnExit;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel toolStripStatusLabelUser;
        private ToolStripStatusLabel toolStripStatusLabelStats;
        private Panel panelDashboard;
        private Label lblActiveProjects;
        private Label lblOverdueTasks;
        private Label lblTotalClients;

        public MainForm()
        {
            InitializeComponent();
            SetupForm();
            LoadDashboardStats();
        }

        private void SetupForm()
        {
            // Панель быстрой статистики
            panelDashboard = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(760, 150),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.WhiteSmoke
            };

            Label lblDashboardTitle = new Label
            {
                Text = "Общая статистика",
                Font = new Font("Arial", 12, FontStyle.Bold),
                Location = new Point(20, 10),
                AutoSize = true
            };

            Label lblActiveTitle = new Label
            {
                Text = "Активные проекты:",
                Location = new Point(30, 50),
                AutoSize = true
            };

            lblActiveProjects = new Label
            {
                Text = "0",
                Font = new Font("Arial", 14, FontStyle.Bold),
                ForeColor = Color.DarkGreen,
                Location = new Point(180, 45),
                AutoSize = true
            };

            Label lblOverdueTitle = new Label
            {
                Text = "Просроченные задачи:",
                Location = new Point(30, 90),
                AutoSize = true
            };

            lblOverdueTasks = new Label
            {
                Text = "0",
                Font = new Font("Arial", 14, FontStyle.Bold),
                ForeColor = Color.DarkRed,
                Location = new Point(180, 85),
                AutoSize = true
            };

            Label lblClientsTitle = new Label
            {
                Text = "Всего клиентов:",
                Location = new Point(30, 130),
                AutoSize = true
            };

            lblTotalClients = new Label
            {
                Text = "0",
                Font = new Font("Arial", 14, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Location = new Point(180, 125),
                AutoSize = true
            };

            // Кнопки навигации
            btnProjects = new Button
            {
                Text = "Проекты",
                Location = new Point(20, 190),
                Size = new Size(150, 60),
                Font = new Font("Arial", 10),
                BackColor = Color.LightBlue,
                Tag = "ProjectsForm"
            };
            btnProjects.Click += NavigationButton_Click;

            btnTasks = new Button
            {
                Text = "Задачи",
                Location = new Point(190, 190),
                Size = new Size(150, 60),
                Font = new Font("Arial", 10),
                BackColor = Color.LightGreen,
                Tag = "TasksForm"
            };
            btnTasks.Click += NavigationButton_Click;

            btnClients = new Button
            {
                Text = "Клиенты",
                Location = new Point(360, 190),
                Size = new Size(150, 60),
                Font = new Font("Arial", 10),
                BackColor = Color.LightYellow,
                Tag = "ClientsForm"
            };
            btnClients.Click += NavigationButton_Click;

            btnMediaPlatforms = new Button
            {
                Text = "Медиаплатформы",
                Location = new Point(20, 270),
                Size = new Size(150, 60),
                Font = new Font("Arial", 10),
                BackColor = Color.LightPink,
                Tag = "MediaPlatformsForm"
            };
            btnMediaPlatforms.Click += NavigationButton_Click;

            btnReports = new Button
            {
                Text = "Отчеты",
                Location = new Point(190, 270),
                Size = new Size(150, 60),
                Font = new Font("Arial", 10),
                BackColor = Color.LightSalmon,
                Tag = "ReportsForm"
            };
            btnReports.Click += NavigationButton_Click;

            btnExit = new Button
            {
                Text = "Выход",
                Location = new Point(360, 270),
                Size = new Size(150, 60),
                Font = new Font("Arial", 10),
                BackColor = Color.LightGray
            };
            btnExit.Click += (s, e) => Application.Exit();

            // Статус-бар
            statusStrip = new StatusStrip();
            toolStripStatusLabelUser = new ToolStripStatusLabel
            {
                Text = "Пользователь: Менеджер",
                BorderSides = ToolStripStatusLabelBorderSides.Right,
                BorderStyle = Border3DStyle.Etched
            };

            toolStripStatusLabelStats = new ToolStripStatusLabel
            {
                Text = "Загрузка статистики...",
                Spring = true,
                TextAlign = ContentAlignment.MiddleRight
            };

            statusStrip.Items.AddRange(new ToolStripItem[] {
                toolStripStatusLabelUser,
                toolStripStatusLabelStats
            });
            statusStrip.Location = new Point(0, 550);
            statusStrip.Size = new Size(800, 22);

            // Добавление элементов на панель
            panelDashboard.Controls.Add(lblDashboardTitle);
            panelDashboard.Controls.Add(lblActiveTitle);
            panelDashboard.Controls.Add(lblActiveProjects);
            panelDashboard.Controls.Add(lblOverdueTitle);
            panelDashboard.Controls.Add(lblOverdueTasks);
            panelDashboard.Controls.Add(lblClientsTitle);
            panelDashboard.Controls.Add(lblTotalClients);

            // Добавление элементов на форму
            this.Controls.Add(panelDashboard);
            this.Controls.Add(btnProjects);
            this.Controls.Add(btnTasks);
            this.Controls.Add(btnClients);
            this.Controls.Add(btnMediaPlatforms);
            this.Controls.Add(btnReports);
            this.Controls.Add(btnExit);
            this.Controls.Add(statusStrip);
        }

        private void NavigationButton_Click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null)
            {
                Form form = null;

                switch (btn.Tag.ToString())
                {
                    case "ProjectsForm":
                        form = new ProjectsForm();
                        break;
                    case "TasksForm":
                        form = new TasksForm();
                        break;
                    case "ClientsForm":
                        form = new ClientsForm();
                        break;
                    case "MediaPlatformsForm":
                        form = new MediaPlatformsForm();
                        break;
                    case "ReportsForm":
                        form = new ReportsForm();
                        break;
                }

                if (form != null)
                {
                    form.ShowDialog();
                    LoadDashboardStats(); // Обновляем статистику после закрытия формы
                }
            }
        }

        private void LoadDashboardStats()
        {
            try
            {
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    connection.Open();

                    // Активные проекты
                    string activeProjectsQuery = "SELECT COUNT(*) FROM Project WHERE EndDate >= GETDATE()";
                    SqlCommand activeCmd = new SqlCommand(activeProjectsQuery, connection);
                    int activeProjects = (int)activeCmd.ExecuteScalar();
                    lblActiveProjects.Text = activeProjects.ToString();

                    // Просроченные задачи
                    string overdueTasksQuery = @"SELECT COUNT(*) FROM Task 
                                               WHERE Status = 'В работе' 
                                                 AND Deadline < GETDATE()";
                    SqlCommand overdueCmd = new SqlCommand(overdueTasksQuery, connection);
                    int overdueTasks = (int)overdueCmd.ExecuteScalar();
                    lblOverdueTasks.Text = overdueTasks.ToString();

                    // Всего клиентов
                    string clientsQuery = "SELECT COUNT(*) FROM Client";
                    SqlCommand clientsCmd = new SqlCommand(clientsQuery, connection);
                    int totalClients = (int)clientsCmd.ExecuteScalar();
                    lblTotalClients.Text = totalClients.ToString();

                    // Обновление статус-бара
                    toolStripStatusLabelStats.Text = $"Обновлено: {DateTime.Now:HH:mm:ss} | Проектов: {activeProjects} | Задач: {overdueTasks}";
                }
            }
            catch (Exception ex)
            {
                toolStripStatusLabelStats.Text = $"Ошибка загрузки данных: {ex.Message}";
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show("Вы действительно хотите выйти из системы?",
                    "Подтверждение выхода",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
        }
    }
}