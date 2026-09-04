using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace PaddyTrolleyManagement
{
    public partial class Dashboard : Form
    {
        // =====================================================
        // DASHBOARD VALUE LABELS
        // =====================================================

        private Label lblMonthPurchaseValue;
        private Label lblMonthDispatchValue;
        private Label lblActiveDumpsValue;
        private Label lblTotalUnitsValue;

        // =====================================================
        // YEAR AND MONTH FILTERS
        // =====================================================

        private ComboBox cmbYear;
        private ComboBox cmbMonth;

        private bool isLoading = false;


        // =====================================================
        // CONSTRUCTOR
        // =====================================================

        public Dashboard()
        {
            InitializeComponent();

            CreateUI();

            LoadYears();
            LoadDashboardData();
        }


        // =====================================================
        // CREATE DASHBOARD UI
        // =====================================================

        private void CreateUI()
        {
            this.Text = "Rana Group - Paddy Management System";
            this.Size = new Size(1100, 760);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.WhiteSmoke;


            // =================================================
            // HEADER
            // =================================================

            Panel headerPanel = new Panel();

            headerPanel.Size = new Size(1100, 110);
            headerPanel.Location = new Point(0, 0);
            headerPanel.BackColor = Color.FromArgb(30, 80, 50);

            this.Controls.Add(headerPanel);


            // COMPANY NAME

            Label lblCompany = new Label();

            lblCompany.Text = "RANA GROUP";

            lblCompany.Font =
                new Font(
                    "Segoe UI",
                    24,
                    FontStyle.Bold
                );

            lblCompany.ForeColor = Color.White;
            lblCompany.AutoSize = true;
            lblCompany.Location = new Point(420, 15);

            headerPanel.Controls.Add(lblCompany);


            // SYSTEM NAME

            Label lblSystem = new Label();

            lblSystem.Text =
                "PADDY MANAGEMENT SYSTEM";

            lblSystem.Font =
                new Font(
                    "Segoe UI",
                    15,
                    FontStyle.Regular
                );

            lblSystem.ForeColor = Color.White;
            lblSystem.AutoSize = true;
            lblSystem.Location = new Point(375, 55);

            headerPanel.Controls.Add(lblSystem);


            // =================================================
            // YEAR SELECTOR
            // =================================================

            Label lblSelectYear = new Label();

            lblSelectYear.Text = "Select Year:";

            lblSelectYear.Font =
                new Font(
                    "Segoe UI",
                    11,
                    FontStyle.Bold
                );

            lblSelectYear.AutoSize = true;
            lblSelectYear.Location =
                new Point(40, 135);

            this.Controls.Add(lblSelectYear);


            cmbYear = new ComboBox();

            cmbYear.Location =
                new Point(150, 130);

            cmbYear.Size =
                new Size(150, 35);

            cmbYear.Font =
                new Font(
                    "Segoe UI",
                    11,
                    FontStyle.Regular
                );

            cmbYear.DropDownStyle =
                ComboBoxStyle.DropDownList;

            this.Controls.Add(cmbYear);


            // =================================================
            // MONTH SELECTOR
            // =================================================

            Label lblSelectMonth = new Label();

            lblSelectMonth.Text = "Select Month:";

            lblSelectMonth.Font =
                new Font(
                    "Segoe UI",
                    11,
                    FontStyle.Bold
                );

            lblSelectMonth.AutoSize = true;
            lblSelectMonth.Location =
                new Point(350, 135);

            this.Controls.Add(lblSelectMonth);


            cmbMonth = new ComboBox();

            cmbMonth.Location =
                new Point(475, 130);

            cmbMonth.Size =
                new Size(180, 35);

            cmbMonth.Font =
                new Font(
                    "Segoe UI",
                    11,
                    FontStyle.Regular
                );

            cmbMonth.DropDownStyle =
                ComboBoxStyle.DropDownList;


            // ADD MONTHS

            cmbMonth.Items.Add("January");
            cmbMonth.Items.Add("February");
            cmbMonth.Items.Add("March");
            cmbMonth.Items.Add("April");
            cmbMonth.Items.Add("May");
            cmbMonth.Items.Add("June");
            cmbMonth.Items.Add("July");
            cmbMonth.Items.Add("August");
            cmbMonth.Items.Add("September");
            cmbMonth.Items.Add("October");
            cmbMonth.Items.Add("November");
            cmbMonth.Items.Add("December");


            // SELECT CURRENT MONTH

            cmbMonth.SelectedIndex =
                DateTime.Today.Month - 1;

            this.Controls.Add(cmbMonth);


            // =================================================
            // SUMMARY CARDS
            // =================================================

            Panel purchaseCard =
                CreateSummaryCard(
                    "MONTH PURCHASE",
                    new Point(40, 190)
                );

            lblMonthPurchaseValue =
                CreateCardValueLabel();

            purchaseCard.Controls.Add(
                lblMonthPurchaseValue
            );

            this.Controls.Add(purchaseCard);


            Panel dispatchCard =
                CreateSummaryCard(
                    "MONTH DISPATCH",
                    new Point(290, 190)
                );

            lblMonthDispatchValue =
                CreateCardValueLabel();

            dispatchCard.Controls.Add(
                lblMonthDispatchValue
            );

            this.Controls.Add(dispatchCard);


            Panel activeDumpCard =
                CreateSummaryCard(
                    "ACTIVE DUMPS",
                    new Point(540, 190)
                );

            lblActiveDumpsValue =
                CreateCardValueLabel();

            activeDumpCard.Controls.Add(
                lblActiveDumpsValue
            );

            this.Controls.Add(activeDumpCard);


            Panel totalUnitCard =
                CreateSummaryCard(
                    "TOTAL UNITS",
                    new Point(790, 190)
                );

            lblTotalUnitsValue =
                CreateCardValueLabel();

            totalUnitCard.Controls.Add(
                lblTotalUnitsValue
            );

            this.Controls.Add(totalUnitCard);


            // =================================================
            // SYSTEM MODULES TITLE
            // =================================================

            Label lblNavigation = new Label();

            lblNavigation.Text =
                "SYSTEM MODULES";

            lblNavigation.Font =
                new Font(
                    "Segoe UI",
                    14,
                    FontStyle.Bold
                );

            lblNavigation.AutoSize = true;

            lblNavigation.Location =
                new Point(40, 370);

            this.Controls.Add(lblNavigation);


            // =================================================
            // DAILY ENTRY BUTTON
            // =================================================

            Button btnDailyEntry =
                CreateNavigationButton(
                    "DAILY ENTRY",
                    new Point(40, 420)
                );

            btnDailyEntry.Click +=
                (s, e) =>
                {
                    DailyEntry dailyEntryForm =
                        new DailyEntry();

                    dailyEntryForm.ShowDialog();

                    LoadYears();
                    LoadDashboardData();
                };

            this.Controls.Add(btnDailyEntry);


            // =================================================
            // VIEW / EDIT ENTRIES BUTTON
            // =================================================

            Button btnViewEdit =
                CreateNavigationButton(
                    "VIEW / EDIT ENTRIES",
                    new Point(290, 420)
                );

            btnViewEdit.Click +=
                (s, e) =>
                {
                    ViewEditDailyEntries viewEditForm =
                        new ViewEditDailyEntries();

                    viewEditForm.ShowDialog();

                    LoadYears();
                    LoadDashboardData();
                };

            this.Controls.Add(btnViewEdit);


            // =================================================
            // DAILY REPORT BUTTON
            // =================================================

            Button btnDailyReport =
                CreateNavigationButton(
                    "DAILY REPORT",
                    new Point(540, 420)
                );

            btnDailyReport.Click +=
                (s, e) =>
                {
                    DailyReport dailyReportForm =
                        new DailyReport();

                    dailyReportForm.ShowDialog();
                };

            this.Controls.Add(btnDailyReport);


            // =================================================
            // MANAGE DUMPS BUTTON
            // =================================================

            Button btnManageDumps =
                CreateNavigationButton(
                    "MANAGE DUMPS",
                    new Point(790, 420)
                );

            btnManageDumps.Click +=
                (s, e) =>
                {
                    ManageDumps manageDumps =
                        new ManageDumps();

                    manageDumps.ShowDialog();

                    LoadDashboardData();
                };

            this.Controls.Add(btnManageDumps);


            // =================================================
            // PADDY ANALYTICS BUTTON
            // =================================================

            Button btnPaddyAnalytics =
                CreateNavigationButton(
                    "PADDY ANALYTICS",
                    new Point(540, 510)
                );

            btnPaddyAnalytics.Click +=
                (s, e) =>
                {
                    PaddyAnalytics analyticsForm =
                        new PaddyAnalytics();

                    analyticsForm.ShowDialog();
                };

            this.Controls.Add(btnPaddyAnalytics);

            // =================================================
            // DATABASE BACKUP BUTTON
            // =================================================

            Button btnDatabaseBackup =
                CreateNavigationButton(
                    "DATABASE BACKUP",
                    new Point(40, 510)
                );

            btnDatabaseBackup.Click +=
                (s, e) =>
                {
                    DatabaseBackup backupForm =
                        new DatabaseBackup();

                    backupForm.ShowDialog();
                };

            this.Controls.Add(btnDatabaseBackup);


            // =================================================
            // DATABASE RESTORE BUTTON
            // =================================================

            Button btnDatabaseRestore =
                CreateNavigationButton(
                    "RESTORE DATABASE",
                    new Point(290, 510)
                );

            btnDatabaseRestore.Click +=
                (s, e) =>
                {
                    DatabaseRestore restoreForm =
                        new DatabaseRestore();

                    restoreForm.ShowDialog();

                    // Refresh dashboard after restore
                    LoadYears();
                    LoadDashboardData();
                };

            this.Controls.Add(btnDatabaseRestore);


            // =================================================
            // FOOTER
            // =================================================

            Label lblFooter = new Label();

            lblFooter.Text =
                "Rana Group | Paddy Monitoring & Reporting System";

            lblFooter.Font =
                new Font(
                    "Segoe UI",
                    9,
                    FontStyle.Regular
                );

            lblFooter.AutoSize = true;

            lblFooter.Location =
                new Point(390, 650);

            this.Controls.Add(lblFooter);


            // =================================================
            // FILTER EVENTS
            // =================================================

            cmbYear.SelectedIndexChanged +=
                (s, e) =>
                {
                    if (!isLoading)
                    {
                        LoadDashboardData();
                    }
                };


            cmbMonth.SelectedIndexChanged +=
                (s, e) =>
                {
                    if (!isLoading)
                    {
                        LoadDashboardData();
                    }
                };
        }


        // =====================================================
        // LOAD AVAILABLE YEARS
        // =====================================================

        private void LoadYears()
        {
            try
            {
                isLoading = true;

                cmbYear.Items.Clear();

                using (
                    SqlConnection con =
                    DatabaseHelper.GetConnection()
                )
                {
                    con.Open();

                    string query = @"
                        SELECT DISTINCT
                            YEAR(EntryDate) AS EntryYear
                        FROM DailyTrolleyDetails
                        WHERE EntryDate IS NOT NULL
                        ORDER BY EntryYear DESC";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(query, con)
                    )
                    {
                        using (
                            SqlDataReader reader =
                            cmd.ExecuteReader()
                        )
                        {
                            while (reader.Read())
                            {
                                cmbYear.Items.Add(
                                    reader["EntryYear"].ToString()
                                );
                            }
                        }
                    }
                }


                string currentYear =
                    DateTime.Today.Year.ToString();

                int currentYearIndex =
                    cmbYear.Items.IndexOf(currentYear);


                if (currentYearIndex >= 0)
                {
                    cmbYear.SelectedIndex =
                        currentYearIndex;
                }
                else if (cmbYear.Items.Count > 0)
                {
                    cmbYear.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load years.\n\n" +
                    ex.Message
                );
            }
            finally
            {
                isLoading = false;
            }
        }


        // =====================================================
        // CREATE SUMMARY CARD
        // =====================================================

        private Panel CreateSummaryCard(
            string title,
            Point location)
        {
            Panel card = new Panel();

            card.Size =
                new Size(220, 130);

            card.Location = location;

            card.BackColor = Color.White;

            card.BorderStyle =
                BorderStyle.FixedSingle;


            Label lblTitle = new Label();

            lblTitle.Text = title;

            lblTitle.Font =
                new Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lblTitle.AutoSize = true;

            lblTitle.Location =
                new Point(20, 20);

            card.Controls.Add(lblTitle);

            return card;
        }


        // =====================================================
        // CREATE CARD VALUE LABEL
        // =====================================================

        private Label CreateCardValueLabel()
        {
            Label lblValue = new Label();

            lblValue.Text = "0";

            lblValue.Font =
                new Font(
                    "Segoe UI",
                    28,
                    FontStyle.Bold
                );

            lblValue.AutoSize = true;

            lblValue.Location =
                new Point(20, 55);

            return lblValue;
        }


        // =====================================================
        // CREATE NAVIGATION BUTTON
        // =====================================================

        private Button CreateNavigationButton(
            string text,
            Point location)
        {
            Button button = new Button();

            button.Text = text;

            button.Size =
                new Size(220, 65);

            button.Location = location;

            button.Font =
                new Font(
                    "Segoe UI",
                    11,
                    FontStyle.Bold
                );

            return button;
        }


        // =====================================================
        // LOAD DASHBOARD DATA
        // =====================================================

        private void LoadDashboardData()
        {
            if (
                cmbYear == null ||
                cmbMonth == null ||
                cmbYear.SelectedItem == null ||
                cmbMonth.SelectedIndex < 0
            )
            {
                return;
            }


            int selectedYear =
                Convert.ToInt32(
                    cmbYear.SelectedItem
                );

            int selectedMonth =
                cmbMonth.SelectedIndex + 1;


            DateTime startDate =
                new DateTime(
                    selectedYear,
                    selectedMonth,
                    1
                );

            DateTime endDate =
                startDate.AddMonths(1);


            try
            {
                using (
                    SqlConnection con =
                    DatabaseHelper.GetConnection()
                )
                {
                    con.Open();


                    // MONTH PURCHASE

                    string purchaseQuery = @"
                        SELECT ISNULL(
                            SUM(
                                dt.DayTrolley +
                                dt.NightTrolley
                            ),
                            0
                        )
                        FROM DailyTrolleyDetails dt
                        INNER JOIN Groups g
                            ON dt.GroupID = g.GroupID
                        WHERE
                            dt.EntryDate >= @StartDate
                            AND dt.EntryDate < @EndDate
                            AND g.GroupType = 'Purchase'";


                    lblMonthPurchaseValue.Text =
                        ExecuteScalarWithDate(
                            con,
                            purchaseQuery,
                            startDate,
                            endDate
                        ).ToString();


                    // MONTH DISPATCH

                    string dispatchQuery = @"
                        SELECT ISNULL(
                            SUM(
                                dt.DayTrolley +
                                dt.NightTrolley
                            ),
                            0
                        )
                        FROM DailyTrolleyDetails dt
                        INNER JOIN Groups g
                            ON dt.GroupID = g.GroupID
                        WHERE
                            dt.EntryDate >= @StartDate
                            AND dt.EntryDate < @EndDate
                            AND g.GroupType = 'Dispatch'";


                    lblMonthDispatchValue.Text =
                        ExecuteScalarWithDate(
                            con,
                            dispatchQuery,
                            startDate,
                            endDate
                        ).ToString();


                    // ACTIVE DUMPS

                    string activeDumpsQuery = @"
                        SELECT COUNT(*)
                        FROM Groups
                        WHERE IsActive = 1";


                    lblActiveDumpsValue.Text =
                        ExecuteScalar(
                            con,
                            activeDumpsQuery
                        ).ToString();


                    // TOTAL ACTIVE UNITS

                    string totalUnitsQuery = @"
                        SELECT COUNT(*)
                        FROM Units
                        WHERE IsActive = 1";


                    lblTotalUnitsValue.Text =
                        ExecuteScalar(
                            con,
                            totalUnitsQuery
                        ).ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load Dashboard data.\n\n" +
                    ex.Message
                );
            }
        }


        // =====================================================
        // EXECUTE QUERY WITH DATE PARAMETERS
        // =====================================================

        private int ExecuteScalarWithDate(
            SqlConnection con,
            string query,
            DateTime startDate,
            DateTime endDate)
        {
            using (
                SqlCommand cmd =
                new SqlCommand(query, con)
            )
            {
                cmd.Parameters.AddWithValue(
                    "@StartDate",
                    startDate
                );

                cmd.Parameters.AddWithValue(
                    "@EndDate",
                    endDate
                );


                object result =
                    cmd.ExecuteScalar();


                if (
                    result == null ||
                    result == DBNull.Value
                )
                {
                    return 0;
                }

                return Convert.ToInt32(result);
            }
        }


        // =====================================================
        // EXECUTE NORMAL SCALAR QUERY
        // =====================================================

        private int ExecuteScalar(
            SqlConnection con,
            string query)
        {
            using (
                SqlCommand cmd =
                new SqlCommand(query, con)
            )
            {
                object result =
                    cmd.ExecuteScalar();


                if (
                    result == null ||
                    result == DBNull.Value
                )
                {
                    return 0;
                }

                return Convert.ToInt32(result);
            }
        }
    }
}