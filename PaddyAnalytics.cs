using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Data.SqlClient;

using iTextSharp.text;
using iTextSharp.text.pdf;

using Excel = Microsoft.Office.Interop.Excel;

namespace PaddyTrolleyManagement
{
    public partial class PaddyAnalytics : Form
    {
        // =====================================================
        // CONTROLS
        // =====================================================

        private DateTimePicker dtFrom;
        private DateTimePicker dtTo;

        private Label lblPurchase;
        private Label lblDumps;

        private Label lblHighestDump;
        private Label lblHighestDate;

        private DataGridView dgvDumpWise;
        private DataGridView dgvDaily;


        // =====================================================
        // CONSTRUCTOR
        // =====================================================

        public PaddyAnalytics()
        {
            InitializeComponent();

            CreateAnalyticsUI();

            // Current month by default
            dtFrom.Value = new DateTime(
                DateTime.Today.Year,
                DateTime.Today.Month,
                1
            );

            dtTo.Value = DateTime.Today;

            LoadAnalytics();
        }


        // =====================================================
        // CREATE UI
        // =====================================================

        private void CreateAnalyticsUI()
        {
            this.Text =
                "Rana Group - Paddy Purchase MIS & Analytics";

            this.Size =
                new Size(1200, 800);

            this.StartPosition =
                FormStartPosition.CenterScreen;

            this.BackColor =
                Color.FromArgb(245, 247, 249);

            this.FormBorderStyle =
                FormBorderStyle.FixedSingle;

            this.MaximizeBox = false;


            // =====================================================
            // HEADER
            // =====================================================

            Panel header =
                new Panel();

            header.Location =
                new Point(0, 0);

            header.Size =
                new Size(1200, 95);

            header.BackColor =
                Color.FromArgb(30, 80, 50);

            this.Controls.Add(header);


            Label lblCompany =
                new Label();

            lblCompany.Text =
                "RANA GROUP";

            lblCompany.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    24,
                    FontStyle.Bold
                );

            lblCompany.ForeColor =
                Color.White;

            lblCompany.AutoSize = true;

            lblCompany.Location =
                new Point(485, 12);

            header.Controls.Add(lblCompany);


            Label lblSystem =
                new Label();

            lblSystem.Text =
                "PADDY PURCHASE MIS & ANALYTICS";

            lblSystem.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    13,
                    FontStyle.Regular
                );

            lblSystem.ForeColor =
                Color.White;

            lblSystem.AutoSize = true;

            lblSystem.Location =
                new Point(430, 58);

            header.Controls.Add(lblSystem);


            // =====================================================
            // FILTER PANEL
            // =====================================================

            Panel filterPanel =
                new Panel();

            filterPanel.Location =
                new Point(15, 110);

            filterPanel.Size =
                new Size(1150, 65);

            filterPanel.BackColor =
                Color.White;

            filterPanel.BorderStyle =
                BorderStyle.FixedSingle;

            this.Controls.Add(filterPanel);


            // FROM DATE
            Label lblFrom =
                new Label();

            lblFrom.Text =
                "From Date:";

            lblFrom.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lblFrom.AutoSize = true;

            lblFrom.Location =
                new Point(15, 20);

            filterPanel.Controls.Add(lblFrom);


            dtFrom =
                new DateTimePicker();

            dtFrom.Format =
                DateTimePickerFormat.Short;

            dtFrom.Location =
                new Point(90, 16);

            dtFrom.Size =
                new Size(135, 30);

            filterPanel.Controls.Add(dtFrom);


            // TO DATE
            Label lblTo =
                new Label();

            lblTo.Text =
                "To Date:";

            lblTo.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lblTo.AutoSize = true;

            lblTo.Location =
                new Point(250, 20);

            filterPanel.Controls.Add(lblTo);


            dtTo =
                new DateTimePicker();

            dtTo.Format =
                DateTimePickerFormat.Short;

            dtTo.Location =
                new Point(315, 16);

            dtTo.Size =
                new Size(135, 30);

            filterPanel.Controls.Add(dtTo);


            // PURCHASE ONLY LABEL
            Label lblReport =
                new Label();

            lblReport.Text =
                "Report:";

            lblReport.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lblReport.AutoSize = true;

            lblReport.Location =
                new Point(475, 20);

            filterPanel.Controls.Add(lblReport);


            Label lblPurchaseOnly =
                new Label();

            lblPurchaseOnly.Text =
                "PURCHASE ONLY";

            lblPurchaseOnly.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lblPurchaseOnly.ForeColor =
                Color.FromArgb(30, 80, 50);

            lblPurchaseOnly.AutoSize = true;

            lblPurchaseOnly.Location =
                new Point(535, 20);

            filterPanel.Controls.Add(
                lblPurchaseOnly
            );


            // =====================================================
            // GENERATE BUTTON
            // =====================================================

            Button btnGenerate =
                CreateActionButton(
                    "GENERATE",
                    new Point(700, 14),
                    new Size(120, 36)
                );

            btnGenerate.Click +=
                (s, e) =>
                {
                    LoadAnalytics();
                };

            filterPanel.Controls.Add(
                btnGenerate
            );


            // =====================================================
            // PDF BUTTON
            // =====================================================

            Button btnExportPdf =
                CreateActionButton(
                    "EXPORT PDF",
                    new Point(830, 14),
                    new Size(120, 36)
                );

            btnExportPdf.Click +=
                BtnExportPdf_Click;

            filterPanel.Controls.Add(
                btnExportPdf
            );


            // =====================================================
            // EXCEL BUTTON
            // =====================================================

            Button btnExportExcel =
                CreateActionButton(
                    "EXPORT EXCEL",
                    new Point(960, 14),
                    new Size(130, 36)
                );

            btnExportExcel.Click +=
                BtnExportExcel_Click;

            filterPanel.Controls.Add(
                btnExportExcel
            );


            // =====================================================
            // TOTAL PURCHASE CARD
            // =====================================================

            Panel purchaseCard =
                CreateCard(
                    "TOTAL PURCHASE",
                    new Point(15, 195),
                    new Size(270, 120)
                );

            lblPurchase =
                CreateValueLabel();

            purchaseCard.Controls.Add(
                lblPurchase
            );

            this.Controls.Add(
                purchaseCard
            );


            // =====================================================
            // ACTIVE PURCHASE DUMPS CARD
            // =====================================================

            Panel dumpsCard =
                CreateCard(
                    "ACTIVE PURCHASE DUMPS",
                    new Point(305, 195),
                    new Size(270, 120)
                );

            lblDumps =
                CreateValueLabel();

            dumpsCard.Controls.Add(
                lblDumps
            );

            this.Controls.Add(
                dumpsCard
            );


            // =====================================================
            // HIGHEST PURCHASE DUMP
            // =====================================================

            Panel highestDumpCard =
                CreateInfoCard(
                    "HIGHEST PURCHASE DUMP",
                    new Point(595, 195),
                    new Size(270, 120)
                );

            lblHighestDump =
                CreateInfoValueLabel();

            highestDumpCard.Controls.Add(
                lblHighestDump
            );

            this.Controls.Add(
                highestDumpCard
            );


            // =====================================================
            // HIGHEST PURCHASE DATE
            // =====================================================

            Panel highestDateCard =
                CreateInfoCard(
                    "HIGHEST PURCHASE DATE",
                    new Point(885, 195),
                    new Size(270, 120)
                );

            lblHighestDate =
                CreateInfoValueLabel();

            highestDateCard.Controls.Add(
                lblHighestDate
            );

            this.Controls.Add(
                highestDateCard
            );


            // =====================================================
            // DUMP-WISE TITLE
            // =====================================================

            Label lblDumpTitle =
                new Label();

            lblDumpTitle.Text =
                "DUMP-WISE PURCHASE ANALYSIS";

            lblDumpTitle.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    12,
                    FontStyle.Bold
                );

            lblDumpTitle.AutoSize = true;

            lblDumpTitle.Location =
                new Point(20, 340);

            this.Controls.Add(
                lblDumpTitle
            );


            // =====================================================
            // DATE-WISE TITLE
            // =====================================================

            Label lblDateTitle =
                new Label();

            lblDateTitle.Text =
                "DATE-WISE PURCHASE ANALYSIS";

            lblDateTitle.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    12,
                    FontStyle.Bold
                );

            lblDateTitle.AutoSize = true;

            lblDateTitle.Location =
                new Point(605, 340);

            this.Controls.Add(
                lblDateTitle
            );


            // =====================================================
            // DUMP-WISE GRID
            // =====================================================

            dgvDumpWise =
                CreateGrid();

            dgvDumpWise.Location =
                new Point(15, 370);

            dgvDumpWise.Size =
                new Size(550, 300);

            this.Controls.Add(
                dgvDumpWise
            );


            // =====================================================
            // DATE-WISE GRID
            // =====================================================

            dgvDaily =
                CreateGrid();

            dgvDaily.Location =
                new Point(600, 370);

            dgvDaily.Size =
                new Size(555, 300);

            this.Controls.Add(
                dgvDaily
            );


            // =====================================================
            // FOOTER
            // =====================================================

            Label lblFooter =
                new Label();

            lblFooter.Text =
                "Rana Group | Paddy Monitoring & Reporting System";

            lblFooter.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    9,
                    FontStyle.Regular
                );

            lblFooter.ForeColor =
                Color.DimGray;

            lblFooter.AutoSize = true;

            lblFooter.Location =
                new Point(430, 710);

            this.Controls.Add(
                lblFooter
            );
        }


        // =====================================================
        // ACTION BUTTON
        // =====================================================

        private Button CreateActionButton(
            string text,
            Point location,
            Size size)
        {
            Button btn =
                new Button();

            btn.Text = text;

            btn.Size = size;

            btn.Location =
                location;

            btn.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    9,
                    FontStyle.Bold
                );

            btn.BackColor =
                Color.White;

            btn.FlatStyle =
                FlatStyle.Standard;

            btn.Cursor =
                Cursors.Hand;

            return btn;
        }


        // =====================================================
        // CREATE CARD
        // =====================================================

        private Panel CreateCard(
            string title,
            Point location,
            Size size)
        {
            Panel card =
                new Panel();

            card.Location =
                location;

            card.Size =
                size;

            card.BackColor =
                Color.White;

            card.BorderStyle =
                BorderStyle.FixedSingle;


            Label lblTitle =
                new Label();

            lblTitle.Text =
                title;

            lblTitle.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lblTitle.ForeColor =
                Color.FromArgb(30, 80, 50);

            lblTitle.AutoSize = true;

            lblTitle.Location =
                new Point(18, 18);

            card.Controls.Add(
                lblTitle
            );

            return card;
        }


        // =====================================================
        // INFORMATION CARD
        // =====================================================

        private Panel CreateInfoCard(
            string title,
            Point location,
            Size size)
        {
            Panel card =
                CreateCard(
                    title,
                    location,
                    size
                );

            return card;
        }


        // =====================================================
        // VALUE LABEL
        // =====================================================

        private Label CreateValueLabel()
        {
            Label lbl =
                new Label();

            lbl.Text =
                "0";

            lbl.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    30,
                    FontStyle.Bold
                );

            lbl.ForeColor =
                Color.FromArgb(20, 20, 20);

            lbl.AutoSize = true;

            lbl.Location =
                new Point(18, 52);

            return lbl;
        }


        // =====================================================
        // INFORMATION VALUE LABEL
        // =====================================================

        private Label CreateInfoValueLabel()
        {
            Label lbl =
                new Label();

            lbl.Text =
                "No data";

            lbl.Font =
                new System.Drawing.Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lbl.ForeColor =
                Color.FromArgb(20, 20, 20);

            lbl.AutoSize = false;

            lbl.Size =
                new Size(235, 60);

            lbl.Location =
                new Point(18, 48);

            return lbl;
        }


        // =====================================================
        // CREATE GRID
        // =====================================================

        private DataGridView CreateGrid()
        {
            DataGridView grid =
                new DataGridView();

            grid.ReadOnly = true;

            grid.AllowUserToAddRows =
                false;

            grid.AllowUserToDeleteRows =
                false;

            grid.AllowUserToResizeRows =
                false;

            grid.RowHeadersVisible =
                false;

            grid.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;

            grid.MultiSelect = false;

            grid.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            grid.BackgroundColor =
                Color.White;

            grid.BorderStyle =
                BorderStyle.FixedSingle;

            grid.EnableHeadersVisualStyles =
                false;

            grid.ColumnHeadersDefaultCellStyle =
                new DataGridViewCellStyle
                {
                    BackColor =
                        Color.FromArgb(30, 80, 50),

                    ForeColor =
                        Color.White,

                    Font =
                        new System.Drawing.Font(
                            "Segoe UI",
                            9,
                            FontStyle.Bold
                        ),

                    Alignment =
                        DataGridViewContentAlignment.MiddleCenter
                };

            grid.DefaultCellStyle =
                new DataGridViewCellStyle
                {
                    Font =
                        new System.Drawing.Font(
                            "Segoe UI",
                            9
                        ),

                    SelectionBackColor =
                        Color.FromArgb(
                            210,
                            230,
                            220
                        ),

                    SelectionForeColor =
                        Color.Black
                };

            grid.RowTemplate.Height =
                27;

            return grid;
        }


        // =====================================================
        // LOAD ANALYTICS
        // =====================================================

        private void LoadAnalytics()
        {
            try
            {
                if (dtFrom.Value.Date >
                    dtTo.Value.Date)
                {
                    MessageBox.Show(
                        "From Date cannot be greater than To Date.",
                        "Invalid Date",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                    return;
                }


                DateTime fromDate =
                    dtFrom.Value.Date;

                DateTime toDate =
                    dtTo.Value.Date.AddDays(1);


                using (
                    SqlConnection con =
                    DatabaseHelper.GetConnection()
                )
                {
                    con.Open();


                    // =================================================
                    // TOTAL PURCHASE
                    // =================================================

                    string purchaseQuery = @"
                        SELECT ISNULL(
                            SUM(
                                ISNULL(DayTrolley, 0) +
                                ISNULL(NightTrolley, 0)
                            ),
                            0
                        )
                        FROM DailyTrolleyDetails
                        WHERE EntryDate >= @FromDate
                        AND EntryDate < @ToDate
                        AND GroupID IN
                        (
                            SELECT GroupID
                            FROM Groups
                            WHERE GroupType = 'Purchase'
                        )";


                    int purchase =
                        ExecuteDateQuery(
                            con,
                            purchaseQuery,
                            fromDate,
                            toDate
                        );


                    lblPurchase.Text =
                        purchase.ToString();


                    // =================================================
                    // ACTIVE PURCHASE DUMPS
                    // =================================================

                    string activeDumpsQuery = @"
                        SELECT COUNT(*)
                        FROM Groups
                        WHERE IsActive = 1
                        AND GroupType = 'Purchase'";


                    int activeDumps =
                        ExecuteSimpleQuery(
                            con,
                            activeDumpsQuery
                        );


                    lblDumps.Text =
                        activeDumps.ToString();


                    // =================================================
                    // DUMP-WISE PURCHASE
                    // =================================================

                    string dumpWiseQuery = @"
                        SELECT
                            g.GroupName AS DumpName,

                            SUM(
                                ISNULL(d.DayTrolley, 0)
                            ) AS DayTrolleys,

                            SUM(
                                ISNULL(d.NightTrolley, 0)
                            ) AS NightTrolleys,

                            SUM(
                                ISNULL(d.DayTrolley, 0) +
                                ISNULL(d.NightTrolley, 0)
                            ) AS TotalTrolleys

                        FROM DailyTrolleyDetails d

                        INNER JOIN Groups g
                            ON d.GroupID = g.GroupID

                        WHERE
                            d.EntryDate >= @FromDate
                            AND d.EntryDate < @ToDate
                            AND g.GroupType = 'Purchase'

                        GROUP BY
                            g.GroupName

                        ORDER BY
                            TotalTrolleys DESC";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(
                            dumpWiseQuery,
                            con
                        )
                    )
                    {
                        cmd.Parameters.AddWithValue(
                            "@FromDate",
                            fromDate
                        );

                        cmd.Parameters.AddWithValue(
                            "@ToDate",
                            toDate
                        );


                        using (
                            SqlDataAdapter adapter =
                            new SqlDataAdapter(cmd)
                        )
                        {
                            DataTable table =
                                new DataTable();

                            adapter.Fill(table);

                            dgvDumpWise.DataSource =
                                table;
                        }
                    }


                    // =================================================
                    // DATE-WISE PURCHASE
                    // =================================================

                    string dailyQuery = @"
                        SELECT

                            CAST(
                                d.EntryDate AS DATE
                            ) AS EntryDate,

                            SUM(
                                ISNULL(d.DayTrolley, 0)
                            ) AS DayTrolleys,

                            SUM(
                                ISNULL(d.NightTrolley, 0)
                            ) AS NightTrolleys,

                            SUM(
                                ISNULL(d.DayTrolley, 0) +
                                ISNULL(d.NightTrolley, 0)
                            ) AS TotalTrolleys

                        FROM DailyTrolleyDetails d

                        INNER JOIN Groups g
                            ON d.GroupID = g.GroupID

                        WHERE
                            d.EntryDate >= @FromDate
                            AND d.EntryDate < @ToDate
                            AND g.GroupType = 'Purchase'

                        GROUP BY
                            CAST(d.EntryDate AS DATE)

                        ORDER BY
                            TotalTrolleys DESC";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(
                            dailyQuery,
                            con
                        )
                    )
                    {
                        cmd.Parameters.AddWithValue(
                            "@FromDate",
                            fromDate
                        );

                        cmd.Parameters.AddWithValue(
                            "@ToDate",
                            toDate
                        );


                        using (
                            SqlDataAdapter adapter =
                            new SqlDataAdapter(cmd)
                        )
                        {
                            DataTable table =
                                new DataTable();

                            adapter.Fill(table);

                            dgvDaily.DataSource =
                                table;
                        }
                    }


                    // =================================================
                    // HIGHEST PURCHASE DUMP
                    // =================================================

                    if (dgvDumpWise.Rows.Count > 0)
                    {
                        DataGridViewRow row =
                            dgvDumpWise.Rows[0];


                        string dump =
                            Convert.ToString(
                                row.Cells[
                                    "DumpName"
                                ].Value
                            );


                        string total =
                            Convert.ToString(
                                row.Cells[
                                    "TotalTrolleys"
                                ].Value
                            );


                        lblHighestDump.Text =
                            dump
                            + "\n"
                            + total
                            + " Trolleys";
                    }
                    else
                    {
                        lblHighestDump.Text =
                            "No purchase data";
                    }


                    // =================================================
                    // HIGHEST PURCHASE DATE
                    // =================================================

                    if (dgvDaily.Rows.Count > 0)
                    {
                        DataGridViewRow row =
                            dgvDaily.Rows[0];


                        DateTime highestDate =
                            Convert.ToDateTime(
                                row.Cells[
                                    "EntryDate"
                                ].Value
                            );


                        string total =
                            Convert.ToString(
                                row.Cells[
                                    "TotalTrolleys"
                                ].Value
                            );


                        lblHighestDate.Text =
                            highestDate.ToString(
                                "dd-MMM-yyyy"
                            )
                            + "\n"
                            + total
                            + " Trolleys";
                    }
                    else
                    {
                        lblHighestDate.Text =
                            "No purchase data";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load purchase analytics.\n\n"
                    + ex.Message,
                    "Analytics Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }


        // =====================================================
        // DATE QUERY
        // =====================================================

        private int ExecuteDateQuery(
            SqlConnection con,
            string query,
            DateTime fromDate,
            DateTime toDate)
        {
            using (
                SqlCommand cmd =
                new SqlCommand(
                    query,
                    con
                )
            )
            {
                cmd.Parameters.AddWithValue(
                    "@FromDate",
                    fromDate
                );

                cmd.Parameters.AddWithValue(
                    "@ToDate",
                    toDate
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


                return Convert.ToInt32(
                    result
                );
            }
        }


        // =====================================================
        // SIMPLE QUERY
        // =====================================================

        private int ExecuteSimpleQuery(
            SqlConnection con,
            string query)
        {
            using (
                SqlCommand cmd =
                new SqlCommand(
                    query,
                    con
                )
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


                return Convert.ToInt32(
                    result
                );
            }
        }


        // =====================================================
        // EXPORT PDF
        // =====================================================

        private void BtnExportPdf_Click(
            object sender,
            EventArgs e)
        {
            if (
                dgvDumpWise.DataSource == null ||
                dgvDaily.DataSource == null
            )
            {
                MessageBox.Show(
                    "Please generate the analytics first."
                );

                return;
            }


            using (
                SaveFileDialog saveDialog =
                new SaveFileDialog()
            )
            {
                saveDialog.Filter =
                    "PDF Files (*.pdf)|*.pdf";

                saveDialog.FileName =
                    "Paddy_Purchase_Analytics_"
                    + dtFrom.Value.ToString(
                        "yyyy-MM-dd"
                    )
                    + "_to_"
                    + dtTo.Value.ToString(
                        "yyyy-MM-dd"
                    )
                    + ".pdf";


                if (
                    saveDialog.ShowDialog()
                    != DialogResult.OK
                )
                {
                    return;
                }


                try
                {
                    Document document =
                        new Document(
                            PageSize.A4.Rotate(),
                            30,
                            30,
                            30,
                            30
                        );


                    PdfWriter.GetInstance(
                        document,
                        new FileStream(
                            saveDialog.FileName,
                            FileMode.Create
                        )
                    );


                    document.Open();


                    // =================================================
                    // TITLE
                    // =================================================

                    iTextSharp.text.Font titleFont =
                        FontFactory.GetFont(
                            FontFactory.HELVETICA_BOLD,
                            20
                        );


                    Paragraph title =
                        new Paragraph(
                            "RANA GROUP",
                            titleFont
                        );

                    title.Alignment =
                        Element.ALIGN_CENTER;

                    document.Add(title);


                    Paragraph subtitle =
                        new Paragraph(
                            "PADDY PURCHASE MIS & ANALYTICS",
                            FontFactory.GetFont(
                                FontFactory.HELVETICA_BOLD,
                                14
                            )
                        );

                    subtitle.Alignment =
                        Element.ALIGN_CENTER;

                    document.Add(subtitle);


                    document.Add(
                        new Paragraph(" ")
                    );


                    // =================================================
                    // PERIOD
                    // =================================================

                    Paragraph period =
                        new Paragraph(
                            "Purchase Report Period: "
                            + dtFrom.Value.ToString(
                                "dd-MMM-yyyy"
                            )
                            + " to "
                            + dtTo.Value.ToString(
                                "dd-MMM-yyyy"
                            ),
                            FontFactory.GetFont(
                                FontFactory.HELVETICA,
                                10
                            )
                        );

                    period.Alignment =
                        Element.ALIGN_CENTER;

                    document.Add(period);


                    document.Add(
                        new Paragraph(" ")
                    );


                    // =================================================
                    // SUMMARY
                    // =================================================

                    PdfPTable summary =
                        new PdfPTable(2);

                    summary.WidthPercentage =
                        100;


                    AddPdfCell(
                        summary,
                        "TOTAL PURCHASE"
                    );

                    AddPdfCell(
                        summary,
                        "ACTIVE PURCHASE DUMPS"
                    );


                    AddPdfCell(
                        summary,
                        lblPurchase.Text
                    );

                    AddPdfCell(
                        summary,
                        lblDumps.Text
                    );


                    document.Add(summary);


                    document.Add(
                        new Paragraph(" ")
                    );


                    // =================================================
                    // HIGHEST
                    // =================================================

                    document.Add(
                        new Paragraph(
                            "HIGHEST PURCHASE DUMP: "
                            + lblHighestDump.Text.Replace(
                                "\n",
                                " - "
                            ),
                            FontFactory.GetFont(
                                FontFactory.HELVETICA_BOLD,
                                11
                            )
                        )
                    );


                    document.Add(
                        new Paragraph(
                            "HIGHEST PURCHASE DATE: "
                            + lblHighestDate.Text.Replace(
                                "\n",
                                " - "
                            ),
                            FontFactory.GetFont(
                                FontFactory.HELVETICA_BOLD,
                                11
                            )
                        )
                    );


                    document.Add(
                        new Paragraph(" ")
                    );


                    // =================================================
                    // DUMP-WISE
                    // =================================================

                    AddGridToPdf(
                        document,
                        "DUMP-WISE PURCHASE ANALYSIS",
                        dgvDumpWise
                    );


                    // =================================================
                    // DATE-WISE
                    // =================================================

                    AddGridToPdf(
                        document,
                        "DATE-WISE PURCHASE ANALYSIS",
                        dgvDaily
                    );


                    document.Close();


                    MessageBox.Show(
                        "Purchase PDF exported successfully.",
                        "Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Unable to export PDF.\n\n"
                        + ex.Message,
                        "PDF Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }


        // =====================================================
        // PDF CELL
        // =====================================================

        private void AddPdfCell(
            PdfPTable table,
            string text)
        {
            PdfPCell cell =
                new PdfPCell(
                    new Phrase(
                        text,
                        FontFactory.GetFont(
                            FontFactory.HELVETICA_BOLD,
                            9
                        )
                    )
                );

            cell.HorizontalAlignment =
                Element.ALIGN_CENTER;

            cell.Padding = 6;

            table.AddCell(cell);
        }


        // =====================================================
        // GRID TO PDF
        // =====================================================

        private void AddGridToPdf(
            Document document,
            string title,
            DataGridView grid)
        {
            document.Add(
                new Paragraph(
                    title,
                    FontFactory.GetFont(
                        FontFactory.HELVETICA_BOLD,
                        12
                    )
                )
            );


            document.Add(
                new Paragraph(" ")
            );


            if (grid.Columns.Count == 0)
            {
                document.Add(
                    new Paragraph(
                        "No data available."
                    )
                );

                return;
            }


            PdfPTable pdfTable =
                new PdfPTable(
                    grid.Columns.Count
                );

            pdfTable.WidthPercentage =
                100;


            // =================================================
            // HEADERS
            // =================================================

            foreach (
                DataGridViewColumn column
                in grid.Columns
            )
            {
                PdfPCell headerCell =
                    new PdfPCell(
                        new Phrase(
                            column.HeaderText,
                            FontFactory.GetFont(
                                FontFactory.HELVETICA_BOLD,
                                8
                            )
                        )
                    );

                headerCell.HorizontalAlignment =
                    Element.ALIGN_CENTER;

                headerCell.Padding = 5;

                pdfTable.AddCell(
                    headerCell
                );
            }


            // =================================================
            // DATA
            // =================================================

            foreach (
                DataGridViewRow row
                in grid.Rows
            )
            {
                if (row.IsNewRow)
                    continue;


                foreach (
                    DataGridViewCell cell
                    in row.Cells
                )
                {
                    string value =
                        cell.Value == null ||
                        cell.Value == DBNull.Value
                        ? ""
                        : cell.Value.ToString();


                    PdfPCell dataCell =
                        new PdfPCell(
                            new Phrase(
                                value,
                                FontFactory.GetFont(
                                    FontFactory.HELVETICA,
                                    8
                                )
                            )
                        );

                    dataCell.Padding = 4;

                    pdfTable.AddCell(
                        dataCell
                    );
                }
            }


            document.Add(
                pdfTable
            );


            document.Add(
                new Paragraph(" ")
            );
        }


        // =====================================================
        // EXPORT EXCEL
        // =====================================================

        private void BtnExportExcel_Click(
            object sender,
            EventArgs e)
        {
            if (
                dgvDumpWise.DataSource == null ||
                dgvDaily.DataSource == null
            )
            {
                MessageBox.Show(
                    "Please generate the analytics first."
                );

                return;
            }


            using (
                SaveFileDialog saveDialog =
                new SaveFileDialog()
            )
            {
                saveDialog.Filter =
                    "Excel Files (*.xlsx)|*.xlsx";

                saveDialog.FileName =
                    "Paddy_Purchase_Analytics_"
                    + dtFrom.Value.ToString(
                        "yyyy-MM-dd"
                    )
                    + "_to_"
                    + dtTo.Value.ToString(
                        "yyyy-MM-dd"
                    )
                    + ".xlsx";


                if (
                    saveDialog.ShowDialog()
                    != DialogResult.OK
                )
                {
                    return;
                }


                Excel.Application excelApp =
                    null;

                Excel.Workbook workbook =
                    null;

                Excel.Worksheet summarySheet =
                    null;

                Excel.Worksheet dumpSheet =
                    null;

                Excel.Worksheet dailySheet =
                    null;


                try
                {
                    excelApp =
                        new Excel.Application();

                    excelApp.Visible = false;

                    excelApp.DisplayAlerts = false;


                    workbook =
                        excelApp.Workbooks.Add();


                    // =================================================
                    // SUMMARY SHEET
                    // =================================================

                    summarySheet =
                        (Excel.Worksheet)
                        workbook.Worksheets[1];

                    summarySheet.Name =
                        "Purchase Summary";


                    summarySheet.Cells[1, 1] =
                        "RANA GROUP";

                    summarySheet.Cells[2, 1] =
                        "PADDY PURCHASE MIS & ANALYTICS";


                    summarySheet.Cells[4, 1] =
                        "From Date";

                    summarySheet.Cells[4, 2] =
                        dtFrom.Value.ToString(
                            "dd-MMM-yyyy"
                        );


                    summarySheet.Cells[5, 1] =
                        "To Date";

                    summarySheet.Cells[5, 2] =
                        dtTo.Value.ToString(
                            "dd-MMM-yyyy"
                        );


                    summarySheet.Cells[7, 1] =
                        "Total Purchase";

                    summarySheet.Cells[7, 2] =
                        lblPurchase.Text;


                    summarySheet.Cells[8, 1] =
                        "Active Purchase Dumps";

                    summarySheet.Cells[8, 2] =
                        lblDumps.Text;


                    summarySheet.Cells[10, 1] =
                        "Highest Purchase Dump";

                    summarySheet.Cells[10, 2] =
                        lblHighestDump.Text.Replace(
                            "\n",
                            " - "
                        );


                    summarySheet.Cells[11, 1] =
                        "Highest Purchase Date";

                    summarySheet.Cells[11, 2] =
                        lblHighestDate.Text.Replace(
                            "\n",
                            " - "
                        );


                    summarySheet.Columns.AutoFit();


                    // =================================================
                    // DUMP-WISE SHEET
                    // =================================================

                    dumpSheet =
                        (Excel.Worksheet)
                        workbook.Worksheets.Add(
                            After:
                            workbook.Worksheets[
                                workbook.Worksheets.Count
                            ]
                        );

                    dumpSheet.Name =
                        "Dump Wise Purchase";


                    ExportGridToExcel(
                        dgvDumpWise,
                        dumpSheet
                    );


                    // =================================================
                    // DATE-WISE SHEET
                    // =================================================

                    dailySheet =
                        (Excel.Worksheet)
                        workbook.Worksheets.Add(
                            After:
                            workbook.Worksheets[
                                workbook.Worksheets.Count
                            ]
                        );

                    dailySheet.Name =
                        "Date Wise Purchase";


                    ExportGridToExcel(
                        dgvDaily,
                        dailySheet
                    );


                    // =================================================
                    // SAVE
                    // =================================================

                    workbook.SaveAs(
                        saveDialog.FileName,
                        Excel.XlFileFormat.xlOpenXMLWorkbook
                    );


                    workbook.Close(false);

                    excelApp.Quit();


                    MessageBox.Show(
                        "Purchase Excel exported successfully.",
                        "Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Unable to export Excel.\n\n"
                        + ex.Message,
                        "Excel Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
                finally
                {
                    ReleaseExcelObject(
                        dailySheet
                    );

                    ReleaseExcelObject(
                        dumpSheet
                    );

                    ReleaseExcelObject(
                        summarySheet
                    );

                    ReleaseExcelObject(
                        workbook
                    );

                    ReleaseExcelObject(
                        excelApp
                    );
                }
            }
        }


        // =====================================================
        // GRID TO EXCEL
        // =====================================================

        private void ExportGridToExcel(
            DataGridView grid,
            Excel.Worksheet sheet)
        {
            int rowIndex = 1;


            // =================================================
            // HEADERS
            // =================================================

            int columnIndex = 1;


            foreach (
                DataGridViewColumn column
                in grid.Columns
            )
            {
                sheet.Cells[
                    rowIndex,
                    columnIndex
                ] =
                    column.HeaderText;

                columnIndex++;
            }


            rowIndex++;


            // =================================================
            // DATA
            // =================================================

            foreach (
                DataGridViewRow row
                in grid.Rows
            )
            {
                if (row.IsNewRow)
                    continue;


                columnIndex = 1;


                foreach (
                    DataGridViewCell cell
                    in row.Cells
                )
                {
                    object value =
                        cell.Value == null
                        ? ""
                        : cell.Value;


                    sheet.Cells[
                        rowIndex,
                        columnIndex
                    ] =
                        value;

                    columnIndex++;
                }


                rowIndex++;
            }


            sheet.Columns.AutoFit();
        }


        // =====================================================
        // RELEASE EXCEL OBJECT
        // =====================================================

        private void ReleaseExcelObject(
            object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(
                        obj
                    );
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }
}