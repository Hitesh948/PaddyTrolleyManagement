using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

using iTextSharp.text;
using iTextSharp.text.pdf;

// Fix Font ambiguity
using DrawingFont = System.Drawing.Font;

namespace PaddyTrolleyManagement
{
    public partial class DailyReport : Form
    {
        private DateTimePicker dtpReportDate;
        private ComboBox cmbUnit;

        private DataGridView dgvPurchase;
        private DataGridView dgvDispatch;

        private Button btnGenerate;
        private Button btnExportPdf;

        private Label lblPurchase;
        private Label lblDispatch;
        private Label lblGrandTotal;

        public DailyReport()
        {
            InitializeComponent();

            CreateUI();
            LoadUnits();
        }

        // ==========================================
        // CREATE UI
        // ==========================================

        private void CreateUI()
        {
            this.Text = "Rana Group Paddy System - Daily Report";
            this.Size = new Size(1200, 780);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.WhiteSmoke;

            // ==========================================
            // TITLE
            // ==========================================

            Label lblTitle = new Label();

            lblTitle.Text = "TROLLEY DETAILS OF PADDY DUMPS";

            lblTitle.Font = new DrawingFont(
                "Segoe UI",
                20,
                FontStyle.Bold);

            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(350, 25);

            this.Controls.Add(lblTitle);


            // ==========================================
            // REPORT DATE
            // ==========================================

            Label lblReportDate = new Label();

            lblReportDate.Text = "Report Date:";

            lblReportDate.Font = new DrawingFont(
                "Segoe UI",
                10,
                FontStyle.Bold);

            lblReportDate.AutoSize = true;
            lblReportDate.Location = new Point(80, 105);

            this.Controls.Add(lblReportDate);


            dtpReportDate = new DateTimePicker();

            dtpReportDate.Format =
                DateTimePickerFormat.Custom;

            dtpReportDate.CustomFormat =
                "dd/MM/yyyy";

            dtpReportDate.Width = 150;

            dtpReportDate.Location =
                new Point(180, 100);

            // Automatically select yesterday
            dtpReportDate.Value =
                DateTime.Today.AddDays(-1);

            this.Controls.Add(dtpReportDate);


            // ==========================================
            // UNIT
            // ==========================================

            Label lblUnit = new Label();

            lblUnit.Text = "Unit:";

            lblUnit.Font = new DrawingFont(
                "Segoe UI",
                10,
                FontStyle.Bold);

            lblUnit.AutoSize = true;

            lblUnit.Location =
                new Point(380, 105);

            this.Controls.Add(lblUnit);


            cmbUnit = new ComboBox();

            cmbUnit.Width = 180;

            cmbUnit.Location =
                new Point(430, 100);

            cmbUnit.DropDownStyle =
                ComboBoxStyle.DropDownList;

            this.Controls.Add(cmbUnit);


            // ==========================================
            // GENERATE REPORT BUTTON
            // ==========================================

            btnGenerate = new Button();

            btnGenerate.Text =
                "Generate Report";

            btnGenerate.Size =
                new Size(150, 38);

            btnGenerate.Location =
                new Point(650, 96);

            btnGenerate.Font =
                new DrawingFont(
                    "Segoe UI",
                    9,
                    FontStyle.Bold);

            btnGenerate.Click +=
                BtnGenerate_Click;

            this.Controls.Add(btnGenerate);


            // ==========================================
            // EXPORT PDF BUTTON
            // ==========================================

            btnExportPdf = new Button();

            btnExportPdf.Text =
                "Export PDF";

            btnExportPdf.Size =
                new Size(140, 38);

            btnExportPdf.Location =
                new Point(830, 96);

            btnExportPdf.Font =
                new DrawingFont(
                    "Segoe UI",
                    9,
                    FontStyle.Bold);

            btnExportPdf.Click +=
                BtnExportPdf_Click;

            this.Controls.Add(btnExportPdf);


            // ==========================================
            // PURCHASE
            // ==========================================

            lblPurchase = new Label();

            lblPurchase.Text = "PURCHASE";

            lblPurchase.Font =
                new DrawingFont(
                    "Segoe UI",
                    13,
                    FontStyle.Bold);

            lblPurchase.AutoSize = true;

            lblPurchase.Location =
                new Point(40, 175);

            this.Controls.Add(lblPurchase);


            dgvPurchase = CreateGrid();

            dgvPurchase.Location =
                new Point(40, 210);

            dgvPurchase.Width = 1100;
            dgvPurchase.Height = 150;

            this.Controls.Add(dgvPurchase);


            // ==========================================
            // DISPATCH
            // ==========================================

            lblDispatch = new Label();

            lblDispatch.Text = "DISPATCH";

            lblDispatch.Font =
                new DrawingFont(
                    "Segoe UI",
                    13,
                    FontStyle.Bold);

            lblDispatch.AutoSize = true;

            lblDispatch.Location =
                new Point(40, 395);

            this.Controls.Add(lblDispatch);


            dgvDispatch = CreateGrid();

            dgvDispatch.Location =
                new Point(40, 430);

            dgvDispatch.Width = 1100;
            dgvDispatch.Height = 120;

            this.Controls.Add(dgvDispatch);


            // ==========================================
            // GRAND TOTAL
            // ==========================================

            lblGrandTotal = new Label();

            lblGrandTotal.Text =
                "GRAND TOTAL → Day: 0 | Night: 0 | Total: 0";

            lblGrandTotal.Font =
                new DrawingFont(
                    "Segoe UI",
                    13,
                    FontStyle.Bold);

            lblGrandTotal.AutoSize = true;

            lblGrandTotal.Location =
                new Point(40, 600);

            this.Controls.Add(lblGrandTotal);
        }


        // ==========================================
        // CREATE GRID
        // ==========================================

        private DataGridView CreateGrid()
        {
            DataGridView grid =
                new DataGridView();

            grid.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            grid.ReadOnly = true;

            grid.AllowUserToAddRows = false;

            grid.AllowUserToDeleteRows = false;

            grid.RowHeadersVisible = false;

            grid.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;

            grid.BackgroundColor =
                Color.White;

            grid.BorderStyle =
                BorderStyle.FixedSingle;

            grid.AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.AllCells;

            return grid;
        }


        // ==========================================
        // LOAD UNITS
        // ==========================================

        private void LoadUnits()
        {
            try
            {
                using (SqlConnection con =
                    DatabaseHelper.GetConnection())
                {
                    con.Open();

                    DataTable table =
                        new DataTable();

                    table.Columns.Add(
                        "UnitID",
                        typeof(int));

                    table.Columns.Add(
                        "UnitName",
                        typeof(string));


                    DataRow allRow =
                        table.NewRow();

                    allRow["UnitID"] = 0;
                    allRow["UnitName"] = "All Units";

                    table.Rows.Add(allRow);


                    string query = @"
                        SELECT
                            UnitID,
                            UnitName
                        FROM Units
                        WHERE IsActive = 1
                        ORDER BY UnitName";


                    using (SqlDataAdapter adapter =
                        new SqlDataAdapter(query, con))
                    {
                        adapter.Fill(table);
                    }


                    cmbUnit.DataSource = table;

                    cmbUnit.DisplayMember =
                        "UnitName";

                    cmbUnit.ValueMember =
                        "UnitID";

                    cmbUnit.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load Units.\n\n" +
                    ex.Message);
            }
        }


        // ==========================================
        // GENERATE REPORT
        // ==========================================

        private void BtnGenerate_Click(
            object sender,
            EventArgs e)
        {
            LoadReport();
        }


        // ==========================================
        // LOAD REPORT
        // ==========================================

        private void LoadReport()
        {
            try
            {
                using (SqlConnection con =
                    DatabaseHelper.GetConnection())
                {
                    con.Open();


                    string query = @"
                        SELECT
                            u.UnitName AS Unit,

                            g.GroupName AS [Dump Name],

                            ISNULL(
                                SUM(dt.DayTrolley),
                                0
                            ) AS Day,

                            ISNULL(
                                SUM(dt.NightTrolley),
                                0
                            ) AS Night,

                            ISNULL(
                                SUM(
                                    dt.DayTrolley +
                                    dt.NightTrolley
                                ),
                                0
                            ) AS Total

                        FROM DailyTrolleyDetails dt

                        INNER JOIN Groups g
                            ON dt.GroupID = g.GroupID

                        INNER JOIN Units u
                            ON g.UnitID = u.UnitID

                        WHERE
                            CAST(dt.EntryDate AS DATE)
                            = @ReportDate

                            AND g.GroupType =
                            @GroupType

                            AND
                            (
                                @UnitID = 0
                                OR g.UnitID = @UnitID
                            )

                        GROUP BY
                            u.UnitName,
                            g.GroupName

                        ORDER BY
                            u.UnitName,
                            g.GroupName";


                    // PURCHASE DATA

                    DataTable purchaseTable =
                        GetReportData(
                            con,
                            query,
                            "Purchase",
                            "Total Purchase Trolley");

                    dgvPurchase.DataSource =
                        purchaseTable;


                    // DISPATCH DATA

                    DataTable dispatchTable =
                        GetReportData(
                            con,
                            query,
                            "Dispatch",
                            "Total Dispatch Trolley");

                    dgvDispatch.DataSource =
                        dispatchTable;


                    // ADJUST PURCHASE GRID

                    AdjustGridHeight(
                        dgvPurchase);


                    // POSITION DISPATCH

                    int dispatchY =
                        dgvPurchase.Bottom + 35;

                    lblDispatch.Location =
                        new Point(
                            40,
                            dispatchY);

                    dgvDispatch.Location =
                        new Point(
                            40,
                            dispatchY + 35);


                    AdjustGridHeight(
                        dgvDispatch);


                    // POSITION GRAND TOTAL

                    lblGrandTotal.Location =
                        new Point(
                            40,
                            dgvDispatch.Bottom + 35);


                    // CALCULATE GRAND TOTAL

                    CalculateGrandTotal(
                        purchaseTable,
                        dispatchTable);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to generate report.\n\n" +
                    ex.Message);
            }
        }


        // ==========================================
        // GET REPORT DATA
        // ==========================================

        private DataTable GetReportData(
            SqlConnection con,
            string query,
            string groupType,
            string totalText)
        {
            using (SqlCommand cmd =
                new SqlCommand(query, con))
            {
                cmd.Parameters.AddWithValue(
                    "@ReportDate",
                    dtpReportDate.Value.Date);

                cmd.Parameters.AddWithValue(
                    "@GroupType",
                    groupType);


                int selectedUnitID = 0;

                if (cmbUnit.SelectedValue != null)
                {
                    selectedUnitID =
                        Convert.ToInt32(
                            cmbUnit.SelectedValue);
                }


                cmd.Parameters.AddWithValue(
                    "@UnitID",
                    selectedUnitID);


                using (SqlDataAdapter adapter =
                    new SqlDataAdapter(cmd))
                {
                    DataTable table =
                        new DataTable();

                    adapter.Fill(table);


                    // ADD SERIAL NUMBER

                    table.Columns.Add(
                        "Sr No",
                        typeof(int));


                    for (
                        int i = 0;
                        i < table.Rows.Count;
                        i++)
                    {
                        table.Rows[i]["Sr No"] =
                            i + 1;
                    }


                    table.Columns["Sr No"]
                        .SetOrdinal(0);


                    // CALCULATE TOTALS

                    int totalDay = 0;
                    int totalNight = 0;


                    foreach (
                        DataRow row
                        in table.Rows)
                    {
                        totalDay +=
                            Convert.ToInt32(
                                row["Day"]);

                        totalNight +=
                            Convert.ToInt32(
                                row["Night"]);
                    }


                    int total =
                        totalDay + totalNight;


                    // ADD TOTAL ROW

                    DataRow totalRow =
                        table.NewRow();

                    totalRow["Sr No"] =
                        DBNull.Value;

                    totalRow["Unit"] =
                        "";

                    totalRow["Dump Name"] =
                        totalText;

                    totalRow["Day"] =
                        totalDay;

                    totalRow["Night"] =
                        totalNight;

                    totalRow["Total"] =
                        total;


                    table.Rows.Add(totalRow);


                    return table;
                }
            }
        }


        // ==========================================
        // ADJUST GRID HEIGHT
        // ==========================================

        private void AdjustGridHeight(
            DataGridView grid)
        {
            int height =
                grid.ColumnHeadersHeight;

            foreach (
                DataGridViewRow row
                in grid.Rows)
            {
                height += row.Height;
            }


            if (height < 70)
            {
                height = 70;
            }

            if (height > 220)
            {
                height = 220;
            }


            grid.Height =
                height + 5;
        }


        // ==========================================
        // CALCULATE GRAND TOTAL
        // ==========================================

        private void CalculateGrandTotal(
            DataTable purchaseTable,
            DataTable dispatchTable)
        {
            int purchaseDay = 0;
            int purchaseNight = 0;
            int purchaseTotal = 0;

            int dispatchDay = 0;
            int dispatchNight = 0;
            int dispatchTotal = 0;


            if (purchaseTable.Rows.Count > 0)
            {
                DataRow row =
                    purchaseTable.Rows[
                        purchaseTable.Rows.Count - 1];

                purchaseDay =
                    Convert.ToInt32(row["Day"]);

                purchaseNight =
                    Convert.ToInt32(row["Night"]);

                purchaseTotal =
                    Convert.ToInt32(row["Total"]);
            }


            if (dispatchTable.Rows.Count > 0)
            {
                DataRow row =
                    dispatchTable.Rows[
                        dispatchTable.Rows.Count - 1];

                dispatchDay =
                    Convert.ToInt32(row["Day"]);

                dispatchNight =
                    Convert.ToInt32(row["Night"]);

                dispatchTotal =
                    Convert.ToInt32(row["Total"]);
            }


            int grandDay =
                purchaseDay + dispatchDay;

            int grandNight =
                purchaseNight + dispatchNight;

            int grandTotal =
                purchaseTotal + dispatchTotal;


            lblGrandTotal.Text =
                $"GRAND TOTAL → " +
                $"Day: {grandDay} | " +
                $"Night: {grandNight} | " +
                $"Total: {grandTotal}";
        }


        // ==========================================
        // EXPORT PDF
        // ==========================================

        private void BtnExportPdf_Click(
            object sender,
            EventArgs e)
        {
            if (
                dgvPurchase.DataSource == null ||
                dgvDispatch.DataSource == null)
            {
                MessageBox.Show(
                    "Please generate the report first.");

                return;
            }


            using (SaveFileDialog saveDialog =
                new SaveFileDialog())
            {
                saveDialog.Filter =
                    "PDF Files (*.pdf)|*.pdf";


                string reportDate =
                    dtpReportDate.Value
                        .ToString("dd-MM-yyyy");


                saveDialog.FileName =
                    "Paddy_Trolley_Report_" +
                    reportDate +
                    ".pdf";


                if (
                    saveDialog.ShowDialog()
                    != DialogResult.OK)
                {
                    return;
                }


                try
                {
                    Document document =
                        new Document(
                            PageSize.A4,
                            30,
                            30,
                            30,
                            30);


                    PdfWriter.GetInstance(
                        document,
                        new FileStream(
                            saveDialog.FileName,
                            FileMode.Create));


                    document.Open();


                    // TITLE

                    Paragraph title =
                        new Paragraph(
                            "TROLLEY DETAILS OF PADDY DUMPS",
                            FontFactory.GetFont(
                                FontFactory.HELVETICA_BOLD,
                                16));

                    title.Alignment =
                        Element.ALIGN_CENTER;

                    document.Add(title);

                    document.Add(
                        new Paragraph(" "));


                    // REPORT INFORMATION

                    Paragraph reportInfo =
                        new Paragraph(
                            "Report Date: " +
                            dtpReportDate.Value
                                .ToString("dd/MM/yyyy") +

                            "\nUnit: " +
                            cmbUnit.Text);


                    reportInfo.Alignment =
                        Element.ALIGN_CENTER;

                    document.Add(reportInfo);

                    document.Add(
                        new Paragraph(" "));


                    // PURCHASE

                    AddPdfSection(
                        document,
                        "PURCHASE",
                        dgvPurchase);


                    // DISPATCH

                    AddPdfSection(
                        document,
                        "DISPATCH",
                        dgvDispatch);


                    document.Add(
                        new Paragraph(" "));


                    // GRAND TOTAL

                    Paragraph grandTotal =
                        new Paragraph(
                            lblGrandTotal.Text,
                            FontFactory.GetFont(
                                FontFactory.HELVETICA_BOLD,
                                12));


                    document.Add(grandTotal);

                    document.Close();


                    MessageBox.Show(
                        "PDF exported successfully.",
                        "Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Unable to export PDF.\n\n" +
                        ex.Message,
                        "PDF Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }


        // ==========================================
        // ADD PDF SECTION
        // ==========================================

        private void AddPdfSection(
            Document document,
            string sectionTitle,
            DataGridView grid)
        {
            Paragraph section =
                new Paragraph(
                    sectionTitle,
                    FontFactory.GetFont(
                        FontFactory.HELVETICA_BOLD,
                        12));


            document.Add(section);


            PdfPTable pdfTable =
                new PdfPTable(
                    grid.ColumnCount);

            pdfTable.WidthPercentage =
                100;


            // HEADERS

            foreach (
                DataGridViewColumn column
                in grid.Columns)
            {
                PdfPCell headerCell =
                    new PdfPCell(
                        new Phrase(
                            column.HeaderText,
                            FontFactory.GetFont(
                                FontFactory.HELVETICA_BOLD,
                                9)));

                pdfTable.AddCell(headerCell);
            }


            // DATA

            foreach (
                DataGridViewRow row
                in grid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }


                foreach (
                    DataGridViewCell cell
                    in row.Cells)
                {
                    string value = "";

                    if (
                        cell.Value != null &&
                        cell.Value != DBNull.Value)
                    {
                        value =
                            cell.Value.ToString();
                    }


                    PdfPCell dataCell =
                        new PdfPCell(
                            new Phrase(
                                value,
                                FontFactory.GetFont(
                                    FontFactory.HELVETICA,
                                    9)));

                    pdfTable.AddCell(dataCell);
                }
            }


            document.Add(pdfTable);

            document.Add(
                new Paragraph(" "));
        }
    }
}