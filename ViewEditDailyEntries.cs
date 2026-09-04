using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace PaddyTrolleyManagement
{
    public partial class ViewEditDailyEntries : Form
    {
        private DataGridView dgvEntries;
        private DateTimePicker dtpFilterDate;
        private Button btnFilter;
        private Button btnShowAll;
        private Button btnEdit;
        private Button btnDelete;
        private Button btnRefresh;

        public ViewEditDailyEntries()
        {
            InitializeComponent();
            CreateUI();
            LoadEntries();
        }

        private void CreateUI()
        {
            // =========================
            // FORM
            // =========================

            this.Text = "Daily Trolley Entries";
            this.Size = new Size(1100, 650);
            this.StartPosition = FormStartPosition.CenterScreen;


            // =========================
            // TITLE
            // =========================

            Label lblTitle = new Label();

            lblTitle.Text = "DAILY TROLLEY ENTRIES";
            lblTitle.Font = new Font(
                "Segoe UI",
                20,
                FontStyle.Bold
            );

            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(30, 25);

            this.Controls.Add(lblTitle);


            // =========================
            // DATE FILTER
            // =========================

            Label lblDate = new Label();

            lblDate.Text = "Date:";
            lblDate.Font = new Font(
                "Segoe UI",
                10,
                FontStyle.Bold
            );

            lblDate.Location = new Point(30, 85);
            lblDate.AutoSize = true;

            this.Controls.Add(lblDate);


            dtpFilterDate = new DateTimePicker();

            dtpFilterDate.Format =
                DateTimePickerFormat.Short;

            dtpFilterDate.Location =
                new Point(80, 80);

            dtpFilterDate.Width = 150;

            this.Controls.Add(dtpFilterDate);


            // =========================
            // FILTER BUTTON
            // =========================

            btnFilter = new Button();

            btnFilter.Text = "Filter";
            btnFilter.Location =
                new Point(250, 78);

            btnFilter.Size =
                new Size(100, 35);

            btnFilter.Click +=
                BtnFilter_Click;

            this.Controls.Add(btnFilter);


            // =========================
            // SHOW ALL BUTTON
            // =========================

            btnShowAll = new Button();

            btnShowAll.Text = "Show All";
            btnShowAll.Location =
                new Point(360, 78);

            btnShowAll.Size =
                new Size(100, 35);

            btnShowAll.Click +=
                BtnShowAll_Click;

            this.Controls.Add(btnShowAll);


            // =========================
            // REFRESH BUTTON
            // =========================

            btnRefresh = new Button();

            btnRefresh.Text = "Refresh";
            btnRefresh.Location =
                new Point(470, 78);

            btnRefresh.Size =
                new Size(100, 35);

            btnRefresh.Click +=
                BtnRefresh_Click;

            this.Controls.Add(btnRefresh);


            // =========================
            // DATA GRID
            // =========================

            dgvEntries = new DataGridView();

            dgvEntries.Name =
                "dgvEntries";

            dgvEntries.Location =
                new Point(30, 135);

            dgvEntries.Size =
                new Size(1020, 380);

            dgvEntries.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            dgvEntries.ReadOnly = true;

            dgvEntries.AllowUserToAddRows =
                false;

            dgvEntries.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;

            dgvEntries.MultiSelect = false;

            this.Controls.Add(dgvEntries);


            // =========================
            // EDIT BUTTON
            // =========================

            btnEdit = new Button();

            btnEdit.Text = "Edit Selected";
            btnEdit.Location =
                new Point(30, 540);

            btnEdit.Size =
                new Size(140, 40);

            btnEdit.Click +=
                BtnEdit_Click;

            this.Controls.Add(btnEdit);


            // =========================
            // DELETE BUTTON
            // =========================

            btnDelete = new Button();

            btnDelete.Text = "Delete Selected";
            btnDelete.Location =
                new Point(190, 540);

            btnDelete.Size =
                new Size(140, 40);

            btnDelete.Click +=
                BtnDelete_Click;

            this.Controls.Add(btnDelete);
        }


        // =====================================================
        // LOAD ALL ENTRIES
        // =====================================================

        private void LoadEntries()
        {
            try
            {
                using (SqlConnection con =
                    DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        SELECT
                            dt.EntryID,
                            dt.EntryDate,
                            u.UnitName AS Unit,
                            g.GroupName AS [WhatsApp Group / Dump],
                            g.GroupType AS Type,
                            dt.DayTrolley AS Day,
                            dt.NightTrolley AS Night,
                            dt.DayTrolley + dt.NightTrolley AS Total

                        FROM DailyTrolleyDetails dt

                        INNER JOIN Groups g
                            ON dt.GroupID = g.GroupID

                        INNER JOIN Units u
                            ON g.UnitID = u.UnitID

                        ORDER BY
                            dt.EntryDate DESC,
                            u.UnitName,
                            g.GroupName";

                    using (SqlDataAdapter adapter =
                        new SqlDataAdapter(query, con))
                    {
                        DataTable table =
                            new DataTable();

                        adapter.Fill(table);

                        dgvEntries.DataSource =
                            table;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load entries.\n\n" +
                    ex.Message,
                    "Database Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }


        // =====================================================
        // FILTER BY DATE
        // =====================================================

        private void BtnFilter_Click(
            object sender,
            EventArgs e)
        {
            try
            {
                using (SqlConnection con =
                    DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        SELECT
                            dt.EntryID,
                            dt.EntryDate,
                            u.UnitName AS Unit,
                            g.GroupName AS [WhatsApp Group / Dump],
                            g.GroupType AS Type,
                            dt.DayTrolley AS Day,
                            dt.NightTrolley AS Night,
                            dt.DayTrolley + dt.NightTrolley AS Total

                        FROM DailyTrolleyDetails dt

                        INNER JOIN Groups g
                            ON dt.GroupID = g.GroupID

                        INNER JOIN Units u
                            ON g.UnitID = u.UnitID

                        WHERE dt.EntryDate = @EntryDate

                        ORDER BY
                            u.UnitName,
                            g.GroupName";

                    using (SqlCommand cmd =
                        new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue(
                            "@EntryDate",
                            dtpFilterDate.Value.Date
                        );

                        using (SqlDataAdapter adapter =
                            new SqlDataAdapter(cmd))
                        {
                            DataTable table =
                                new DataTable();

                            adapter.Fill(table);

                            dgvEntries.DataSource =
                                table;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to filter entries.\n\n" +
                    ex.Message,
                    "Database Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }


        // =====================================================
        // SHOW ALL
        // =====================================================

        private void BtnShowAll_Click(
            object sender,
            EventArgs e)
        {
            LoadEntries();
        }


        // =====================================================
        // REFRESH
        // =====================================================

        private void BtnRefresh_Click(
            object sender,
            EventArgs e)
        {
            LoadEntries();
        }


        // =====================================================
        // EDIT
        // =====================================================

        private void BtnEdit_Click(
     object sender,
     EventArgs e)
        {
            if (dgvEntries.SelectedRows.Count == 0)
            {
                MessageBox.Show(
                    "Please select an entry first."
                );

                return;
            }

            int entryID = Convert.ToInt32(
                dgvEntries.SelectedRows[0]
                .Cells["EntryID"].Value
            );

            int currentDay = Convert.ToInt32(
                dgvEntries.SelectedRows[0]
                .Cells["Day"].Value
            );

            int currentNight = Convert.ToInt32(
                dgvEntries.SelectedRows[0]
                .Cells["Night"].Value
            );


            // Create Edit Window
            Form editForm = new Form();

            editForm.Text = "Edit Daily Entry";
            editForm.Size = new Size(350, 250);
            editForm.StartPosition =
                FormStartPosition.CenterParent;


            // DAY LABEL
            Label lblDay = new Label();

            lblDay.Text = "Day Trolley:";
            lblDay.Location =
                new Point(30, 30);

            lblDay.AutoSize = true;

            editForm.Controls.Add(lblDay);


            // DAY TEXTBOX
            TextBox txtDay = new TextBox();

            txtDay.Location =
                new Point(150, 27);

            txtDay.Width = 120;

            txtDay.Text =
                currentDay.ToString();

            editForm.Controls.Add(txtDay);


            // NIGHT LABEL
            Label lblNight = new Label();

            lblNight.Text = "Night Trolley:";
            lblNight.Location =
                new Point(30, 80);

            lblNight.AutoSize = true;

            editForm.Controls.Add(lblNight);


            // NIGHT TEXTBOX
            TextBox txtNight = new TextBox();

            txtNight.Location =
                new Point(150, 77);

            txtNight.Width = 120;

            txtNight.Text =
                currentNight.ToString();

            editForm.Controls.Add(txtNight);


            // SAVE BUTTON
            Button btnSave = new Button();

            btnSave.Text = "Save";

            btnSave.Location =
                new Point(70, 140);

            btnSave.Size =
                new Size(90, 35);


            // CANCEL BUTTON
            Button btnCancel = new Button();

            btnCancel.Text = "Cancel";

            btnCancel.Location =
                new Point(180, 140);

            btnCancel.Size =
                new Size(90, 35);

            btnCancel.DialogResult =
                DialogResult.Cancel;

            editForm.CancelButton =
                btnCancel;


            // SAVE BUTTON CLICK
            btnSave.Click += (s, ev) =>
            {
                if (!int.TryParse(
                    txtDay.Text,
                    out int newDay))
                {
                    MessageBox.Show(
                        "Please enter a valid Day trolley number."
                    );

                    return;
                }

                if (!int.TryParse(
                    txtNight.Text,
                    out int newNight))
                {
                    MessageBox.Show(
                        "Please enter a valid Night trolley number."
                    );

                    return;
                }

                if (newDay < 0 || newNight < 0)
                {
                    MessageBox.Show(
                        "Trolley counts cannot be negative."
                    );

                    return;
                }

                try
                {
                    using (SqlConnection con =
                        DatabaseHelper.GetConnection())
                    {
                        con.Open();

                        string query = @"
                    UPDATE DailyTrolleyDetails
                    SET
                        DayTrolley = @DayTrolley,
                        NightTrolley = @NightTrolley
                    WHERE EntryID = @EntryID";

                        using (SqlCommand cmd =
                            new SqlCommand(query, con))
                        {
                            cmd.Parameters.AddWithValue(
                                "@DayTrolley",
                                newDay
                            );

                            cmd.Parameters.AddWithValue(
                                "@NightTrolley",
                                newNight
                            );

                            cmd.Parameters.AddWithValue(
                                "@EntryID",
                                entryID
                            );

                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show(
                        "Entry updated successfully!"
                    );

                    editForm.Close();

                    LoadEntries();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Unable to update entry.\n\n" +
                        ex.Message
                    );
                }
            };


            editForm.Controls.Add(btnSave);
            editForm.Controls.Add(btnCancel);


            editForm.ShowDialog();
        }


        // =====================================================
        // DELETE
        // =====================================================

        private void BtnDelete_Click(
            object sender,
            EventArgs e)
        {
            if (dgvEntries.SelectedRows.Count == 0)
            {
                MessageBox.Show(
                    "Please select an entry first."
                );

                return;
            }

            int entryID =
                Convert.ToInt32(
                    dgvEntries.SelectedRows[0]
                    .Cells["EntryID"]
                    .Value
                );

            DialogResult result =
                MessageBox.Show(
                    "Are you sure you want to delete this entry?",
                    "Confirm Delete",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

            if (result != DialogResult.Yes)
                return;

            try
            {
                using (SqlConnection con =
                    DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        DELETE FROM DailyTrolleyDetails
                        WHERE EntryID = @EntryID";

                    using (SqlCommand cmd =
                        new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue(
                            "@EntryID",
                            entryID
                        );

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show(
                    "Entry deleted successfully!"
                );

                LoadEntries();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to delete entry.\n\n" +
                    ex.Message,
                    "Database Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}