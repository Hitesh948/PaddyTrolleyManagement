using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace PaddyTrolleyManagement
{
    public partial class ManageDumps : Form
    {
        private ComboBox cmbUnit;
        private TextBox txtDumpName;
        private ComboBox cmbType;

        private Button btnAdd;
        private Button btnUpdate;
        private Button btnActivate;
        private Button btnDeactivate;
        private Button btnDelete;
        private Button btnClear;

        private DataGridView dgvDumps;

        private int selectedGroupID = 0;

        public ManageDumps()
        {
            InitializeComponent();
            CreateUI();
            LoadUnits();
            LoadDumps();
        }


        // ============================================
        // CREATE USER INTERFACE
        // ============================================

        private void CreateUI()
        {
            this.Text = "Manage Paddy Dumps";
            this.Size = new Size(1000, 650);
            this.StartPosition = FormStartPosition.CenterScreen;


            // ================= TITLE =================

            Label lblTitle = new Label();

            lblTitle.Text = "MANAGE PADDY DUMPS";
            lblTitle.Font =
                new Font("Segoe UI", 20, FontStyle.Bold);

            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(330, 25);

            this.Controls.Add(lblTitle);


            // ================= UNIT =================

            Label lblUnit = new Label();

            lblUnit.Text = "Unit:";
            lblUnit.Font =
                new Font("Segoe UI", 10, FontStyle.Bold);

            lblUnit.AutoSize = true;
            lblUnit.Location = new Point(40, 100);

            this.Controls.Add(lblUnit);


            cmbUnit = new ComboBox();

            cmbUnit.Location =
                new Point(100, 97);

            cmbUnit.Width = 180;

            cmbUnit.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cmbUnit.SelectedIndexChanged +=
                CmbUnit_SelectedIndexChanged;

            this.Controls.Add(cmbUnit);


            // ================= DUMP NAME =================

            Label lblDumpName = new Label();

            lblDumpName.Text = "Dump Name:";
            lblDumpName.Font =
                new Font("Segoe UI", 10, FontStyle.Bold);

            lblDumpName.AutoSize = true;
            lblDumpName.Location = new Point(320, 100);

            this.Controls.Add(lblDumpName);


            txtDumpName = new TextBox();

            txtDumpName.Location =
                new Point(420, 97);

            txtDumpName.Width = 200;

            this.Controls.Add(txtDumpName);


            // ================= TYPE =================

            Label lblType = new Label();

            lblType.Text = "Type:";
            lblType.Font =
                new Font("Segoe UI", 10, FontStyle.Bold);

            lblType.AutoSize = true;
            lblType.Location = new Point(660, 100);

            this.Controls.Add(lblType);


            cmbType = new ComboBox();

            cmbType.Location =
                new Point(710, 97);

            cmbType.Width = 180;

            cmbType.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cmbType.Items.Add("Purchase");
            cmbType.Items.Add("Dispatch");

            this.Controls.Add(cmbType);


            // ================= BUTTONS =================

            btnAdd = new Button();

            btnAdd.Text = "Add Dump";
            btnAdd.Size = new Size(130, 38);
            btnAdd.Location = new Point(40, 160);

            btnAdd.Click += BtnAdd_Click;

            this.Controls.Add(btnAdd);


            btnUpdate = new Button();

            btnUpdate.Text = "Update Dump";
            btnUpdate.Size = new Size(130, 38);
            btnUpdate.Location = new Point(185, 160);

            btnUpdate.Click += BtnUpdate_Click;

            this.Controls.Add(btnUpdate);


            btnActivate = new Button();

            btnActivate.Text = "Activate";
            btnActivate.Size = new Size(120, 38);
            btnActivate.Location = new Point(330, 160);

            btnActivate.Click += BtnActivate_Click;

            this.Controls.Add(btnActivate);


            btnDeactivate = new Button();

            btnDeactivate.Text = "Deactivate";
            btnDeactivate.Size = new Size(120, 38);
            btnDeactivate.Location = new Point(465, 160);

            btnDeactivate.Click += BtnDeactivate_Click;

            this.Controls.Add(btnDeactivate);


            btnDelete = new Button();

            btnDelete.Text = "Delete";
            btnDelete.Size = new Size(120, 38);
            btnDelete.Location = new Point(600, 160);

            btnDelete.Click += BtnDelete_Click;

            this.Controls.Add(btnDelete);


            btnClear = new Button();

            btnClear.Text = "Clear";
            btnClear.Size = new Size(120, 38);
            btnClear.Location = new Point(735, 160);

            btnClear.Click += BtnClear_Click;

            this.Controls.Add(btnClear);


            // ================= DATA GRID =================

            dgvDumps = new DataGridView();

            dgvDumps.Location =
                new Point(40, 230);

            dgvDumps.Size =
                new Size(850, 320);

            dgvDumps.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            dgvDumps.ReadOnly = true;

            dgvDumps.AllowUserToAddRows = false;

            dgvDumps.AllowUserToDeleteRows = false;

            dgvDumps.RowHeadersVisible = false;

            dgvDumps.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;

            dgvDumps.CellClick +=
                DgvDumps_CellClick;

            this.Controls.Add(dgvDumps);
        }


        // ============================================
        // LOAD UNITS
        // ============================================

        private void LoadUnits()
        {
            try
            {
                using (SqlConnection con =
                    DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        SELECT
                            UnitID,
                            UnitName
                        FROM Units
                        WHERE IsActive = 1
                        ORDER BY UnitName";

                    SqlDataAdapter adapter =
                        new SqlDataAdapter(query, con);

                    DataTable table =
                        new DataTable();

                    adapter.Fill(table);

                    cmbUnit.DataSource =
                        table;

                    cmbUnit.DisplayMember =
                        "UnitName";

                    cmbUnit.ValueMember =
                        "UnitID";

                    cmbUnit.SelectedIndex =
                        -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load Units.\n\n" +
                    ex.Message
                );
            }
        }


        // ============================================
        // LOAD DUMPS
        // ============================================

        private void LoadDumps()
        {
            try
            {
                using (SqlConnection con =
                    DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        SELECT
                            g.GroupID,
                            u.UnitName AS Unit,
                            g.GroupName AS [Dump Name],
                            g.GroupType AS Type,
                            CASE
                                WHEN g.IsActive = 1
                                THEN 'Active'
                                ELSE 'Inactive'
                            END AS Status
                        FROM Groups g
                        INNER JOIN Units u
                            ON g.UnitID = u.UnitID
                        ORDER BY
                            u.UnitName,
                            g.GroupName";

                    SqlDataAdapter adapter =
                        new SqlDataAdapter(
                            query,
                            con
                        );

                    DataTable table =
                        new DataTable();

                    adapter.Fill(table);

                    dgvDumps.DataSource =
                        table;


                    // Hide technical ID
                    if (dgvDumps.Columns["GroupID"] != null)
                    {
                        dgvDumps.Columns["GroupID"].Visible =
                            false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load Dumps.\n\n" +
                    ex.Message
                );
            }
        }


        // ============================================
        // FILTER BY UNIT
        // ============================================

        private void CmbUnit_SelectedIndexChanged(
            object sender,
            EventArgs e)
        {
            if (cmbUnit.SelectedValue == null)
                return;

            if (!(cmbUnit.SelectedValue is int))
                return;

            LoadDumpsByUnit(
                Convert.ToInt32(
                    cmbUnit.SelectedValue
                )
            );
        }


        private void LoadDumpsByUnit(int unitID)
        {
            try
            {
                using (SqlConnection con =
                    DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        SELECT
                            g.GroupID,
                            u.UnitName AS Unit,
                            g.GroupName AS [Dump Name],
                            g.GroupType AS Type,
                            CASE
                                WHEN g.IsActive = 1
                                THEN 'Active'
                                ELSE 'Inactive'
                            END AS Status
                        FROM Groups g
                        INNER JOIN Units u
                            ON g.UnitID = u.UnitID
                        WHERE g.UnitID = @UnitID
                        ORDER BY g.GroupName";

                    SqlCommand cmd =
                        new SqlCommand(
                            query,
                            con
                        );

                    cmd.Parameters.AddWithValue(
                        "@UnitID",
                        unitID
                    );

                    SqlDataAdapter adapter =
                        new SqlDataAdapter(cmd);

                    DataTable table =
                        new DataTable();

                    adapter.Fill(table);

                    dgvDumps.DataSource =
                        table;

                    if (dgvDumps.Columns["GroupID"] != null)
                    {
                        dgvDumps.Columns["GroupID"].Visible =
                            false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message
                );
            }
        }


        // ============================================
        // ADD NEW DUMP
        // ============================================

        private void BtnAdd_Click(
            object sender,
            EventArgs e)
        {
            if (cmbUnit.SelectedIndex == -1)
            {
                MessageBox.Show(
                    "Please select a Unit."
                );

                return;
            }

            if (string.IsNullOrWhiteSpace(
                txtDumpName.Text))
            {
                MessageBox.Show(
                    "Please enter Dump Name."
                );

                return;
            }

            if (cmbType.SelectedIndex == -1)
            {
                MessageBox.Show(
                    "Please select Purchase or Dispatch."
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
                        INSERT INTO Groups
                        (
                            UnitID,
                            GroupName,
                            GroupType,
                            IsActive
                        )
                        VALUES
                        (
                            @UnitID,
                            @GroupName,
                            @GroupType,
                            1
                        )";

                    SqlCommand cmd =
                        new SqlCommand(
                            query,
                            con
                        );

                    cmd.Parameters.AddWithValue(
                        "@UnitID",
                        cmbUnit.SelectedValue
                    );

                    cmd.Parameters.AddWithValue(
                        "@GroupName",
                        txtDumpName.Text.Trim()
                    );

                    cmd.Parameters.AddWithValue(
                        "@GroupType",
                        cmbType.Text
                    );

                    cmd.ExecuteNonQuery();

                    MessageBox.Show(
                        "Dump added successfully."
                    );

                    ClearFields();
                    LoadDumps();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to add Dump.\n\n" +
                    ex.Message
                );
            }
        }


        // ============================================
        // SELECT DUMP FROM GRID
        // ============================================

        private void DgvDumps_CellClick(
            object sender,
            DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            DataGridViewRow row =
                dgvDumps.Rows[e.RowIndex];

            selectedGroupID =
                Convert.ToInt32(
                    row.Cells["GroupID"].Value
                );

            txtDumpName.Text =
                row.Cells["Dump Name"].Value
                .ToString();

            cmbType.Text =
                row.Cells["Type"].Value
                .ToString();

            string unitName =
                row.Cells["Unit"].Value
                .ToString();

            cmbUnit.Text =
                unitName;
        }


        // ============================================
        // UPDATE DUMP
        // ============================================

        private void BtnUpdate_Click(
            object sender,
            EventArgs e)
        {
            if (selectedGroupID == 0)
            {
                MessageBox.Show(
                    "Please select a Dump first."
                );

                return;
            }

            if (string.IsNullOrWhiteSpace(
                txtDumpName.Text))
            {
                MessageBox.Show(
                    "Dump Name cannot be empty."
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
                        UPDATE Groups
                        SET
                            UnitID = @UnitID,
                            GroupName = @GroupName,
                            GroupType = @GroupType
                        WHERE GroupID = @GroupID";

                    SqlCommand cmd =
                        new SqlCommand(
                            query,
                            con
                        );

                    cmd.Parameters.AddWithValue(
                        "@UnitID",
                        cmbUnit.SelectedValue
                    );

                    cmd.Parameters.AddWithValue(
                        "@GroupName",
                        txtDumpName.Text.Trim()
                    );

                    cmd.Parameters.AddWithValue(
                        "@GroupType",
                        cmbType.Text
                    );

                    cmd.Parameters.AddWithValue(
                        "@GroupID",
                        selectedGroupID
                    );

                    cmd.ExecuteNonQuery();

                    MessageBox.Show(
                        "Dump updated successfully."
                    );

                    ClearFields();
                    LoadDumps();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message
                );
            }
        }


        // ============================================
        // ACTIVATE DUMP
        // ============================================

        private void BtnActivate_Click(
            object sender,
            EventArgs e)
        {
            UpdateDumpStatus(1);
        }


        // ============================================
        // DEACTIVATE DUMP
        // ============================================

        private void BtnDeactivate_Click(
            object sender,
            EventArgs e)
        {
            UpdateDumpStatus(0);
        }


        private void UpdateDumpStatus(
            int status)
        {
            if (selectedGroupID == 0)
            {
                MessageBox.Show(
                    "Please select a Dump first."
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
                        UPDATE Groups
                        SET IsActive = @Status
                        WHERE GroupID = @GroupID";

                    SqlCommand cmd =
                        new SqlCommand(
                            query,
                            con
                        );

                    cmd.Parameters.AddWithValue(
                        "@Status",
                        status
                    );

                    cmd.Parameters.AddWithValue(
                        "@GroupID",
                        selectedGroupID
                    );

                    cmd.ExecuteNonQuery();

                    MessageBox.Show(
                        status == 1
                        ? "Dump activated successfully."
                        : "Dump deactivated successfully."
                    );

                    ClearFields();
                    LoadDumps();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message
                );
            }
        }


        // ============================================
        // DELETE DUMP SAFELY
        // ============================================

        private void BtnDelete_Click(
            object sender,
            EventArgs e)
        {
            if (selectedGroupID == 0)
            {
                MessageBox.Show(
                    "Please select a Dump first."
                );

                return;
            }

            DialogResult result =
                MessageBox.Show(
                    "Are you sure you want to delete this Dump?",
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


                    // Check if daily entries exist

                    string checkQuery = @"
                        SELECT COUNT(*)
                        FROM DailyTrolleyDetails
                        WHERE GroupID = @GroupID";

                    SqlCommand checkCmd =
                        new SqlCommand(
                            checkQuery,
                            con
                        );

                    checkCmd.Parameters.AddWithValue(
                        "@GroupID",
                        selectedGroupID
                    );

                    int count =
                        Convert.ToInt32(
                            checkCmd.ExecuteScalar()
                        );


                    if (count > 0)
                    {
                        MessageBox.Show(
                            "This Dump already has daily trolley entries.\n\n" +
                            "It cannot be permanently deleted.\n" +
                            "Please use Deactivate instead."
                        );

                        return;
                    }


                    // Delete if no entries exist

                    string deleteQuery = @"
                        DELETE FROM Groups
                        WHERE GroupID = @GroupID";

                    SqlCommand deleteCmd =
                        new SqlCommand(
                            deleteQuery,
                            con
                        );

                    deleteCmd.Parameters.AddWithValue(
                        "@GroupID",
                        selectedGroupID
                    );

                    deleteCmd.ExecuteNonQuery();

                    MessageBox.Show(
                        "Dump deleted successfully."
                    );

                    ClearFields();
                    LoadDumps();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to delete Dump.\n\n" +
                    ex.Message
                );
            }
        }


        // ============================================
        // CLEAR BUTTON
        // ============================================

        private void BtnClear_Click(
            object sender,
            EventArgs e)
        {
            ClearFields();
            LoadDumps();
        }


        private void ClearFields()
        {
            selectedGroupID = 0;

            txtDumpName.Clear();

            cmbUnit.SelectedIndex = -1;

            cmbType.SelectedIndex = -1;
        }

        private void ManageDumps_Load(object sender, EventArgs e)
        {
 
        }
    }
}