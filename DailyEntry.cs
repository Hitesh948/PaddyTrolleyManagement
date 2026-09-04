using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace PaddyTrolleyManagement
{
    public partial class DailyEntry : Form
    {
        private DateTimePicker dtpDate;
        private ComboBox cmbUnit;
        private ComboBox cmbGroup;
        private Label lblTypeValue;
        private NumericUpDown nudDay;
        private NumericUpDown nudNight;
        private Button btnSave;

        public DailyEntry()
        {
            InitializeComponent();
            CreateUI();
            LoadUnits();
        }

        private void CreateUI()
        {
            // Form settings
            this.Text = "Daily Paddy Entry";
            this.Size = new Size(600, 550);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Title
            Label lblTitle = new Label();
            lblTitle.Text = "DAILY PADDY ENTRY";
            lblTitle.Font = new Font("Segoe UI", 18, FontStyle.Bold);
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(170, 30);
            this.Controls.Add(lblTitle);


            // ---------------- DATE ----------------

            Label lblDate = new Label();
            lblDate.Text = "Date";
            lblDate.Location = new Point(80, 100);
            lblDate.AutoSize = true;
            this.Controls.Add(lblDate);

            dtpDate = new DateTimePicker();
            dtpDate.Name = "dtpDate";
            dtpDate.Format = DateTimePickerFormat.Short;
            dtpDate.Location = new Point(250, 95);
            dtpDate.Width = 250;
            this.Controls.Add(dtpDate);


            // ---------------- UNIT ----------------

            Label lblUnit = new Label();
            lblUnit.Text = "Unit";
            lblUnit.Location = new Point(80, 150);
            lblUnit.AutoSize = true;
            this.Controls.Add(lblUnit);

            cmbUnit = new ComboBox();
            cmbUnit.Name = "cmbUnit";
            cmbUnit.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbUnit.Location = new Point(250, 145);
            cmbUnit.Width = 250;
            cmbUnit.SelectedIndexChanged += CmbUnit_SelectedIndexChanged;
            this.Controls.Add(cmbUnit);


            // ---------------- GROUP ----------------

            Label lblGroup = new Label();
            lblGroup.Text = "WhatsApp Group / Dump";
            lblGroup.Location = new Point(80, 200);
            lblGroup.AutoSize = true;
            this.Controls.Add(lblGroup);

            cmbGroup = new ComboBox();
            cmbGroup.Name = "cmbGroup";
            cmbGroup.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbGroup.Location = new Point(250, 195);
            cmbGroup.Width = 250;
            cmbGroup.SelectedIndexChanged += CmbGroup_SelectedIndexChanged;
            this.Controls.Add(cmbGroup);


            // ---------------- TYPE ----------------

            Label lblType = new Label();
            lblType.Text = "Type";
            lblType.Location = new Point(80, 250);
            lblType.AutoSize = true;
            this.Controls.Add(lblType);

            lblTypeValue = new Label();
            lblTypeValue.Text = "-";
            lblTypeValue.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblTypeValue.Location = new Point(250, 250);
            lblTypeValue.AutoSize = true;
            this.Controls.Add(lblTypeValue);


            // ---------------- DAY ----------------

            Label lblDay = new Label();
            lblDay.Text = "Day Trolley";
            lblDay.Location = new Point(80, 300);
            lblDay.AutoSize = true;
            this.Controls.Add(lblDay);

            nudDay = new NumericUpDown();
            nudDay.Name = "nudDay";
            nudDay.Minimum = 0;
            nudDay.Maximum = 10000;
            nudDay.Location = new Point(250, 295);
            nudDay.Width = 250;
            this.Controls.Add(nudDay);


            // ---------------- NIGHT ----------------

            Label lblNight = new Label();
            lblNight.Text = "Night Trolley";
            lblNight.Location = new Point(80, 350);
            lblNight.AutoSize = true;
            this.Controls.Add(lblNight);

            nudNight = new NumericUpDown();
            nudNight.Name = "nudNight";
            nudNight.Minimum = 0;
            nudNight.Maximum = 10000;
            nudNight.Location = new Point(250, 345);
            nudNight.Width = 250;
            this.Controls.Add(nudNight);


            // ---------------- SAVE ----------------

            btnSave = new Button();
            btnSave.Name = "btnSave";
            btnSave.Text = "SAVE ENTRY";
            btnSave.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnSave.Location = new Point(250, 410);
            btnSave.Size = new Size(250, 45);
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);
        }


        // =====================================================
        // LOAD UNITS
        // =====================================================

        private void LoadUnits()
        {
            try
            {
                using (SqlConnection con = DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        SELECT UnitID, UnitName
                        FROM Units
                        WHERE IsActive = 1
                        ORDER BY UnitName";

                    using (SqlCommand cmd = new SqlCommand(query, con))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        cmbUnit.Items.Clear();

                        while (reader.Read())
                        {
                            cmbUnit.Items.Add(
                                new ComboBoxItem
                                {
                                    ID = Convert.ToInt32(reader["UnitID"]),
                                    Name = reader["UnitName"].ToString()
                                }
                            );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load Units.\n\n" + ex.Message,
                    "Database Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }


        // =====================================================
        // LOAD GROUPS WHEN UNIT CHANGES
        // =====================================================

        private void CmbUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbUnit.SelectedItem == null)
                return;

            ComboBoxItem selectedUnit =
                (ComboBoxItem)cmbUnit.SelectedItem;

            LoadGroups(selectedUnit.ID);
        }


        private void LoadGroups(int unitID)
        {
            try
            {
                using (SqlConnection con = DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        SELECT GroupID, GroupName, GroupType
                        FROM Groups
                        WHERE UnitID = @UnitID
                        AND IsActive = 1
                        ORDER BY GroupName";

                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@UnitID", unitID);

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            cmbGroup.Items.Clear();

                            while (reader.Read())
                            {
                                cmbGroup.Items.Add(
                                    new GroupComboItem
                                    {
                                        ID = Convert.ToInt32(reader["GroupID"]),
                                        Name = reader["GroupName"].ToString(),
                                        Type = reader["GroupType"].ToString()
                                    }
                                );
                            }
                        }
                    }
                }

                lblTypeValue.Text = "-";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to load Groups.\n\n" + ex.Message,
                    "Database Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }


        // =====================================================
        // SHOW PURCHASE / DISPATCH
        // =====================================================

        private void CmbGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbGroup.SelectedItem == null)
                return;

            GroupComboItem selectedGroup =
                (GroupComboItem)cmbGroup.SelectedItem;

            lblTypeValue.Text = selectedGroup.Type;
        }


        // =====================================================
        // SAVE DAILY ENTRY
        // =====================================================

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (cmbUnit.SelectedItem == null)
            {
                MessageBox.Show("Please select a Unit.");
                return;
            }

            if (cmbGroup.SelectedItem == null)
            {
                MessageBox.Show("Please select a WhatsApp Group / Dump.");
                return;
            }

            GroupComboItem selectedGroup =
                (GroupComboItem)cmbGroup.SelectedItem;

            try
            {
                using (SqlConnection con = DatabaseHelper.GetConnection())
                {
                    con.Open();

                    string query = @"
                        INSERT INTO DailyTrolleyDetails
                        (
                            GroupID,
                            EntryDate,
                            DayTrolley,
                            NightTrolley
                        )
                        VALUES
                        (
                            @GroupID,
                            @EntryDate,
                            @DayTrolley,
                            @NightTrolley
                        )";

                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue(
                            "@GroupID",
                            selectedGroup.ID
                        );

                        cmd.Parameters.AddWithValue(
                            "@EntryDate",
                            dtpDate.Value.Date
                        );

                        cmd.Parameters.AddWithValue(
                            "@DayTrolley",
                            Convert.ToInt32(nudDay.Value)
                        );

                        cmd.Parameters.AddWithValue(
                            "@NightTrolley",
                            Convert.ToInt32(nudNight.Value)
                        );

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show(
                    "Daily trolley entry saved successfully!",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                nudDay.Value = 0;
                nudNight.Value = 0;
            }
            catch (SqlException ex)
            {
                if (ex.Number == 2627 || ex.Number == 2601)
                {
                    MessageBox.Show(
                        "An entry already exists for this Group and Date.",
                        "Duplicate Entry",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
                else
                {
                    MessageBox.Show(
                        "Database error:\n\n" + ex.Message,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error:\n\n" + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }


        // =====================================================
        // COMBOBOX ITEM CLASSES
        // =====================================================

        private class ComboBoxItem
        {
            public int ID { get; set; }
            public string Name { get; set; }

            public override string ToString()
            {
                return Name;
            }
        }


        private class GroupComboItem
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public string Type { get; set; }

            public override string ToString()
            {
                return Name;
            }
        }
    }
}