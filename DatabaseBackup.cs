using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace PaddyTrolleyManagement
{
    public partial class DatabaseBackup : Form
    {
        public DatabaseBackup()
        {
            InitializeComponent();
            CreateUI();
        }

        private void CreateUI()
        {
            this.Text = "Database Backup";
            this.Size = new Size(600, 350);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.WhiteSmoke;


            // TITLE

            Label lblTitle = new Label();

            lblTitle.Text = "DATABASE BACKUP";

            lblTitle.Font =
                new Font(
                    "Segoe UI",
                    20,
                    FontStyle.Bold
                );

            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(170, 30);

            this.Controls.Add(lblTitle);


            // INFORMATION

            Label lblInfo = new Label();

            lblInfo.Text =
                "Create a backup of the Paddy Management System database.";

            lblInfo.Font =
                new Font(
                    "Segoe UI",
                    10,
                    FontStyle.Regular
                );

            lblInfo.AutoSize = true;
            lblInfo.Location = new Point(90, 90);

            this.Controls.Add(lblInfo);


            // BACKUP BUTTON

            Button btnBackup = new Button();

            btnBackup.Text = "CREATE BACKUP";

            btnBackup.Size =
                new Size(220, 55);

            btnBackup.Location =
                new Point(180, 150);

            btnBackup.Font =
                new Font(
                    "Segoe UI",
                    11,
                    FontStyle.Bold
                );

            btnBackup.Click += BtnBackup_Click;

            this.Controls.Add(btnBackup);


            // STATUS LABEL

            Label lblFooter = new Label();

            lblFooter.Text =
                "Backup files will be saved as .bak files.";

            lblFooter.AutoSize = true;

            lblFooter.Location =
                new Point(180, 240);

            this.Controls.Add(lblFooter);
        }


        // =====================================================
        // CREATE DATABASE BACKUP
        // =====================================================

        private void BtnBackup_Click(
            object sender,
            EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog =
                    new SaveFileDialog();

                saveDialog.Filter =
                    "SQL Server Backup File (*.bak)|*.bak";

                saveDialog.FileName =
                    "PaddyManagement_" +
                    DateTime.Now.ToString(
                        "yyyyMMdd_HHmmss"
                    ) +
                    ".bak";


                if (
                    saveDialog.ShowDialog() ==
                    DialogResult.OK
                )
                {
                    string backupPath =
                        saveDialog.FileName;


                    using (
                        SqlConnection con =
                        DatabaseHelper.GetConnection()
                    )
                    {
                        con.Open();


                        string databaseName =
                            con.Database;


                        string query =
                            $"BACKUP DATABASE [{databaseName}] " +
                            $"TO DISK = N'{backupPath}' " +
                            "WITH INIT, FORMAT";


                        using (
                            SqlCommand cmd =
                            new SqlCommand(
                                query,
                                con
                            )
                        )
                        {
                            cmd.CommandTimeout = 0;

                            cmd.ExecuteNonQuery();
                        }
                    }


                    MessageBox.Show(
                        "Database backup created successfully!",
                        "Backup Complete",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Backup failed.\n\n" +
                    ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}