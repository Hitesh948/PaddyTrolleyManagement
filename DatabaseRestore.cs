using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace PaddyTrolleyManagement
{
    public partial class DatabaseRestore : Form
    {
        private TextBox txtBackupPath;
        private string selectedBackupPath = "";

        public DatabaseRestore()
        {
            InitializeComponent();
            CreateUI();
        }

        // =====================================================
        // CREATE UI
        // =====================================================

        private void CreateUI()
        {
            this.Text = "Restore Database";
            this.Size = new Size(700, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.WhiteSmoke;

            // TITLE

            Label lblTitle = new Label();

            lblTitle.Text = "DATABASE RESTORE";

            lblTitle.Font =
                new Font(
                    "Segoe UI",
                    20,
                    FontStyle.Bold
                );

            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(220, 30);

            this.Controls.Add(lblTitle);


            // WARNING

            Label lblWarning = new Label();

            lblWarning.Text =
                "Warning: Restoring will replace the current database data.";

            lblWarning.Font =
                new Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lblWarning.ForeColor = Color.DarkRed;
            lblWarning.AutoSize = true;
            lblWarning.Location = new Point(130, 90);

            this.Controls.Add(lblWarning);


            // BACKUP FILE LABEL

            Label lblBackupFile = new Label();

            lblBackupFile.Text = "Backup File:";

            lblBackupFile.Font =
                new Font(
                    "Segoe UI",
                    10,
                    FontStyle.Bold
                );

            lblBackupFile.AutoSize = true;
            lblBackupFile.Location = new Point(50, 145);

            this.Controls.Add(lblBackupFile);


            // BACKUP PATH TEXTBOX

            txtBackupPath = new TextBox();

            txtBackupPath.Location =
                new Point(50, 175);

            txtBackupPath.Size =
                new Size(470, 30);

            txtBackupPath.ReadOnly = true;

            this.Controls.Add(txtBackupPath);


            // BROWSE BUTTON

            Button btnBrowse = new Button();

            btnBrowse.Text = "BROWSE";

            btnBrowse.Size =
                new Size(100, 30);

            btnBrowse.Location =
                new Point(535, 175);

            btnBrowse.Click += BtnBrowse_Click;

            this.Controls.Add(btnBrowse);


            // RESTORE BUTTON

            Button btnRestore = new Button();

            btnRestore.Text = "RESTORE DATABASE";

            btnRestore.Size =
                new Size(230, 55);

            btnRestore.Location =
                new Point(230, 250);

            btnRestore.Font =
                new Font(
                    "Segoe UI",
                    11,
                    FontStyle.Bold
                );

            btnRestore.Click += BtnRestore_Click;

            this.Controls.Add(btnRestore);
        }


        // =====================================================
        // BROWSE BACKUP FILE
        // =====================================================

        private void BtnBrowse_Click(
            object sender,
            EventArgs e)
        {
            OpenFileDialog openDialog =
                new OpenFileDialog();

            openDialog.Filter =
                "SQL Server Backup File (*.bak)|*.bak";

            if (
                openDialog.ShowDialog() ==
                DialogResult.OK
            )
            {
                selectedBackupPath =
                    openDialog.FileName;

                txtBackupPath.Text =
                    selectedBackupPath;
            }
        }


        // =====================================================
        // RESTORE DATABASE
        // =====================================================

        private void BtnRestore_Click(
            object sender,
            EventArgs e)
        {
            if (
                string.IsNullOrWhiteSpace(
                    selectedBackupPath
                )
            )
            {
                MessageBox.Show(
                    "Please select a backup (.bak) file first.",
                    "No Backup Selected",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return;
            }


            if (!File.Exists(selectedBackupPath))
            {
                MessageBox.Show(
                    "The selected backup file does not exist.",
                    "File Not Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                return;
            }


            DialogResult result =
                MessageBox.Show(
                    "WARNING!\n\n" +
                    "This will replace the current database with the selected backup.\n\n" +
                    "Do you want to continue?",
                    "Confirm Database Restore",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );


            if (result != DialogResult.Yes)
            {
                return;
            }


            try
            {
                this.Cursor = Cursors.WaitCursor;


                // =================================================
                // GET CURRENT CONNECTION INFORMATION
                // =================================================

                string originalConnectionString;

                using (
                    SqlConnection originalConnection =
                    DatabaseHelper.GetConnection()
                )
                {
                    originalConnectionString =
                        originalConnection.ConnectionString;
                }


                SqlConnectionStringBuilder builder =
                    new SqlConnectionStringBuilder(
                        originalConnectionString
                    );


                string databaseName =
                    builder.InitialCatalog;


                if (
                    string.IsNullOrWhiteSpace(
                        databaseName
                    )
                )
                {
                    MessageBox.Show(
                        "Database name could not be found.",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                    return;
                }


                // =================================================
                // CONNECT TO MASTER DATABASE
                // =================================================

                builder.InitialCatalog = "master";

                string masterConnectionString =
                    builder.ConnectionString;


                using (
                    SqlConnection con =
                    new SqlConnection(
                        masterConnectionString
                    )
                )
                {
                    con.Open();


                    // =============================================
                    // READ LOGICAL FILE NAMES FROM BACKUP
                    // =============================================

                    DataTable fileList =
                        new DataTable();


                    string escapedBackupPath =
                        selectedBackupPath.Replace(
                            "'",
                            "''"
                        );


                    string fileListQuery =
                        $"RESTORE FILELISTONLY " +
                        $"FROM DISK = N'{escapedBackupPath}'";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(
                            fileListQuery,
                            con
                        )
                    )
                    {
                        cmd.CommandTimeout = 0;

                        using (
                            SqlDataAdapter adapter =
                            new SqlDataAdapter(cmd)
                        )
                        {
                            adapter.Fill(fileList);
                        }
                    }


                    if (fileList.Rows.Count < 2)
                    {
                        MessageBox.Show(
                            "Invalid backup file.",
                            "Restore Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        return;
                    }


                    string logicalDataName = "";
                    string logicalLogName = "";


                    foreach (
                        DataRow row in fileList.Rows
                    )
                    {
                        string type =
                            row["Type"].ToString();

                        if (type == "D")
                        {
                            logicalDataName =
                                row["LogicalName"].ToString();
                        }
                        else if (type == "L")
                        {
                            logicalLogName =
                                row["LogicalName"].ToString();
                        }
                    }


                    // =============================================
                    // GET SQL SERVER DEFAULT DATA PATH
                    // =============================================

                    string dataPath = "";

                    string dataPathQuery =
                        "SELECT CAST(SERVERPROPERTY(" +
                        "'InstanceDefaultDataPath' " +
                        ") AS NVARCHAR(4000))";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(
                            dataPathQuery,
                            con
                        )
                    )
                    {
                        object value =
                            cmd.ExecuteScalar();

                        if (
                            value != null &&
                            value != DBNull.Value
                        )
                        {
                            dataPath =
                                value.ToString();
                        }
                    }


                    // =============================================
                    // GET SQL SERVER DEFAULT LOG PATH
                    // =============================================

                    string logPath = "";

                    string logPathQuery =
                        "SELECT CAST(SERVERPROPERTY(" +
                        "'InstanceDefaultLogPath' " +
                        ") AS NVARCHAR(4000))";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(
                            logPathQuery,
                            con
                        )
                    )
                    {
                        object value =
                            cmd.ExecuteScalar();

                        if (
                            value != null &&
                            value != DBNull.Value
                        )
                        {
                            logPath =
                                value.ToString();
                        }
                    }


                    // =============================================
                    // CHECK PATHS
                    // =============================================

                    if (
                        string.IsNullOrWhiteSpace(dataPath) ||
                        string.IsNullOrWhiteSpace(logPath)
                    )
                    {
                        MessageBox.Show(
                            "SQL Server data or log path could not be detected.",
                            "Restore Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        return;
                    }


                    string dataFile =
                        Path.Combine(
                            dataPath,
                            databaseName + ".mdf"
                        );


                    string logFile =
                        Path.Combine(
                            logPath,
                            databaseName + "_log.ldf"
                        );


                    // =============================================
                    // CLOSE EXISTING DATABASE CONNECTIONS
                    // =============================================

                    string setSingleUserQuery =
                        $"ALTER DATABASE [{databaseName}] " +
                        "SET SINGLE_USER " +
                        "WITH ROLLBACK IMMEDIATE";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(
                            setSingleUserQuery,
                            con
                        )
                    )
                    {
                        cmd.CommandTimeout = 0;

                        cmd.ExecuteNonQuery();
                    }


                    // =============================================
                    // RESTORE DATABASE
                    // =============================================

                    string restoreQuery =
                        $"RESTORE DATABASE [{databaseName}] " +
                        $"FROM DISK = N'{escapedBackupPath}' " +
                        "WITH REPLACE, " +
                        $"MOVE N'{logicalDataName}' " +
                        $"TO N'{dataFile}', " +
                        $"MOVE N'{logicalLogName}' " +
                        $"TO N'{logFile}'";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(
                            restoreQuery,
                            con
                        )
                    )
                    {
                        cmd.CommandTimeout = 0;

                        cmd.ExecuteNonQuery();
                    }


                    // =============================================
                    // SET DATABASE BACK TO MULTI USER
                    // =============================================

                    string setMultiUserQuery =
                        $"ALTER DATABASE [{databaseName}] " +
                        "SET MULTI_USER";


                    using (
                        SqlCommand cmd =
                        new SqlCommand(
                            setMultiUserQuery,
                            con
                        )
                    )
                    {
                        cmd.CommandTimeout = 0;

                        cmd.ExecuteNonQuery();
                    }
                }


                MessageBox.Show(
                    "Database restored successfully!\n\n" +
                    "Please restart the application.",
                    "Restore Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Database restore failed.\n\n" +
                    ex.Message,
                    "Restore Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
    }
}