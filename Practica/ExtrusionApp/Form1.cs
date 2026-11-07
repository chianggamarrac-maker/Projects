using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace ExtrusionApp
{
    public partial class Form1 : Form{
        private readonly string connectionString = "Server=localhost;Database=master;Trusted_Connection=True;Encrypt=False;";
        private TextBox txtSearch;
        private DataGridView dgvOrders;
        private Label lblSelectedOrder;
        private NumericUpDown numGoodRolls;
        private NumericUpDown numNonConformingRolls;
        private TextBox txtDefectNote;
        private Button btnRegister;
        private Label lblTotalGood;
        private Label lblTotalNonConforming;
        private Label lblTotalRemaining;
        private Label lblTotalWeight;
        private string selectedOrderCode = null;
        private int selectedRemainingRolls = 0;
        private string selectedStatus = "";
        public Form1(){
            InitializeComponent();
            SetupCustomUI();
            LoadProductionOrders("");
        }
        private void SetupCustomUI(){
            this.Text = "Extrusion Production Reporting";
            this.Size = new Size(1100, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            // Search
            Label lblSearch = new Label { Text = "Search (Code/Material):", Location = new Point(20, 20), AutoSize = true };
            txtSearch = new TextBox { Location = new Point(170, 18), Width = 300 };
            txtSearch.TextChanged += (s, e) => LoadProductionOrders(txtSearch.Text.Trim());
            // Orders
            dgvOrders = new DataGridView{
                Location = new Point(20, 60),
                Size = new Size(1040, 280),
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvOrders.SelectionChanged += DgvOrders_SelectionChanged;
            // Production Logging Panel
            GroupBox gbReport = new GroupBox { Text = "Register Production Report", Location = new Point(20, 360), Size = new Size(650, 280) };
            lblSelectedOrder = new Label { Text = "Selected Order: None", Location = new Point(20, 30), AutoSize = true, Font = new Font(FontFamily.GenericSansSerif, 9, FontStyle.Bold) };
            Label lblGood = new Label { Text = "Good rolls:", Location = new Point(20, 70), AutoSize = true };
            numGoodRolls = new NumericUpDown { Location = new Point(180, 68), Width = 100, Minimum = 0, Maximum = 1000 };
            Label lblNonConf = new Label { Text = "Non-conforming rolls:", Location = new Point(20, 110), AutoSize = true };
            numNonConformingRolls = new NumericUpDown { Location = new Point(180, 108), Width = 100, Minimum = 0, Maximum = 1000 };
            Label lblNote = new Label { Text = "Defect note (max 250):", Location = new Point(20, 150), AutoSize = true };
            txtDefectNote = new TextBox { Location = new Point(180, 148), Width = 430, Multiline = true, Height = 60, MaxLength = 250 };
            btnRegister = new Button{
                Text = "Register production",
                Location = new Point(180, 225),
                Width = 160,
                Height = 35,
                BackColor = Color.LightGray
            };
            btnRegister.Click += BtnRegister_Click;
            gbReport.Controls.Add(lblSelectedOrder);
            gbReport.Controls.Add(lblGood);
            gbReport.Controls.Add(numGoodRolls);
            gbReport.Controls.Add(lblNonConf);
            gbReport.Controls.Add(numNonConformingRolls);
            gbReport.Controls.Add(lblNote);
            gbReport.Controls.Add(txtDefectNote);
            gbReport.Controls.Add(btnRegister);
            // Visible Totals Panel
            GroupBox gbTotals = new GroupBox { Text = "Visible Totals", Location = new Point(690, 360), Size = new Size(370, 280) };
            lblTotalGood = new Label { Text = "Total Good rolls: 0", Location = new Point(20, 40), AutoSize = true };
            lblTotalNonConforming = new Label { Text = "Total Non-conforming rolls: 0", Location = new Point(20, 80), AutoSize = true };
            lblTotalRemaining = new Label { Text = "Total Remaining good rolls: 0", Location = new Point(20, 120), AutoSize = true };
            lblTotalWeight = new Label { Text = "Total Estimated good weight (kg): 0.000", Location = new Point(20, 160), AutoSize = true };
            gbTotals.Controls.Add(lblTotalGood);
            gbTotals.Controls.Add(lblTotalNonConforming);
            gbTotals.Controls.Add(lblTotalRemaining);
            gbTotals.Controls.Add(lblTotalWeight);
            // Add controls to the window
            this.Controls.Add(lblSearch);
            this.Controls.Add(txtSearch);
            this.Controls.Add(dgvOrders);
            this.Controls.Add(gbReport);
            this.Controls.Add(gbTotals);
        }
        private void LoadProductionOrders(string filter){
            try{
                using (SqlConnection conn = new SqlConnection(connectionString)){
                    conn.Open();
                    string query = @"
                        SELECT 
                            o.Code, 
                            o.Material, 
                            o.WidthMm, 
                            o.RequiredRolls, 
                            ISNULL(SUM(r.GoodRolls), 0) AS GoodRolls,
                            ISNULL(SUM(r.NonConformingRolls), 0) AS NonConformingRolls,
                            (o.RequiredRolls - ISNULL(SUM(r.GoodRolls), 0)) AS RemainingGoodRolls,
                            CAST((ISNULL(SUM(r.GoodRolls), 0) * o.NominalRollWeightKg) AS DECIMAL(18,3)) AS EstimatedGoodWeightKg,
                            CASE 
                                WHEN ISNULL(SUM(r.GoodRolls), 0) >= o.RequiredRolls THEN 'Completed'
                                WHEN o.IsOnHold = 1 THEN 'On hold'
                                ELSE 'Open'
                            END AS Status,
                            o.NominalRollWeightKg,
                            o.IsOnHold
                        FROM ProductionOrder o
                        LEFT JOIN ProductionReport r ON o.Code = r.OrderCode
                        WHERE (@Filter = '' OR o.Code LIKE '%' + @Filter + '%' OR o.Material LIKE '%' + @Filter + '%')
                        GROUP BY o.Code, o.Material, o.WidthMm, o.RequiredRolls, o.NominalRollWeightKg, o.IsOnHold";

                    using (SqlCommand cmd = new SqlCommand(query, conn)){
                        cmd.Parameters.AddWithValue("@Filter", filter);
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        dgvOrders.DataSource = dt;
                        // Hide technical columns while keeping them accessible
                        if (dgvOrders.Columns["NominalRollWeightKg"] != null) dgvOrders.Columns["NominalRollWeightKg"].Visible = false;
                        if (dgvOrders.Columns["IsOnHold"] != null) dgvOrders.Columns["IsOnHold"].Visible = false;
                        CalculateVisibleTotals(dt);
                    }
                }
            }
            catch (Exception ex){
                MessageBox.Show($"Error loading data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CalculateVisibleTotals(DataTable dt){
            int sumGood = 0;
            int sumNonConf = 0;
            int sumRemaining = 0;
            decimal sumWeight = 0;
            foreach (DataRow row in dt.Rows){
                sumGood += Convert.ToInt32(row["GoodRolls"]);
                sumNonConf += Convert.ToInt32(row["NonConformingRolls"]);
                sumRemaining += Convert.ToInt32(row["RemainingGoodRolls"]);
                sumWeight += Convert.ToDecimal(row["EstimatedGoodWeightKg"]);
            }
            lblTotalGood.Text = $"Total Good rolls: {sumGood}";
            lblTotalNonConforming.Text = $"Total Non-conforming rolls: {sumNonConf}";
            lblTotalRemaining.Text = $"Total Remaining good rolls: {sumRemaining}";
            lblTotalWeight.Text = $"Total Estimated good weight (kg): {sumWeight:F3}";
        }
        private void DgvOrders_SelectionChanged(object sender, EventArgs e){
            if (dgvOrders.CurrentRow != null){
                selectedOrderCode = dgvOrders.CurrentRow.Cells["Code"].Value.ToString();
                selectedRemainingRolls = Convert.ToInt32(dgvOrders.CurrentRow.Cells["RemainingGoodRolls"].Value);
                selectedStatus = dgvOrders.CurrentRow.Cells["Status"].Value.ToString();
                decimal nominalWeight = Convert.ToDecimal(dgvOrders.CurrentRow.Cells["NominalRollWeightKg"].Value);
                decimal width = Convert.ToDecimal(dgvOrders.CurrentRow.Cells["WidthMm"].Value);
                int totalProduced = Convert.ToInt32(dgvOrders.CurrentRow.Cells["GoodRolls"].Value) + 
                                   Convert.ToInt32(dgvOrders.CurrentRow.Cells["NonConformingRolls"].Value);
                decimal estimatedProducedWeight = totalProduced * nominalWeight;

                lblSelectedOrder.Text = $"Selected Order: {selectedOrderCode} | Width: {width}mm | Weight: {nominalWeight}kg | Status: {selectedStatus}";
            }
        }
        private void BtnRegister_Click(object sender, EventArgs e){
            if (string.IsNullOrEmpty(selectedOrderCode)){
                MessageBox.Show("Please select an order first.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // Rule: Accept reports only for open orders
            if (selectedStatus != "Open"){
                MessageBox.Show($"Cannot register reports for an order with status '{selectedStatus}'.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            int goodRolls = (int)numGoodRolls.Value;
            int nonConforming = (int)numNonConformingRolls.Value;
            string note = txtDefectNote.Text.Trim();
            // Rule: Sum of good and non-conforming rolls must be positive
            if (goodRolls + nonConforming <= 0){
                MessageBox.Show("The sum of good and non-conforming rolls must be greater than zero.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // Rule: New good rolls cannot exceed the remaining ones
            if (goodRolls > selectedRemainingRolls){
                MessageBox.Show($"Good rolls ({goodRolls}) cannot exceed remaining good rolls ({selectedRemainingRolls}).", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // Rule: Defect note is required if there are non-conforming rolls
            if (nonConforming > 0 && string.IsNullOrEmpty(note)){
                MessageBox.Show("A defect note is required when reporting non-conforming rolls.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (note.Length > 250){
                MessageBox.Show("Defect note cannot exceed 250 characters.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try{
                using (SqlConnection conn = new SqlConnection(connectionString)){
                    conn.Open();
                    string insertQuery = @"INSERT INTO ProductionReport (OrderCode, GoodRolls, NonConformingRolls, DefectNote, CreatedUtc) 
                                           VALUES (@OrderCode, @GoodRolls, @NonConformingRolls, @DefectNote, GETUTCDATE())";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn)){
                        cmd.Parameters.AddWithValue("@OrderCode", selectedOrderCode);
                        cmd.Parameters.AddWithValue("@GoodRolls", goodRolls);
                        cmd.Parameters.AddWithValue("@NonConformingRolls", nonConforming);
                        cmd.Parameters.AddWithValue("@DefectNote", string.IsNullOrEmpty(note) ? (object)DBNull.Value : note);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Production report registered successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Clear fields only after success
                numGoodRolls.Value = 0;
                numNonConformingRolls.Value = 0;
                txtDefectNote.Clear();
                // Reload data to update the grid and totals
                LoadProductionOrders(txtSearch.Text.Trim());
            }
            catch (Exception ex){
                MessageBox.Show($"Error saving report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}