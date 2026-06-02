namespace DoseXPDMTool
{
    partial class DoseX_Point_Dose_Tool
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            TestType = new ComboBox();
            Energy = new ComboBox();
            Results = new ListBox();
            Measurement = new TextBox();
            AutoAccept = new CheckBox();
            Accept = new Button();
            TestLabel = new Label();
            energyLabel = new Label();
            MeasurementLabel = new Label();
            ClearResults = new Button();
            Discover = new Button();
            Device = new TextBox();
            electrometerLabel = new Label();
            statusLabel = new Label();
            Chamber_Depth = new Label();
            Arm = new Button();
            SaveHistory = new Button();
            allEnergies = new CheckBox();
            Delete = new Button();
            Insert = new Button();
            Background = new Button();
            SuspendLayout();
            // 
            // TestType
            // 
            TestType.FormattingEnabled = true;
            TestType.Location = new Point(41, 35);
            TestType.Name = "TestType";
            TestType.Size = new Size(121, 23);
            TestType.TabIndex = 0;
            TestType.SelectedIndexChanged += TestType_SelectedIndexChanged;
            // 
            // Energy
            // 
            Energy.FormattingEnabled = true;
            Energy.Location = new Point(43, 119);
            Energy.Name = "Energy";
            Energy.Size = new Size(119, 23);
            Energy.TabIndex = 1;
            Energy.SelectedIndexChanged += Energy_SelectedIndexChanged;
            // 
            // Results
            // 
            Results.FormattingEnabled = true;
            Results.ItemHeight = 15;
            Results.Location = new Point(43, 176);
            Results.Name = "Results";
            Results.Size = new Size(119, 229);
            Results.TabIndex = 2;
            Results.SelectedIndexChanged += Results_SelectedIndexChanged;
            // 
            // Measurement
            // 
            Measurement.Location = new Point(260, 35);
            Measurement.Name = "Measurement";
            Measurement.Size = new Size(129, 23);
            Measurement.TabIndex = 3;
            Measurement.TextChanged += Measurement_TextChanged;
            Measurement.KeyDown += Measurement_KeyDown;
            // 
            // AutoAccept
            // 
            AutoAccept.AutoSize = true;
            AutoAccept.Location = new Point(261, 110);
            AutoAccept.Name = "AutoAccept";
            AutoAccept.Size = new Size(92, 19);
            AutoAccept.TabIndex = 4;
            AutoAccept.Text = "Auto Accept";
            AutoAccept.UseVisualStyleBackColor = true;
            // 
            // Accept
            // 
            Accept.Location = new Point(260, 81);
            Accept.Name = "Accept";
            Accept.Size = new Size(75, 23);
            Accept.TabIndex = 5;
            Accept.Text = "Accept";
            Accept.UseVisualStyleBackColor = true;
            Accept.Click += Accept_Click;
            // 
            // TestLabel
            // 
            TestLabel.AutoSize = true;
            TestLabel.Location = new Point(42, 19);
            TestLabel.Name = "TestLabel";
            TestLabel.Size = new Size(27, 15);
            TestLabel.TabIndex = 6;
            TestLabel.Text = "Test";
            // 
            // energyLabel
            // 
            energyLabel.AutoSize = true;
            energyLabel.Location = new Point(42, 101);
            energyLabel.Name = "energyLabel";
            energyLabel.Size = new Size(43, 15);
            energyLabel.TabIndex = 7;
            energyLabel.Text = "Energy";
            // 
            // MeasurementLabel
            // 
            MeasurementLabel.AutoSize = true;
            MeasurementLabel.Location = new Point(261, 19);
            MeasurementLabel.Name = "MeasurementLabel";
            MeasurementLabel.Size = new Size(104, 15);
            MeasurementLabel.TabIndex = 8;
            MeasurementLabel.Text = "Live Measurement";
            // 
            // ClearResults
            // 
            ClearResults.Location = new Point(42, 440);
            ClearResults.Name = "ClearResults";
            ClearResults.Size = new Size(120, 23);
            ClearResults.TabIndex = 9;
            ClearResults.Text = "Clear Results";
            ClearResults.UseVisualStyleBackColor = true;
            ClearResults.Click += ClearResults_Click;
            // 
            // Discover
            // 
            Discover.Location = new Point(260, 211);
            Discover.Name = "Discover";
            Discover.Size = new Size(76, 23);
            Discover.TabIndex = 10;
            Discover.Text = "Connect";
            Discover.UseVisualStyleBackColor = true;
            Discover.Click += Discover_Click;
            // 
            // Device
            // 
            Device.Location = new Point(261, 182);
            Device.Name = "Device";
            Device.Size = new Size(128, 23);
            Device.TabIndex = 11;
            // 
            // electrometerLabel
            // 
            electrometerLabel.AutoSize = true;
            electrometerLabel.Location = new Point(261, 164);
            electrometerLabel.Name = "electrometerLabel";
            electrometerLabel.Size = new Size(74, 15);
            electrometerLabel.TabIndex = 12;
            electrometerLabel.Text = "Electrometer";
            // 
            // statusLabel
            // 
            statusLabel.AutoSize = true;
            statusLabel.Location = new Point(261, 63);
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(0, 15);
            statusLabel.TabIndex = 13;
            // 
            // Chamber_Depth
            // 
            Chamber_Depth.AutoSize = true;
            Chamber_Depth.Location = new Point(42, 63);
            Chamber_Depth.Name = "Chamber_Depth";
            Chamber_Depth.Size = new Size(0, 15);
            Chamber_Depth.TabIndex = 14;
            // 
            // Arm
            // 
            Arm.Location = new Point(260, 281);
            Arm.Name = "Arm";
            Arm.Size = new Size(129, 23);
            Arm.TabIndex = 15;
            Arm.Text = "Arm Electrometer";
            Arm.UseVisualStyleBackColor = true;
            Arm.Click += Arm_Click;
            // 
            // SaveHistory
            // 
            SaveHistory.Location = new Point(260, 440);
            SaveHistory.Name = "SaveHistory";
            SaveHistory.Size = new Size(129, 23);
            SaveHistory.TabIndex = 16;
            SaveHistory.Text = "Save History";
            SaveHistory.UseVisualStyleBackColor = true;
            SaveHistory.Click += SaveHistory_Click;
            // 
            // allEnergies
            // 
            allEnergies.AutoSize = true;
            allEnergies.Location = new Point(43, 148);
            allEnergies.Name = "allEnergies";
            allEnergies.Size = new Size(87, 19);
            allEnergies.TabIndex = 17;
            allEnergies.Text = "All Energies";
            allEnergies.UseVisualStyleBackColor = true;
            allEnergies.CheckedChanged += allEnergies_CheckedChanged;
            // 
            // Delete
            // 
            Delete.Location = new Point(44, 415);
            Delete.Name = "Delete";
            Delete.Size = new Size(58, 23);
            Delete.TabIndex = 19;
            Delete.Text = "Delete Result";
            Delete.UseVisualStyleBackColor = true;
            Delete.Click += Delete_Click;
            // 
            // Insert
            // 
            Insert.Location = new Point(106, 415);
            Insert.Name = "Insert";
            Insert.Size = new Size(56, 23);
            Insert.TabIndex = 20;
            Insert.Text = "Insert";
            Insert.UseVisualStyleBackColor = true;
            Insert.Click += Insert_Click;
            // 
            // Background
            // 
            Background.Location = new Point(260, 310);
            Background.Name = "Background";
            Background.Size = new Size(129, 23);
            Background.TabIndex = 21;
            Background.Text = "Background";
            Background.UseVisualStyleBackColor = true;
            Background.Click += Background_Click;
            // 
            // DoseX_Point_Dose_Tool
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(401, 475);
            Controls.Add(Background);
            Controls.Add(Insert);
            Controls.Add(Delete);
            Controls.Add(allEnergies);
            Controls.Add(SaveHistory);
            Controls.Add(Arm);
            Controls.Add(Chamber_Depth);
            Controls.Add(statusLabel);
            Controls.Add(electrometerLabel);
            Controls.Add(Device);
            Controls.Add(Discover);
            Controls.Add(ClearResults);
            Controls.Add(MeasurementLabel);
            Controls.Add(energyLabel);
            Controls.Add(TestLabel);
            Controls.Add(Accept);
            Controls.Add(AutoAccept);
            Controls.Add(Measurement);
            Controls.Add(Results);
            Controls.Add(Energy);
            Controls.Add(TestType);
            Name = "DoseX_Point_Dose_Tool";
            Text = "DoseX Point Dose Data Aquisition Tool";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ComboBox TestType;
        private ComboBox Energy;
        private ListBox Results;
        private TextBox Measurement;
        private CheckBox AutoAccept;
        private Button Accept;
        private Label TestLabel;
        private Label energyLabel;
        private Label MeasurementLabel;
        private Button ClearResults;
        private Button Discover;
        private TextBox Device;
        private Label electrometerLabel;
        private Label statusLabel;
        private Label Chamber_Depth;
        private Button Arm;
        private Button SaveHistory;
        private CheckBox allEnergies;
        private Button Delete;
        private Button Insert;
        private Button Background;
    }
}