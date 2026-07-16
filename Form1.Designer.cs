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
            MainTabs = new TabControl();
            PointDoseTab = new TabPage();
            Background = new Button();
            Insert = new Button();
            Delete = new Button();
            allEnergies = new CheckBox();
            SaveHistory = new Button();
            Arm = new Button();
            Chamber_Depth = new Label();
            statusLabel = new Label();
            electrometerLabel = new Label();
            Device = new TextBox();
            Discover = new Button();
            ClearResults = new Button();
            MeasurementLabel = new Label();
            energyLabel = new Label();
            TestLabel = new Label();
            Accept = new Button();
            AutoAccept = new CheckBox();
            Measurement = new TextBox();
            Results = new ListBox();
            Energy = new ComboBox();
            TestType = new ComboBox();
            Tg51Tab = new TabPage();
            Tg51EnergyCombo = new ComboBox();
            Tg51DepthValue = new Label();
            Tg51DepthCmText = new TextBox();
            Tg51Bias300List = new ListBox();
            Tg51BiasNeg300List = new ListBox();
            Tg51Bias150List = new ListBox();
            Tg51Bias300Label = new Label();
            Tg51BiasNeg300Label = new Label();
            Tg51Bias150Label = new Label();
            Tg51MoveDepthAuto = new CheckBox();
            Tg51ConnectTank = new Button();
            Tg51MoveDepth = new Button();
            Tg51DisconnectTank = new Button();
            Tg51SetTarget = new Button();
            Tg51ClearEnergy = new Button();
            Tg51TestStart = new Button();
            Tg51TestStop = new Button();
            Tg51Background = new Button();
            Tg51DeleteReading = new Button();
            Tg51InsertReading = new Button();
            Tg51OverwriteReading = new Button();
            Tg51ImportXml = new Button();
            Tg51EnergyLabel = new Label();
            Tg51BridgeStatus = new Label();
            Tg51Status = new Label();
            Tg51Preview = new TextBox();
            Tg51BridgeUrl = new TextBox();
            Tg51XmlPath = new TextBox();
            MainTabs.SuspendLayout();
            PointDoseTab.SuspendLayout();
            Tg51Tab.SuspendLayout();
            SuspendLayout();
            // 
            // MainTabs
            // 
            MainTabs.Controls.Add(PointDoseTab);
            MainTabs.Controls.Add(Tg51Tab);
            MainTabs.Dock = DockStyle.Fill;
            MainTabs.Location = new Point(0, 0);
            MainTabs.Margin = new Padding(3, 4, 3, 4);
            MainTabs.Name = "MainTabs";
            MainTabs.SelectedIndex = 0;
            MainTabs.Size = new Size(553, 681);
            MainTabs.TabIndex = 0;
            // 
            // PointDoseTab
            // 
            PointDoseTab.Controls.Add(Background);
            PointDoseTab.Controls.Add(Insert);
            PointDoseTab.Controls.Add(Delete);
            PointDoseTab.Controls.Add(allEnergies);
            PointDoseTab.Controls.Add(SaveHistory);
            PointDoseTab.Controls.Add(Arm);
            PointDoseTab.Controls.Add(Chamber_Depth);
            PointDoseTab.Controls.Add(statusLabel);
            PointDoseTab.Controls.Add(electrometerLabel);
            PointDoseTab.Controls.Add(Device);
            PointDoseTab.Controls.Add(Discover);
            PointDoseTab.Controls.Add(ClearResults);
            PointDoseTab.Controls.Add(MeasurementLabel);
            PointDoseTab.Controls.Add(energyLabel);
            PointDoseTab.Controls.Add(TestLabel);
            PointDoseTab.Controls.Add(Accept);
            PointDoseTab.Controls.Add(AutoAccept);
            PointDoseTab.Controls.Add(Measurement);
            PointDoseTab.Controls.Add(Results);
            PointDoseTab.Controls.Add(Energy);
            PointDoseTab.Controls.Add(TestType);
            PointDoseTab.Location = new Point(4, 29);
            PointDoseTab.Margin = new Padding(3, 4, 3, 4);
            PointDoseTab.Name = "PointDoseTab";
            PointDoseTab.Padding = new Padding(3, 4, 3, 4);
            PointDoseTab.Size = new Size(545, 648);
            PointDoseTab.TabIndex = 0;
            PointDoseTab.Text = "Point Dose";
            PointDoseTab.UseVisualStyleBackColor = true;
            // 
            // Background
            // 
            Background.Location = new Point(297, 413);
            Background.Margin = new Padding(3, 4, 3, 4);
            Background.Name = "Background";
            Background.Size = new Size(147, 31);
            Background.TabIndex = 21;
            Background.Text = "Background";
            Background.UseVisualStyleBackColor = true;
            Background.Click += Background_Click;
            // 
            // Insert
            // 
            Insert.Location = new Point(121, 553);
            Insert.Margin = new Padding(3, 4, 3, 4);
            Insert.Name = "Insert";
            Insert.Size = new Size(64, 31);
            Insert.TabIndex = 20;
            Insert.Text = "Insert";
            Insert.UseVisualStyleBackColor = true;
            Insert.Click += Insert_Click;
            // 
            // Delete
            // 
            Delete.Location = new Point(50, 553);
            Delete.Margin = new Padding(3, 4, 3, 4);
            Delete.Name = "Delete";
            Delete.Size = new Size(66, 31);
            Delete.TabIndex = 19;
            Delete.Text = "Delete Result";
            Delete.UseVisualStyleBackColor = true;
            Delete.Click += Delete_Click;
            // 
            // allEnergies
            // 
            allEnergies.AutoSize = true;
            allEnergies.Location = new Point(49, 197);
            allEnergies.Margin = new Padding(3, 4, 3, 4);
            allEnergies.Name = "allEnergies";
            allEnergies.Size = new Size(109, 24);
            allEnergies.TabIndex = 17;
            allEnergies.Text = "All Energies";
            allEnergies.UseVisualStyleBackColor = true;
            allEnergies.CheckedChanged += allEnergies_CheckedChanged;
            // 
            // SaveHistory
            // 
            SaveHistory.Location = new Point(297, 587);
            SaveHistory.Margin = new Padding(3, 4, 3, 4);
            SaveHistory.Name = "SaveHistory";
            SaveHistory.Size = new Size(147, 31);
            SaveHistory.TabIndex = 16;
            SaveHistory.Text = "Save History";
            SaveHistory.UseVisualStyleBackColor = true;
            SaveHistory.Click += SaveHistory_Click;
            // 
            // Arm
            // 
            Arm.Location = new Point(297, 375);
            Arm.Margin = new Padding(3, 4, 3, 4);
            Arm.Name = "Arm";
            Arm.Size = new Size(147, 31);
            Arm.TabIndex = 15;
            Arm.Text = "Arm Electrometer";
            Arm.UseVisualStyleBackColor = true;
            Arm.Click += Arm_Click;
            // 
            // Chamber_Depth
            // 
            Chamber_Depth.AutoSize = true;
            Chamber_Depth.Location = new Point(48, 84);
            Chamber_Depth.Name = "Chamber_Depth";
            Chamber_Depth.Size = new Size(0, 20);
            Chamber_Depth.TabIndex = 14;
            // 
            // statusLabel
            // 
            statusLabel.AutoSize = true;
            statusLabel.Location = new Point(298, 84);
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(0, 20);
            statusLabel.TabIndex = 13;
            // 
            // electrometerLabel
            // 
            electrometerLabel.AutoSize = true;
            electrometerLabel.Location = new Point(298, 219);
            electrometerLabel.Name = "electrometerLabel";
            electrometerLabel.Size = new Size(94, 20);
            electrometerLabel.TabIndex = 12;
            electrometerLabel.Text = "Electrometer";
            // 
            // Device
            // 
            Device.Location = new Point(298, 243);
            Device.Margin = new Padding(3, 4, 3, 4);
            Device.Name = "Device";
            Device.Size = new Size(146, 27);
            Device.TabIndex = 11;
            // 
            // Discover
            // 
            Discover.Location = new Point(297, 281);
            Discover.Margin = new Padding(3, 4, 3, 4);
            Discover.Name = "Discover";
            Discover.Size = new Size(87, 31);
            Discover.TabIndex = 10;
            Discover.Text = "Connect";
            Discover.UseVisualStyleBackColor = true;
            Discover.Click += Discover_Click;
            // 
            // ClearResults
            // 
            ClearResults.Location = new Point(48, 587);
            ClearResults.Margin = new Padding(3, 4, 3, 4);
            ClearResults.Name = "ClearResults";
            ClearResults.Size = new Size(137, 31);
            ClearResults.TabIndex = 9;
            ClearResults.Text = "Clear Results";
            ClearResults.UseVisualStyleBackColor = true;
            ClearResults.Click += ClearResults_Click;
            // 
            // MeasurementLabel
            // 
            MeasurementLabel.AutoSize = true;
            MeasurementLabel.Location = new Point(298, 25);
            MeasurementLabel.Name = "MeasurementLabel";
            MeasurementLabel.Size = new Size(129, 20);
            MeasurementLabel.TabIndex = 8;
            MeasurementLabel.Text = "Live Measurement";
            // 
            // energyLabel
            // 
            energyLabel.AutoSize = true;
            energyLabel.Location = new Point(48, 135);
            energyLabel.Name = "energyLabel";
            energyLabel.Size = new Size(54, 20);
            energyLabel.TabIndex = 7;
            energyLabel.Text = "Energy";
            // 
            // TestLabel
            // 
            TestLabel.AutoSize = true;
            TestLabel.Location = new Point(48, 25);
            TestLabel.Name = "TestLabel";
            TestLabel.Size = new Size(35, 20);
            TestLabel.TabIndex = 6;
            TestLabel.Text = "Test";
            // 
            // Accept
            // 
            Accept.Location = new Point(297, 108);
            Accept.Margin = new Padding(3, 4, 3, 4);
            Accept.Name = "Accept";
            Accept.Size = new Size(86, 31);
            Accept.TabIndex = 5;
            Accept.Text = "Accept";
            Accept.UseVisualStyleBackColor = true;
            Accept.Click += Accept_Click;
            // 
            // AutoAccept
            // 
            AutoAccept.AutoSize = true;
            AutoAccept.Location = new Point(298, 147);
            AutoAccept.Margin = new Padding(3, 4, 3, 4);
            AutoAccept.Name = "AutoAccept";
            AutoAccept.Size = new Size(113, 24);
            AutoAccept.TabIndex = 4;
            AutoAccept.Text = "Auto Accept";
            AutoAccept.UseVisualStyleBackColor = true;
            // 
            // Measurement
            // 
            Measurement.Location = new Point(297, 47);
            Measurement.Margin = new Padding(3, 4, 3, 4);
            Measurement.Name = "Measurement";
            Measurement.Size = new Size(147, 27);
            Measurement.TabIndex = 3;
            Measurement.TextChanged += Measurement_TextChanged;
            Measurement.KeyDown += Measurement_KeyDown;
            // 
            // Results
            // 
            Results.FormattingEnabled = true;
            Results.ItemHeight = 20;
            Results.Location = new Point(49, 235);
            Results.Margin = new Padding(3, 4, 3, 4);
            Results.Name = "Results";
            Results.Size = new Size(135, 304);
            Results.TabIndex = 2;
            Results.SelectedIndexChanged += Results_SelectedIndexChanged;
            // 
            // Energy
            // 
            Energy.FormattingEnabled = true;
            Energy.Location = new Point(49, 159);
            Energy.Margin = new Padding(3, 4, 3, 4);
            Energy.Name = "Energy";
            Energy.Size = new Size(135, 28);
            Energy.TabIndex = 1;
            Energy.SelectedIndexChanged += Energy_SelectedIndexChanged;
            // 
            // TestType
            // 
            TestType.FormattingEnabled = true;
            TestType.Location = new Point(47, 47);
            TestType.Margin = new Padding(3, 4, 3, 4);
            TestType.Name = "TestType";
            TestType.Size = new Size(138, 28);
            TestType.TabIndex = 0;
            TestType.SelectedIndexChanged += TestType_SelectedIndexChanged;
            // 
            // Tg51Tab
            // 
            Tg51Tab.Controls.Add(Tg51EnergyCombo);
            Tg51Tab.Controls.Add(Tg51DepthValue);
            Tg51Tab.Controls.Add(Tg51DepthCmText);
            Tg51Tab.Controls.Add(Tg51Bias300List);
            Tg51Tab.Controls.Add(Tg51BiasNeg300List);
            Tg51Tab.Controls.Add(Tg51Bias150List);
            Tg51Tab.Controls.Add(Tg51Bias300Label);
            Tg51Tab.Controls.Add(Tg51BiasNeg300Label);
            Tg51Tab.Controls.Add(Tg51Bias150Label);
            Tg51Tab.Controls.Add(Tg51MoveDepthAuto);
            Tg51Tab.Controls.Add(Tg51ConnectTank);
            Tg51Tab.Controls.Add(Tg51MoveDepth);
            Tg51Tab.Controls.Add(Tg51DisconnectTank);
            Tg51Tab.Controls.Add(Tg51SetTarget);
            Tg51Tab.Controls.Add(Tg51ClearEnergy);
            Tg51Tab.Controls.Add(Tg51TestStart);
            Tg51Tab.Controls.Add(Tg51TestStop);
            Tg51Tab.Controls.Add(Tg51Background);
            Tg51Tab.Controls.Add(Tg51DeleteReading);
            Tg51Tab.Controls.Add(Tg51InsertReading);
            Tg51Tab.Controls.Add(Tg51OverwriteReading);
            Tg51Tab.Controls.Add(Tg51ImportXml);
            Tg51Tab.Controls.Add(Tg51EnergyLabel);
            Tg51Tab.Controls.Add(Tg51BridgeStatus);
            Tg51Tab.Controls.Add(Tg51Status);
            Tg51Tab.Controls.Add(Tg51Preview);
            Tg51Tab.Location = new Point(4, 29);
            Tg51Tab.Margin = new Padding(3, 4, 3, 4);
            Tg51Tab.Name = "Tg51Tab";
            Tg51Tab.Padding = new Padding(3, 4, 3, 4);
            Tg51Tab.Size = new Size(545, 648);
            Tg51Tab.TabIndex = 1;
            Tg51Tab.Text = "TG-51";
            Tg51Tab.UseVisualStyleBackColor = true;
            // 
            // Tg51EnergyCombo
            // 
            Tg51EnergyCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            Tg51EnergyCombo.FormattingEnabled = true;
            Tg51EnergyCombo.Location = new Point(48, 47);
            Tg51EnergyCombo.Margin = new Padding(3, 4, 3, 4);
            Tg51EnergyCombo.Name = "Tg51EnergyCombo";
            Tg51EnergyCombo.Size = new Size(138, 28);
            Tg51EnergyCombo.TabIndex = 20;
            Tg51EnergyCombo.SelectedIndexChanged += Tg51EnergyCombo_SelectedIndexChanged;
            Tg51EnergyCombo.SelectionChangeCommitted += Tg51EnergyCombo_SelectedIndexChanged;
            Tg51EnergyCombo.TextChanged += Tg51EnergyCombo_SelectedIndexChanged;
            // 
            // Tg51DepthValue
            // 
            Tg51DepthValue.AutoSize = true;
            Tg51DepthValue.Location = new Point(48, 89);
            Tg51DepthValue.Name = "Tg51DepthValue";
            Tg51DepthValue.Size = new Size(69, 20);
            Tg51DepthValue.TabIndex = 21;
            Tg51DepthValue.Text = "Depth: --";
            // 
            // Tg51DepthCmText
            // 
            Tg51DepthCmText.Location = new Point(48, 117);
            Tg51DepthCmText.Margin = new Padding(3, 4, 3, 4);
            Tg51DepthCmText.Name = "Tg51DepthCmText";
            Tg51DepthCmText.Size = new Size(77, 27);
            Tg51DepthCmText.TabIndex = 22;
            Tg51DepthCmText.Text = "10";
            // 
            // Tg51Bias300List
            // 
            Tg51Bias300List.FormattingEnabled = true;
            Tg51Bias300List.ItemHeight = 20;
            Tg51Bias300List.Location = new Point(48, 233);
            Tg51Bias300List.Margin = new Padding(3, 4, 3, 4);
            Tg51Bias300List.Name = "Tg51Bias300List";
            Tg51Bias300List.Size = new Size(135, 224);
            Tg51Bias300List.TabIndex = 22;
            Tg51Bias300List.Click += Tg51Bias300Bucket_Selected;
            Tg51Bias300List.Enter += Tg51Bias300Bucket_Selected;
            // 
            // Tg51BiasNeg300List
            // 
            Tg51BiasNeg300List.FormattingEnabled = true;
            Tg51BiasNeg300List.ItemHeight = 20;
            Tg51BiasNeg300List.Location = new Point(207, 233);
            Tg51BiasNeg300List.Margin = new Padding(3, 4, 3, 4);
            Tg51BiasNeg300List.Name = "Tg51BiasNeg300List";
            Tg51BiasNeg300List.Size = new Size(135, 224);
            Tg51BiasNeg300List.TabIndex = 23;
            Tg51BiasNeg300List.Click += Tg51BiasNeg300Bucket_Selected;
            Tg51BiasNeg300List.Enter += Tg51BiasNeg300Bucket_Selected;
            // 
            // Tg51Bias150List
            // 
            Tg51Bias150List.FormattingEnabled = true;
            Tg51Bias150List.ItemHeight = 20;
            Tg51Bias150List.Location = new Point(366, 233);
            Tg51Bias150List.Margin = new Padding(3, 4, 3, 4);
            Tg51Bias150List.Name = "Tg51Bias150List";
            Tg51Bias150List.Size = new Size(135, 224);
            Tg51Bias150List.TabIndex = 24;
            Tg51Bias150List.Click += Tg51Bias150Bucket_Selected;
            Tg51Bias150List.Enter += Tg51Bias150Bucket_Selected;
            // 
            // Tg51Bias300Label
            // 
            Tg51Bias300Label.AutoSize = true;
            Tg51Bias300Label.Location = new Point(48, 209);
            Tg51Bias300Label.Name = "Tg51Bias300Label";
            Tg51Bias300Label.Size = new Size(56, 20);
            Tg51Bias300Label.TabIndex = 25;
            Tg51Bias300Label.Text = "+300 V";
            // 
            // Tg51BiasNeg300Label
            // 
            Tg51BiasNeg300Label.AutoSize = true;
            Tg51BiasNeg300Label.Location = new Point(207, 209);
            Tg51BiasNeg300Label.Name = "Tg51BiasNeg300Label";
            Tg51BiasNeg300Label.Size = new Size(52, 20);
            Tg51BiasNeg300Label.TabIndex = 26;
            Tg51BiasNeg300Label.Text = "-300 V";
            // 
            // Tg51Bias150Label
            // 
            Tg51Bias150Label.AutoSize = true;
            Tg51Bias150Label.Location = new Point(366, 209);
            Tg51Bias150Label.Name = "Tg51Bias150Label";
            Tg51Bias150Label.Size = new Size(37, 20);
            Tg51Bias150Label.TabIndex = 27;
            Tg51Bias150Label.Text = "50%";
            // 
            // Tg51MoveDepthAuto
            // 
            Tg51MoveDepthAuto.AutoSize = true;
            Tg51MoveDepthAuto.Location = new Point(207, 51);
            Tg51MoveDepthAuto.Margin = new Padding(3, 4, 3, 4);
            Tg51MoveDepthAuto.Name = "Tg51MoveDepthAuto";
            Tg51MoveDepthAuto.Size = new Size(205, 24);
            Tg51MoveDepthAuto.TabIndex = 28;
            Tg51MoveDepthAuto.Text = "Move depth automatically";
            Tg51MoveDepthAuto.UseVisualStyleBackColor = true;
            // 
            // Tg51ConnectTank
            // 
            Tg51ConnectTank.Location = new Point(207, 84);
            Tg51ConnectTank.Margin = new Padding(3, 4, 3, 4);
            Tg51ConnectTank.Name = "Tg51ConnectTank";
            Tg51ConnectTank.Size = new Size(105, 33);
            Tg51ConnectTank.TabIndex = 29;
            Tg51ConnectTank.Text = "Connect Tank";
            Tg51ConnectTank.UseVisualStyleBackColor = true;
            Tg51ConnectTank.Click += Tg51ConnectTank_Click;
            // 
            // Tg51MoveDepth
            // 
            Tg51MoveDepth.Location = new Point(319, 84);
            Tg51MoveDepth.Margin = new Padding(3, 4, 3, 4);
            Tg51MoveDepth.Name = "Tg51MoveDepth";
            Tg51MoveDepth.Size = new Size(105, 33);
            Tg51MoveDepth.TabIndex = 30;
            Tg51MoveDepth.Text = "Move Depth";
            Tg51MoveDepth.UseVisualStyleBackColor = true;
            Tg51MoveDepth.Click += Tg51MoveDepth_Click;
            // 
            // Tg51DisconnectTank
            // 
            Tg51DisconnectTank.Location = new Point(431, 84);
            Tg51DisconnectTank.Margin = new Padding(3, 4, 3, 4);
            Tg51DisconnectTank.Name = "Tg51DisconnectTank";
            Tg51DisconnectTank.Size = new Size(105, 33);
            Tg51DisconnectTank.TabIndex = 31;
            Tg51DisconnectTank.Text = "Disconnect";
            Tg51DisconnectTank.UseVisualStyleBackColor = true;
            Tg51DisconnectTank.Click += Tg51DisconnectTank_Click;
            // 
            // Tg51SetTarget
            // 
            Tg51SetTarget.Location = new Point(42, 152);
            Tg51SetTarget.Margin = new Padding(3, 4, 3, 4);
            Tg51SetTarget.Name = "Tg51SetTarget";
            Tg51SetTarget.Size = new Size(141, 33);
            Tg51SetTarget.TabIndex = 32;
            Tg51SetTarget.Text = "Set Target Here";
            Tg51SetTarget.UseVisualStyleBackColor = true;
            Tg51SetTarget.Visible = false;
            Tg51SetTarget.Click += Tg51SetTarget_Click;
            // 
            // Tg51ClearEnergy
            // 
            Tg51ClearEnergy.Location = new Point(48, 481);
            Tg51ClearEnergy.Margin = new Padding(3, 4, 3, 4);
            Tg51ClearEnergy.Name = "Tg51ClearEnergy";
            Tg51ClearEnergy.Size = new Size(136, 31);
            Tg51ClearEnergy.TabIndex = 31;
            Tg51ClearEnergy.Text = "Clear Energy";
            Tg51ClearEnergy.UseVisualStyleBackColor = true;
            Tg51ClearEnergy.Click += Tg51ClearEnergy_Click;
            // 
            // Tg51TestStart
            // 
            Tg51TestStart.Location = new Point(207, 135);
            Tg51TestStart.Margin = new Padding(3, 4, 3, 4);
            Tg51TestStart.Name = "Tg51TestStart";
            Tg51TestStart.Size = new Size(105, 33);
            Tg51TestStart.TabIndex = 33;
            Tg51TestStart.Text = "Test Start";
            Tg51TestStart.UseVisualStyleBackColor = true;
            Tg51TestStart.Click += Tg51TestStart_Click;
            // 
            // Tg51TestStop
            // 
            Tg51TestStop.Location = new Point(319, 135);
            Tg51TestStop.Margin = new Padding(3, 4, 3, 4);
            Tg51TestStop.Name = "Tg51TestStop";
            Tg51TestStop.Size = new Size(105, 33);
            Tg51TestStop.TabIndex = 34;
            Tg51TestStop.Text = "Test Stop";
            Tg51TestStop.UseVisualStyleBackColor = true;
            Tg51TestStop.Click += Tg51TestStop_Click;
            // 
            // Tg51Background
            // 
            Tg51Background.Location = new Point(431, 135);
            Tg51Background.Margin = new Padding(3, 4, 3, 4);
            Tg51Background.Name = "Tg51Background";
            Tg51Background.Size = new Size(105, 33);
            Tg51Background.TabIndex = 35;
            Tg51Background.Text = "Background";
            Tg51Background.UseVisualStyleBackColor = true;
            Tg51Background.Click += Tg51Background_Click;
            // 
            // Tg51DeleteReading
            // 
            Tg51DeleteReading.Location = new Point(207, 481);
            Tg51DeleteReading.Margin = new Padding(3, 4, 3, 4);
            Tg51DeleteReading.Name = "Tg51DeleteReading";
            Tg51DeleteReading.Size = new Size(86, 31);
            Tg51DeleteReading.TabIndex = 36;
            Tg51DeleteReading.Text = "Delete";
            Tg51DeleteReading.UseVisualStyleBackColor = true;
            Tg51DeleteReading.Click += Tg51DeleteReading_Click;
            // 
            // Tg51InsertReading
            // 
            Tg51InsertReading.Location = new Point(299, 481);
            Tg51InsertReading.Margin = new Padding(3, 4, 3, 4);
            Tg51InsertReading.Name = "Tg51InsertReading";
            Tg51InsertReading.Size = new Size(86, 31);
            Tg51InsertReading.TabIndex = 37;
            Tg51InsertReading.Text = "Insert";
            Tg51InsertReading.UseVisualStyleBackColor = true;
            Tg51InsertReading.Click += Tg51InsertReading_Click;
            // 
            // Tg51OverwriteReading
            // 
            Tg51OverwriteReading.Location = new Point(392, 481);
            Tg51OverwriteReading.Margin = new Padding(3, 4, 3, 4);
            Tg51OverwriteReading.Name = "Tg51OverwriteReading";
            Tg51OverwriteReading.Size = new Size(110, 31);
            Tg51OverwriteReading.TabIndex = 38;
            Tg51OverwriteReading.Text = "Overwrite";
            Tg51OverwriteReading.UseVisualStyleBackColor = true;
            Tg51OverwriteReading.Click += Tg51OverwriteReading_Click;
            // 
            // Tg51ImportXml
            // 
            Tg51ImportXml.Location = new Point(431, 175);
            Tg51ImportXml.Margin = new Padding(3, 4, 3, 4);
            Tg51ImportXml.Name = "Tg51ImportXml";
            Tg51ImportXml.Size = new Size(105, 33);
            Tg51ImportXml.TabIndex = 39;
            Tg51ImportXml.Text = "Import XML";
            Tg51ImportXml.UseVisualStyleBackColor = true;
            Tg51ImportXml.Click += Tg51ImportXml_Click;
            // 
            // Tg51EnergyLabel
            // 
            Tg51EnergyLabel.AutoSize = true;
            Tg51EnergyLabel.Location = new Point(21, 21);
            Tg51EnergyLabel.Name = "Tg51EnergyLabel";
            Tg51EnergyLabel.Size = new Size(54, 20);
            Tg51EnergyLabel.TabIndex = 15;
            Tg51EnergyLabel.Text = "Energy";
            // 
            // Tg51BridgeStatus
            // 
            Tg51BridgeStatus.AutoSize = true;
            Tg51BridgeStatus.Location = new Point(48, 547);
            Tg51BridgeStatus.Name = "Tg51BridgeStatus";
            Tg51BridgeStatus.Size = new Size(114, 20);
            Tg51BridgeStatus.TabIndex = 10;
            Tg51BridgeStatus.Text = "Bridge not used";
            // 
            // Tg51Status
            // 
            Tg51Status.AutoSize = true;
            Tg51Status.Location = new Point(48, 511);
            Tg51Status.Name = "Tg51Status";
            Tg51Status.Size = new Size(0, 20);
            Tg51Status.TabIndex = 9;
            // 
            // Tg51Preview
            // 
            Tg51Preview.Location = new Point(48, 579);
            Tg51Preview.Multiline = true;
            Tg51Preview.Name = "Tg51Preview";
            Tg51Preview.ReadOnly = true;
            Tg51Preview.ScrollBars = ScrollBars.Vertical;
            Tg51Preview.Size = new Size(453, 61);
            Tg51Preview.TabIndex = 40;
            Tg51Preview.Text = "Preview only - not written to XML";
            // 
            // Tg51BridgeUrl
            // 
            Tg51BridgeUrl.Location = new Point(249, 530);
            Tg51BridgeUrl.Name = "Tg51BridgeUrl";
            Tg51BridgeUrl.Size = new Size(198, 27);
            Tg51BridgeUrl.TabIndex = 12;
            Tg51BridgeUrl.Visible = false;
            // 
            // Tg51XmlPath
            // 
            Tg51XmlPath.Location = new Point(703, 530);
            Tg51XmlPath.Name = "Tg51XmlPath";
            Tg51XmlPath.ReadOnly = true;
            Tg51XmlPath.Size = new Size(247, 27);
            Tg51XmlPath.TabIndex = 14;
            Tg51XmlPath.Visible = false;
            // 
            // DoseX_Point_Dose_Tool
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(553, 681);
            Controls.Add(MainTabs);
            Margin = new Padding(3, 4, 3, 4);
            Name = "DoseX_Point_Dose_Tool";
            Text = "DoseX Point Dose Data Aquisition Tool";
            MainTabs.ResumeLayout(false);
            PointDoseTab.ResumeLayout(false);
            PointDoseTab.PerformLayout();
            Tg51Tab.ResumeLayout(false);
            Tg51Tab.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TabControl MainTabs;
        private TabPage PointDoseTab;
        private TabPage Tg51Tab;
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
        private Label Tg51Status;
        private Label Tg51BridgeStatus;
        private TextBox Tg51BridgeUrl;
        private TextBox Tg51XmlPath;
        private TextBox Tg51Preview;
        private Label Tg51EnergyLabel;
        private ComboBox Tg51EnergyCombo;
        private Label Tg51DepthValue;
        private TextBox Tg51DepthCmText;
        private ListBox Tg51Bias300List;
        private ListBox Tg51BiasNeg300List;
        private ListBox Tg51Bias150List;
        private Label Tg51Bias300Label;
        private Label Tg51BiasNeg300Label;
        private Label Tg51Bias150Label;
        private CheckBox Tg51MoveDepthAuto;
        private Button Tg51ConnectTank;
        private Button Tg51MoveDepth;
        private Button Tg51DisconnectTank;
        private Button Tg51SetTarget;
        private Button Tg51ClearEnergy;
        private Button Tg51TestStart;
        private Button Tg51TestStop;
        private Button Tg51Background;
        private Button Tg51DeleteReading;
        private Button Tg51InsertReading;
        private Button Tg51OverwriteReading;
        private Button Tg51ImportXml;
    }
}
