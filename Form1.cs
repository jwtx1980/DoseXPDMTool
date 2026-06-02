namespace DoseXPDMTool
{

    using Excel = Microsoft.Office.Interop.Excel;
    using System.Runtime.InteropServices;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Drawing;
    using System.Windows.Forms;
    using Zeroconf;
    using System.Net.WebSockets;
    using System.Threading;
    using System.Text;
    using System.Collections.Generic;
    using Action = Action;
    using System.Net;
    using Newtonsoft.Json.Linq;
    using System.Diagnostics;
    using System.Text.RegularExpressions;
    using Newtonsoft.Json;
    using Nukepayload2.Interop.Office365.Excel;
    using System.Xml.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;


    public partial class DoseX_Point_Dose_Tool : Form
    {
        Machine machine = new Machine();
        private Excel.Application excelApp;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        private string currentNamedRange;
        private ClientWebSocket wsClient;
        private CancellationTokenSource cancellationTokenSource;
        private string IPAddress;
        private int port = 8083;
        private double lastKnownCharge = 0.0;
        double knowncharge2 = 0.0;
        double knowncharge3 = 0.0;
        private bool isMeasurementRunning = false;
        private bool previousIsCollecting = false;
        private double previousCharge = 0.0;
        string lasthistory = "null";
        private int falseCountAfterTrue = 0;
        private bool hasSeenTrue = false;
        private bool armed = true;
        private bool blocked = true;
        private bool requested = false;
        private bool globalHighVoltageEnabled;
        private int globalBiasVoltage;
        private bool doseXConnected = false;
        private bool excelConnected = false;
        private bool allowAllEnergies = false;
        private Logger logger;
        private TaskCompletionSource<bool> loginTcs;
        private string PW = "VarianAOS1!";

        public class BeamConfiguration
        {
            public double BeamOnTime { get; set; }
            public double BeamOffTime { get; set; }
            public double PreTriggerTime { get; set; }
            public double PostTriggerTime { get; set; }

            public BeamConfiguration(double beamOnTime, double beamOffTime, double preTriggerTime, double postTriggerTime)
            {
                BeamOnTime = beamOnTime;
                BeamOffTime = beamOffTime;
                PreTriggerTime = preTriggerTime;
                PostTriggerTime = postTriggerTime;
            }
        }


        public DoseX_Point_Dose_Tool()
        {
            InitializeComponent();
            PopulateComboBox();
            DiscoverServices();
            UpdateBeamConfiguration(0.1, 0.3, 0.1, 0.1);


            // Initialize machine with default energies in case reading config fails or has no energies
            machine = new Machine()
            {
                Energies = new List<string>
                {
                "4x", "6x", "6xfff", "8x", "10x", "10xfff",
                "15x", "16x", "18x", "20x", "23x", "6e", "9e",
                "12e", "15e", "16e", "18e", "20e", "22e"
                }
            };

            try
            {
                // Define the AppData path and subdirectory for your application
                string appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string appSpecificFolder = Path.Combine(appDataFolder, "AQAFT");

                // Ensure the directory exists
                if (!Directory.Exists(appSpecificFolder))
                {
                    Directory.CreateDirectory(appSpecificFolder);
                }

                // Define the path for MachineConfig.txt in AppData
                string localMachineConfigPath = Path.Combine(appSpecificFolder, "MachineConfig.txt");
                // Load the machine configuration from the AppData MachineConfig.txt
                Machine loadedMachine = ReadMachineConfig(localMachineConfigPath);
                if (loadedMachine.Energies.Any())
                {
                    this.machine = loadedMachine; // Only overwrite if there are any energies defined
                    InitializeExcel();
                }
            }
            catch (Exception ex)
            {
                // Handle errors (log or show error message)
                Console.WriteLine($"Failed to read machine config: {ex.Message}");
                // The machine continues with default energies, no need to reassign.
            }
            // Create a logger instance
            // Create a logger instance
            logger = new Logger(machine.Path);
            ExecuteWithDelay();
        }
        public async Task ExecuteWithDelay()
        {
            // Wait for 3 seconds
            await Task.Delay(1000);

            // Execute the command
            ConnectToWebSocketAsync();
        }

        private async void ConnectToWebSocketAsync()
        {
            wsClient = new ClientWebSocket();
            cancellationTokenSource = new CancellationTokenSource();
            try
            {
                Uri serverUri = new Uri($"wss://{IPAddress}:{port}");
                wsClient.Options.RemoteCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
                await wsClient.ConnectAsync(serverUri, cancellationTokenSource.Token);
                StartReceiving();
                await ReadHighVoltageAndBiasVoltageAsync();
                Discover.Text = "Connected";
                doseXConnected = true;
            }
            catch (WebSocketException wsEx)
            {
                Discover.Text = "Discover";
                MessageBox.Show($"WebSocket Error: {wsEx.Message}");
                Console.WriteLine($"WebSocket Error: {wsEx.Message}");
            }
            catch (Exception ex)
            {
                Discover.Text = "Discover";
                MessageBox.Show($"Error connecting to WebSocket server: {ex.Message}");
                Console.WriteLine($"Error connecting to WebSocket server: {ex.Message}");
            }
        }






        private async void StartReceiving()
        {
            var buffer = new byte[1024];
            while (wsClient.State == WebSocketState.Open)
            {
                try
                {
                    var result = await wsClient.ReceiveAsync(new ArraySegment<byte>(buffer), cancellationTokenSource.Token);
                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        await wsClient.CloseAsync(WebSocketCloseStatus.NormalClosure, string.Empty, CancellationToken.None);
                    }
                    else
                    {
                        var message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                        Console.WriteLine($"Received message: {message}");
                        HandleWebSocketMessage(message);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error receiving message: " + ex.Message);
                }
            }
        }

        private bool previousArmedState = true;
        int armedupdate = 0;
        private void HandleWebSocketMessage(string message)
        {
            try
            {
                Console.WriteLine("Received message: " + message);
                var json = JObject.Parse(message);
                string cmd = json["cmd"]?.ToString();

                switch (cmd)
                {
                    case "measurement_data":
                        HandleMeasurementData(json);
                        break;
                    case "remote_status":
                        HandleRemoteStatus(json);
                        break;
                    case "value_init":
                        HandleConfigurationItemsResponse(json);
                        break;
                    case "get_values":
                        HandleMeasurementHistoryResponse(json, message);
                        break;
                    case "login":
                        HandleAdminLoginResponse(json);
                        break;
                    default:
                        Console.WriteLine("Unknown command received.");
                        break;
                }
            }
            catch (JsonReaderException jsonEx)
            {
                WriteMessageToFile(message, "JsonReaderException");
                Console.WriteLine("JSON error: " + jsonEx.Message);
            }
            catch (Exception ex)
            {
                WriteMessageToFile(message, "GeneralException");
                Console.WriteLine("General error: " + ex.Message);
            }
        }

        private void HandleAdminLoginResponse(JObject json)
        {
            bool result = json["result"]?.ToObject<bool>() ?? false;
            if (!result) { MessageBox.Show($"Failed to Login to admin mode"); }
            Console.WriteLine($"Admin login response received: {result}");

            if (result)
            {
                loginTcs?.SetResult(true);
            }
            else
            {
                loginTcs?.SetResult(false);
            }
        }



        private void HandleRemoteStatus(JObject json)
        {
            var values = json["values"];
            bool controlGranted = values?["control"]?.ToObject<bool>() ?? false;
            bool blocked = values?["blocked"]?.ToObject<bool>() ?? false;

            controlTcs?.TrySetResult(controlGranted && !blocked);

            if (controlGranted && !blocked && requested)
            {
                MessageBox.Show($"Bias Set to 300V\n Aquisition timers set to:\n\t Beam on time = {beamConfig.BeamOnTime}\n\t Beam off time = {beamConfig.BeamOffTime}\n\t Pretrigger time = {beamConfig.PreTriggerTime}\n\t Posttrigger time = {beamConfig.PostTriggerTime}");
                Console.WriteLine("Remote control granted.");
            }
            else if (requested)
            {
                MessageBox.Show("Remote control not granted. Bias Not Set. Please free the electrometer by pressing the cloud at the top of the device screen and try again.");
                Console.WriteLine("Remote control not granted.");
            }
            requested = false;
        }

        private void HandleConfigurationItemsResponse(JObject json)
        {
            var values = json["values"] as JObject;
            if (values != null)
            {
                foreach (var property in values.Properties())
                {
                    Console.WriteLine($"{property.Name}: {property.Value}");
                    if (property.Name == "highVoltageEnabled")
                    {
                        globalHighVoltageEnabled = property.Value.ToObject<bool>();
                    }
                    else if (property.Name == "biasVoltage")
                    {
                        globalBiasVoltage = property.Value["value"]?.ToObject<int>() ?? 0;
                    }
                }

                armed = globalHighVoltageEnabled && globalBiasVoltage == 300;

            }
        }
        private void HandleMeasurementData(JObject json)
        {
            var values = json["values"];
            var charge = values?["charge"];
            var running = values?["measurementRunning"]?.ToObject<bool>() ?? false;

            // Update the measurement status
            UpdateStatusLabel(running);
            LogRunningState(running);



            if (charge != null)
            {
                lastKnownCharge = charge.ToObject<double>() / 1e-9;
                Measurement.Text = $"{lastKnownCharge:0.0000}";
            }

        }
        private void HandleMeasurementHistoryResponse(JObject json, string rawMessage)
        {
            var values = json["values"] as JObject;
            if (values != null && values["measurementHistory"] != null)
            {
                Console.WriteLine("Measurement History:");
                foreach (var entry in values["measurementHistory"])
                {
                    var measurement = entry as JObject;
                    if (measurement != null)
                    {
                        var time = measurement["time"]?.ToString();
                        var bias = measurement["biasVoltage"]?["value"]?.ToString();
                        var sensitivity = measurement["sensitivity"]?.ToString();
                        var mode = measurement["measurementMode"]?.ToString();
                        var highVoltageEnabled = measurement["highVoltageEnabled"]?.ToString();
                        var charge = measurement["charge"]?.ToString();

                        string essentialData = $"Time: {time}, Bias: {bias}, Sensitivity: {sensitivity}, " +
                                               $"Mode: {mode}, High Voltage Enabled: {highVoltageEnabled}, Charge: {charge}";

                        // Write the essential data to a text file
                        WriteMessageToFile(essentialData, "MeasurementHistory");
                        Console.WriteLine(essentialData);
                    }
                }
            }
            else
            {
                Console.WriteLine("No measurement history found in the response.");
            }
        }

        private void WriteMessageToFile(string message, string fileType)
        {
            try
            {
                string directoryPath = Path.GetDirectoryName(machine.Path);
                if (!string.IsNullOrEmpty(directoryPath))
                {
                    string fileName = $"PDM_history_{fileType}.txt";
                    string filePath = Path.Combine(directoryPath, fileName);
                    File.AppendAllText(filePath, message + Environment.NewLine);
                    Console.WriteLine($"Message written to {filePath}");
                }
                else
                {
                    Console.WriteLine("Directory path is invalid.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error writing message to file: " + ex.Message);
                Console.WriteLine("Error writing message to file: " + ex.Message);
            }
        }
        private async Task ReadHighVoltageAndBiasVoltageAsync()
        {
            try
            {
                var request = new
                {
                    cmd = "get_values",
                    values = new[] { "highVoltageEnabled", "biasVoltage" }
                };
                string jsonRequest = JsonConvert.SerializeObject(request);
                var buffer = Encoding.UTF8.GetBytes(jsonRequest);
                var segment = new ArraySegment<byte>(buffer);

                await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
                Console.WriteLine("High voltage and bias voltage read requested.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error requesting high voltage and bias voltage: " + ex.Message);
                Console.WriteLine("Error requesting high voltage and bias voltage: " + ex.Message);
            }
        }

        private async void UpdateStatusLabel(bool isCollecting)
        {
            // Ensure this method runs on the UI thread
            if (InvokeRequired)
            {
                Invoke(new Action<bool>(UpdateStatusLabel), isCollecting);
                return;
            }

            if (isCollecting)
            {
                hasSeenTrue = true;  // Set flag when we see 'true'
                falseCountAfterTrue = 0;  // Reset count to zero whenever 'true' is received
                statusLabel.Text = "Collecting Data...";
                statusLabel.BackColor = Color.LightGreen;
                statusLabel.ForeColor = Color.Black;
            }
            else
            {
                if (hasSeenTrue)
                {
                    falseCountAfterTrue++;  // Increment count for each 'false' after a 'true'

                    if (falseCountAfterTrue >= 5)  // Check if five 'false' values have been seen
                    {
                        int rangeSize = GetNamedRangeSize(currentNamedRange); // Get the current named range size
                        int indexToUpdate = Results.SelectedIndex;

                        if (indexToUpdate >= 0)
                        {
                            // Update existing item
                            Results.Items[indexToUpdate] = Measurement.Text;
                            // UpdateExcelCell(indexToUpdate + 1, Measurement.Text);
                            statusLabel.Text = "";
                            statusLabel.BackColor = this.BackColor;
                            statusLabel.ForeColor = Color.Black;
                            if (AutoAccept.Checked) { UpdateExcelRange(); }
                            Results.SelectedIndex = -1;
                        }
                        else if (Results.Items.Count < rangeSize)
                        {
                            // Add new measurement if there is space in the named range
                            Results.Items.Add(Measurement.Text);
                            statusLabel.Text = "Data collection complete.";
                            statusLabel.BackColor = this.BackColor;
                            statusLabel.ForeColor = Color.Black;
                            if (AutoAccept.Checked) { UpdateExcelRange(); }
                            // UpdateExcelCell(Results.Items.Count, Measurement.Text);
                        }
                        else
                        {
                            MessageBox.Show("Cannot add more items than the available cells in the named range.");
                        }

                        // Reset counters and flags after adding to results or reaching limit
                        falseCountAfterTrue = 0;
                        hasSeenTrue = false;
                    }
                }
            }

            previousIsCollecting = isCollecting;  // Maintain the last collecting state

            // Update the armed status and UI
            if (armed != previousArmedState)
            {
                if (!armed)
                {
                    Measurement.BackColor = Color.Yellow;
                    MessageBox.Show("Voltage is not enabled or the bias voltage is not set to 300V.");
                }
                else
                {
                    Measurement.BackColor = SystemColors.Window;
                }
                previousArmedState = armed;
            }
            else
            {
                Measurement.BackColor = armed ? SystemColors.Window : Color.Yellow;
            }
        }




        private void LogRunningState(bool running)
        {
            // Debug.WriteLine($"Running state updated to: {running}");
        }

        private async void DiscoverServices()
        {
            try
            {
                // Replace "_http._tcp.local." with your specific service type identifier
                var results = await ZeroconfResolver.ResolveAsync("_http._tcp.local.");

                // Find the first host with "DoseX" in the display name, case insensitive
                var targetHost = results.FirstOrDefault(host => host.DisplayName.IndexOf("DoseX", StringComparison.OrdinalIgnoreCase) >= 0);

                if (targetHost != null)
                {
                    Device.Text = targetHost.DisplayName;
                    Discover.Text = "Connect";
                    IPAddress = targetHost.IPAddress;  // Assuming IPAddress is a field or a property in this context

                    foreach (var service in targetHost.Services)
                    {
                        string serviceKey = service.Key;
                        var serviceProperties = service.Value;

                        // Example of accessing properties, commented out for now
                        // MessageBox.Show($"Service on Port: {serviceProperties.Port}");
                        foreach (var property in serviceProperties.Properties)
                        {
                            foreach (var record in property)
                            {
                                // MessageBox.Show($"{record.Key}: {record.Value}");
                            }
                        }
                    }
                }
                else
                {
                    Discover.Text = "Discover";
                    MessageBox.Show("No DoseX services found.  Please ensure the electrometer is on and check the ethernet conections at both ends.  Then wait 30 seconds before attempting to discover again.");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error discovering services: {ex.Message}");
                Discover.Text = "Discover";
            }
        }


        private void InitializeExcel()
        {
            if (excelApp == null)
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;  // Only make visible for debugging, can be set to false otherwise
            }

            if (string.IsNullOrEmpty(machine.Path))
            {
                MessageBox.Show("The file path is not specified.");
                return;
            }

            if (workbook == null)
            {
                workbook = excelApp.Workbooks.Open(machine.Path);
            }

            // Here, specify the sheet you are primarily working with by name or index
            if (worksheet == null)
            {
                worksheet = (Excel.Worksheet)workbook.Sheets[1]; worksheet.Activate();
                excelConnected = true;
            }
        }

        public class Machine
        {
            public string Path { get; set; }
            public string Name { get; set; }
            public string Type { get; set; }
            public List<string> Energies { get; set; }

            public Machine()
            {
                Energies = new List<string>();
            }
        }

        public class LogEntry
        {
            public string Path { get; set; }
            public string Test { get; set; }
            public int RangeNumber { get; set; }
            public DateTime Date { get; set; }

        }


        public Machine ReadMachineConfig(string filePath)
        {
            var machine = new Machine();
            try
            {
                var lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    if (line.StartsWith("Annual QA Path:"))
                    {
                        machine.Path = line.Substring("Annual QA Path:".Length).Trim();
                    }
                    else if (line.StartsWith("Machine Name:"))
                    {
                        machine.Name = line.Substring("Machine Name:".Length).Trim();
                    }
                    else if (line.StartsWith("Machine Type:"))
                    {
                        machine.Type = line.Substring("Machine Type:".Length).Trim();
                    }
                    else if (line.StartsWith("- "))
                    {
                        machine.Energies.Add(line.Substring("- ".Length).Trim());
                    }
                }

                return machine;
            }
            catch
            {

                return machine;
            }
        }
        private void PopulateComboBox()
        {
            // Define the list of names
            List<string> testTypes = new List<string>
            {
                "Field Size Factor",
                "Photon Linearity",
                "Electron Linearity",
                "Dose Rate Constancy",
                "Dosimetric Leaf Gap"
            };

            // Add items to the ComboBox
            TestType.Items.Clear(); // Clear existing items if necessary
            foreach (string testType in testTypes)
            {
                TestType.Items.Add(testType);
            }

            // Optionally set the default selected item
            TestType.SelectedIndex = -1;
        }

        private void Accept_Click(object sender, EventArgs e)
        {
            if (excelConnected) { UpdateExcelRange(); }

        }
        private void UpdateExcelRange()
        {
            if (worksheet == null)
            {
                MessageBox.Show("Worksheet is not initialized.");
                return;
            }

            if (string.IsNullOrEmpty(currentNamedRange))
            {
                MessageBox.Show("Named range is not specified.");
                return;
            }

            if (Results == null || Results.Items == null)
            {
                MessageBox.Show("Results list is not initialized.");
                return;
            }

            if (machine == null || string.IsNullOrEmpty(machine.Path))
            {
                MessageBox.Show("Machine configuration is not initialized.");
                return;
            }

            if (logger == null)
            {
                MessageBox.Show("Logger is not initialized.");
                return;
            }

            try
            {
                Excel.Range namedRange = worksheet.Range[currentNamedRange];
                if (namedRange == null)
                {
                    MessageBox.Show("Named range not found in the worksheet.");
                    return;
                }

                Excel.Range cells = namedRange.Cells;
                int itemCount = Results.Items.Count;

                // Use a single loop to iterate through the cells in the named range
                int itemIndex = 0;
                foreach (Excel.Range cell in cells)
                {
                    if (itemIndex >= itemCount)
                    {
                        break; // Stop if there are no more items to place
                    }

                    // Update each cell with corresponding item from the list
                    string itemValue = Results.Items[itemIndex].ToString();
                    if (double.TryParse(itemValue, out double chargeValue))
                    {
                        cell.Value2 = itemValue;

                        // Log the update
                        logger.AddOrUpdateLog(currentNamedRange, itemIndex + 1, machine.Path, chargeValue);
                    }
                    else
                    {
                        MessageBox.Show($"Invalid item value: {itemValue}");
                    }

                    itemIndex++;
                }

                // Optionally save the workbook after updating all cells
                // workbook.Save();

                // Save the log entries to the XML file after updating all cells
                logger.SaveLogEntries();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating Excel range: " + ex.Message);
            }
        }




        private void Measurement_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter && TestType.SelectedIndex != -1 && Energy.SelectedIndex != -1)
            {
                e.SuppressKeyPress = true;  // Prevent the ding sound

                int rangeSize = GetNamedRangeSize(currentNamedRange);
                int indexToUpdate = Results.SelectedIndex;

                if (indexToUpdate >= 0)
                {
                    // Update existing item
                    Results.Items[indexToUpdate] = Measurement.Text;
                    if (AutoAccept.Checked) { UpdateExcelRange(); }

                }
                else if (Results.Items.Count < rangeSize)
                {
                    // Add new measurement if there is space in the named range
                    Results.Items.Add(Measurement.Text);
                    if (AutoAccept.Checked) { UpdateExcelRange(); }

                }
                else
                {
                    MessageBox.Show("Cannot add more items than the available cells in the named range.");
                }

                Results.ClearSelected();
                Measurement.Clear(); // Clear the textbox for new entry
            }
        }


        private void Results_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Toggle selection: If the user clicks the same item again, it will deselect it
            if (Results.SelectedIndex == Results.Tag as int?)
            {
                Results.SelectedIndex = -1; // Deselect item
            }
            Results.Tag = Results.SelectedIndex; // Store current index to check for toggle
        }





        private int GetNamedRangeSize(string namedRange)
        {
            if (workbook == null)
            {
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Open(machine.Path);
            }

            try
            {

                Excel.Name excelNamedRange = workbook.Names.Item(namedRange);

                if (excelNamedRange != null)
                {
                    Excel.Range range = excelNamedRange.RefersToRange;

                    return range.Cells.Count;
                }
                else
                {
                    MessageBox.Show("Excel sheet not connected or Named range not found.");
                    return 0;  // Named range not found
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to retrieve named range size: " + ex.Message);
                return 0;  // Handle exceptions by returning 0 or another appropriate value
            }
        }


        private void TestType_SelectedIndexChanged(object sender, EventArgs e)
        {
            Energy.SelectedIndex = -1;
            // Clear previous entries in Energy ComboBox to prepare for new items
            Energy.Items.Clear();

            if (TestType.SelectedIndex == -1)
            {
                Chamber_Depth.Text = "";
            }
            else if (TestType.SelectedItem.ToString() == "Field Size Factor")
            {
                Chamber_Depth.Text = "CC13 Chamber, 10cm depth, 100SSD";
                UpdateBeamConfiguration(0.2, 0.2, 0.2, 0.2);
                SetMeasurementConfigurationAsync();
            }
            else if (TestType.SelectedItem.ToString() == "Electron Linearity")
            {
                Chamber_Depth.Text = "Farmer Chamber, 2cm depth, 100SSD";
                UpdateBeamConfiguration(0.2, 0.2, 0.5, 0.5);
                SetMeasurementConfigurationAsync();
            }
            else if (TestType.SelectedItem.ToString() == "Dosimetric Leaf Gap")
            {
                Chamber_Depth.Text = "Farmer Chamber, 5cm depth, 95SSD";
                UpdateBeamConfiguration(0.5, 0.5, 0.5, 0.5);
                SetMeasurementConfigurationAsync();
            }
            else
            {
                Chamber_Depth.Text = "Farmer Chamber, 10cm depth, 100SSD";
                UpdateBeamConfiguration(0.1, 0.1, 0.1, 0.1);
                SetMeasurementConfigurationAsync();
            }

            if (TestType.SelectedItem != null) // Ensure there is a selected item to avoid null reference
            {
                string selectedTestType = TestType.SelectedItem.ToString();

                // Determine which type of energy to add based on the selected test type
                string energyType = selectedTestType == "Electron Linearity" ? "e" : "x";

                // List of all energies to be used when allEnergies is checked
                List<string> allEnergiesList = new List<string>
                {
                    "4x", "6x", "6xfff", "8x", "10x", "10xfff",
                    "15x", "16x", "18x", "20x", "23x", "6e", "9e",
                    "12e", "15e", "16e", "18e", "20e", "22e"
                };

                // Choose the correct energy list based on the allEnergies checkbox
                List<string> energiesToUse = allEnergies.Checked ? allEnergiesList : machine.Energies;

                // Add appropriate energy types to Energy ComboBox
                foreach (string energy in energiesToUse)
                {
                    if (energy.Contains(energyType))
                    {
                        Energy.Items.Add(energy);
                    }
                }

                Energy.SelectedIndex = -1;
            }
        }

        private Excel.Range OpenExcelAndSelectSheetByNamedRange(string filePath, string namedRange)
        {
            Excel.Workbook workbook = null;
            Excel.Range range = null;
            try
            {
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Open(filePath);

                Excel.Name excelNamedRange = workbook.Names.Item(namedRange);
                range = excelNamedRange.RefersToRange;
                range.Worksheet.Activate();
                range.Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error accessing named range/nPlease verify that a vaild energy and test is selected in the dropdowns. " + ex.Message);
            }
            return range;
        }

        private void ClearResults_Click(object sender, EventArgs e)
        {
            Results.Items.Clear();
        }
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e); // Ensure base class cleanup happens first

            // Prompt the user to save changes if the workbook is open and not null
            if (workbook != null)
            {
                var result = MessageBox.Show("Do you want to save changes to your workbook?", "Save Changes", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        workbook.Save();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to save the workbook: {ex.Message}");
                    }
                }

                // Close the workbook safely
                try
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to close the workbook: {ex.Message}");
                }
                finally
                {
                    workbook = null;
                }
            }

            // Safely release the worksheet
            if (worksheet != null)
            {
                Marshal.ReleaseComObject(worksheet);
                worksheet = null;
            }

            // Safely quit Excel application
            if (excelApp != null)
            {
                try
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to quit Excel application: {ex.Message}");
                }
                finally
                {
                    excelApp = null;
                }
            }

            // Handle WebSocket if open
            if (wsClient != null && wsClient.State == WebSocketState.Open)
            {
                try
                {
                    cancellationTokenSource.Cancel(); // Cancel any ongoing operations
                    wsClient.CloseAsync(WebSocketCloseStatus.NormalClosure, "Closing", CancellationToken.None).Wait();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to close WebSocket: {ex.Message}");
                }
            }
        }

        public int ExtractNumber(string input)
        {
            var numericText = Regex.Replace(input, "[^0-9]", "");
            return string.IsNullOrEmpty(numericText) ? 0 : int.Parse(numericText);
        }

        private void Energy_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TestType.SelectedItem != null && Energy.SelectedItem != null)
            {
                string selectedEnergy = Energy.SelectedItem.ToString();

                // Regular expression pattern to match numeric characters
                string numericPattern = @"\d+";
                int energyValue = ExtractNumber(Energy.SelectedText);
                               
                // Check if selectedEnergy contains a number
                if (Regex.IsMatch(selectedEnergy, numericPattern))
                {
                    string prefix = TestType.SelectedItem.ToString() switch
                    {
                        string t when t.Contains("Linearity") => "MUL_",
                        string t when t.Contains("Field Size Factor") => "FSF_",
                        string t when t.Contains("Dose Rate Constancy") => "DRC_",
                        string t when t.Contains("Dosimetric Leaf Gap") => "DLG_",
                        _ => string.Empty
                    };
                    if ( TestType.SelectedItem.ToString() == "Dosimetric Leaf Gap" && energyValue > 16)
                    { Chamber_Depth.Text = "Farmer Chamber, 10cm depth, 90SSD"; }
                    if (!string.IsNullOrEmpty(prefix))
                    {
                        currentNamedRange = prefix + selectedEnergy.ToUpper();
                        //MessageBox.Show(currentNamedRange);
                        UpdateWorksheet(currentNamedRange); // Update the worksheet based on the selected named range
                        LoadResultsFromExcel(machine.Path, currentNamedRange);
                    }
                }
                else
                {
                    // Handle case where selected energy does not contain a number
                    MessageBox.Show("Selected energy is invalid.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        private void UpdateWorksheet(string namedRange)
        {
            if (excelApp == null || workbook == null)
            {
                MessageBox.Show("Excel application or workbook is not initialized.");
                return;
            }

            try
            {
                Excel.Name excelNamedRange = workbook.Names.Item(namedRange);
                if (excelNamedRange != null)
                {
                    worksheet = excelNamedRange.RefersToRange.Worksheet;
                    worksheet.Activate();  // Optionally activate the worksheet
                }
                else
                {
                    MessageBox.Show("Named range does not exist: " + namedRange);
                    worksheet = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to initialize worksheet: " + ex.Message);
                worksheet = null;
            }
        }

        private void LoadResultsFromExcel(string filePath, string namedRange)
        {

            Excel.Range range = OpenExcelAndSelectSheetByNamedRange(filePath, namedRange);
            if (range != null)
            {
                Results.Items.Clear();  // Clear previous items
                foreach (Excel.Range cell in range.Cells)
                {
                    if (cell.Value2 != null)
                        Results.Items.Add(cell.Value2.ToString());
                }
            }

        }

        private void Measurement_TextChanged(object sender, EventArgs e)
        {

        }

        private void Discover_Click(object sender, EventArgs e)
        {
            if (Discover.Text == "Discover") { DiscoverServices(); }
            else { ConnectToWebSocketAsync(); }
        }

        private TaskCompletionSource<bool> controlTcs;

        private async Task<bool> RequestControlTokenAsync()
        {
            try
            {
                controlTcs = new TaskCompletionSource<bool>();

                var request = new { cmd = "control", value = "request" };
                string jsonRequest = JsonConvert.SerializeObject(request);
                var buffer = Encoding.UTF8.GetBytes(jsonRequest);
                var segment = new ArraySegment<byte>(buffer);

                await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
                Console.WriteLine("Control token requested.");

                // Wait for the response
                return await controlTcs.Task;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error requesting control token: " + ex.Message);
                Console.WriteLine("Error requesting control token: " + ex.Message);
                return false; // Return false if there is an error
            }
        }

        private BeamConfiguration beamConfig;
        public void UpdateBeamConfiguration(double beamOnTime, double beamOffTime, double preTriggerTime, double postTriggerTime)
        {
            beamConfig = new BeamConfiguration(beamOnTime, beamOffTime, preTriggerTime, postTriggerTime);
        }

        private async Task SetMeasurementConfigurationAsync()
        {
            try
            {
                var config = new
                {
                    cmd = "change_values",
                    values = new Dictionary<string, object>
            {
                { "biasVoltage", new { value = 300, unit = "V" } },  // Set bias voltage to 300V
                { "highVoltageEnabled", true },  // Enable high voltage
                { "measurementMode", "charge" },  // Set measurement mode to charge
                { "autoReset", true },
                { "sensitivity", "mid" },
                { "measurementTimerMode", "trigger" },  // Set start type to trigger
                { "beamOnTime", new { value = beamConfig.BeamOnTime, unit = "s" } },  // Use beam on time from config
                { "beamOffTime", new { value = beamConfig.BeamOffTime, unit = "s" } },  // Use beam off time from config
                { "preTriggerTime", new { value = beamConfig.PreTriggerTime, unit = "s" } },  // Use pre-trigger time from config
                { "postTriggerTime", new { value = beamConfig.PostTriggerTime, unit = "s" } }  // Use post-trigger time from config
            }
                };

                string jsonConfig = JsonConvert.SerializeObject(config);
                var buffer = Encoding.UTF8.GetBytes(jsonConfig);
                var segment = new ArraySegment<byte>(buffer);

                await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
                Console.WriteLine("Measurement configuration sent: " + jsonConfig);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error setting measurement configuration: " + ex.Message);
                Console.WriteLine("Error setting measurement configuration: " + ex.Message);
            }
        }


        private async Task<bool> AuthenticateAdminAsync(string passWord)
        {
            try
            {
                loginTcs = new TaskCompletionSource<bool>();

                var request = new { cmd = "login", pass = passWord };
                string jsonRequest = JsonConvert.SerializeObject(request);
                var buffer = Encoding.UTF8.GetBytes(jsonRequest);
                var segment = new ArraySegment<byte>(buffer);

                //MessageBox.Show("Sending admin login request..." + request);
                await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);

                // Wait for the response
                return await loginTcs.Task;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error requesting admin login: " + ex.Message);
                Console.WriteLine("Error requesting admin login: " + ex.Message);
                return false; // Return false if there is an error
            }
        }






        private async Task StartMeasurementAsync(bool autoReset = true)
        {
            try
            {
                var request = new
                {
                    cmd = "measurement",
                    value = "start",
                    autoReset = autoReset
                };
                string jsonRequest = JsonConvert.SerializeObject(request);
                var buffer = Encoding.UTF8.GetBytes(jsonRequest);
                var segment = new ArraySegment<byte>(buffer);

                await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
                Console.WriteLine("Measurement start requested.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error starting measurement: " + ex.Message);
                Console.WriteLine("Error starting measurement: " + ex.Message);
            }
        }

        private async Task ReleaseControlAsync()
        {
            try
            {
                var request = new { cmd = "control", value = "release" };
                string jsonRequest = JsonConvert.SerializeObject(request);
                var buffer = Encoding.UTF8.GetBytes(jsonRequest);
                var segment = new ArraySegment<byte>(buffer);

                await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
                Console.WriteLine("Control released.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error releasing control: " + ex.Message);
                Console.WriteLine("Error releasing control: " + ex.Message);
            }
        }
        private async Task RequestMeasurementHistoryAsync()
        {
            if (doseXConnected)
            {
                try
                {
                    var request = new
                    {
                        cmd = "get_values",
                        values = new[] { "measurementHistory" }
                    };
                    string jsonRequest = JsonConvert.SerializeObject(request);
                    var buffer = Encoding.UTF8.GetBytes(jsonRequest);
                    var segment = new ArraySegment<byte>(buffer);

                    await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
                    Console.WriteLine("Measurement history requested.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error requesting measurement history: " + ex.Message);
                    Console.WriteLine("Error requesting measurement history: " + ex.Message);
                }


            }
            else { MessageBox.Show("Connect Electrometer"); }

        }


        private async void Arm_Click(object sender, EventArgs e)
        {
            if (doseXConnected)
            {
                requested = true;


                bool isControlAvailable = await RequestControlTokenAsync();  // Check control status before proceeding

                if (!isControlAvailable)
                {
                    MessageBox.Show("Please free control from the electrometer by pressing the 'cloud' at the top.");
                }
                else
                {
                    await AuthenticateAdminAsync(PW);
                    await SetMeasurementConfigurationAsync();
                    await StartMeasurementAsync();
                    await ReleaseControlAsync();
                }
            }
            else
            {
                MessageBox.Show("Connect Electrometer");
            }
        }





        private void SaveHistory_Click(object sender, EventArgs e)
        {
            logger.SaveLogEntries();
            string message =
                            "It's my duty to inform you that the 'Save History' button you just pressed is vestigial. \nIt will be removed or repurposed soon. I hope you haven't grown terribly attached to this button. \n" +
                            "You see, the application has been automatically saving every time new data is written to Excel. \nThis includes saving the charge, timestamp, test type, range number, and other details to a log file. \n" +
                            "Therefore, there should be no need for you to press this button anymore. \nIt's just one more thing to forget, and we already have enough of those, don't we? \n" +
                            "I'll save it out again right now but, really, nothing has changed...                                                                       \n\n" +
                            "+=+-+=+-+-=--=+-==+--+--=--:::::+::::::::::::::::::::::::::::::::::==:-=:----==-===-===--+-=====-+=+\n" +
                            "---:--=:-:=--:----:::::-::::::::::::::::::::::::::::::::::::::::::::::::::-::::::::-----:-:-----:=-=\n" +
                            "=-=:=-=:--=--==-=-:--:-:::-:::::::::::::::::::::::::::::::::::::::::::::::--:-----=---=--=---:-=:=--\n" +
                            "+=+:*=*-+=*-+=*-==+--::::--::::::::::::::::::::::::::::::::::::::::::::::::--::--::=::+=-+--+-=+:*=+\n" +
                            "+=*-+=*-+=*-+=::=-*::::::::::::::::::::::::-*@@@@@@@@@@@@@#:::::::::::::::::::::::::=-+-::-=+-=+:*=+\n" +
                            "--=:=-=:::=:::::=-::::::::::::::::::::--#%@%*:::::::=::==##%@@+-::::::::::::::::::::-=::::::::--::::\n" +
                            "=-=----:=:+:::=:::::::::::::::::::::-=#%#=::::::::::::---==+=**%==::::::::::::::::::::--::--::---===\n" +
                            "-----:-::::::::::::::::::::::::::::=*#-.........:::::::::--=-+=+%#+-::::::::::::::::::::=-::+=++:-==\n" +
                            "---:-:----+::::::::::::::::::::::=**+.............:::::::::--==-#+%++-::::::::::::::::::::=-::-=:=--\n" +
                            "=:--=--:--:::::::::::::::::::::=#*=:..............:::::::::-----#=+###+:::::::::::::::::::::-----=--\n" +
                            "*===::=-=-+::::::::::::::::::-##-:................::::::::::::::=-=*+##@=:::::::::::::::::::----:==+\n" +
                            "+=-:==:::::::::::::::::::::::-##::::...............:::::::::::-:+=+*+**@+:::::::::::::::::::-:-+:-==\n" +
                            "=--:--::-:::::::::::::::::::*%=:...................:::::::::::=-@=+*=****#-:::::::::::::::::=---:=--\n" +
                            "+-=-:-+---::::::::::::::::::*%-.::.......:-=++*#%%%%%%##*+=:::--@=+*=+**#%-::::::::::::::::::::::+==\n" +
                            "+--===*::::::::::::::::::+**=-:...::=#%@##*+-::::::::::::-+*#%@%#==*=+*+##**=-::::::::::::::*::=:+=-\n" +
                            "+-----::--:::::::::::::-*+*@+..-*%%+=:::++***#############****=-=+#@%**=**@#+=::::::::::::::---=:=-=\n" +
                            "=-=:=--::::::::::::::::=@=-%++@*:::=+%%%@@@@@@@@@@@@@@@@@@@@@@@@@#-:-%@#**@*-+::::::::::::::::--:=-:\n" +
                            "=---=-=::::::::::::::::=@==@@*:-=#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@--+%@@#++:::::::::::::::--::-==\n" +
                            "=---=-:=--:::::::::::::=@**@::=#@@@@#--=@@@+==+@@@@@@@@*::-*@@@@+*@@@@#::+@#*+:::::::::::::::--::-==\n" +
                            "*=*-+=*::::::::::::::::-@@@#=+%@@@@@@@-.....-%@@@@@@@@@@@......:*%@@@@@@+=@@@+:::::::::::::::---=:=-\n" +
                            "*-====--:-::::::::::::::::-+%@+*@@@@@@@@:::#@@@@@@@@@@@@@@#-::*@@@@@@@#=%@%+::::::::::::::::*:==:+=-\n" +
                            "=----=-:--::::::::::::::::-=*#*+=*%@@@@@@@@@@@@%%%%%@@%@@@@@@@@@@@@%#=+*#*-:::::::::::::::::+::=:=-:\n" +
                            "=--:----:::::::::::::::::::::-**+==%@@@@@@@%*::--------::=#@@@@@@@%+=**#=:::::::::::::::::::--=-:=-:\n" +
                            "+--=::-::::::::::::::::::::::::=**+=+#@@@@@#**@@@@@@@@@@@=*@@@@@#*+###=:::::::::::::::::::::-:-=:=--\n" +
                            "+=-=-:-=:::::::::::::::::::::::::=+#*****#@@@@@@@@@@@@@@@@@@%#***#@*+:::::::::::::::::::::::+::=:==+\n" +
                            "+=-=---::::::::::::::::::::::::-+++**+*##*+++++*%@@@@@@#*+++##%#**#++-::::::::::::::::::::::*-=+:+=+\n" +
                            "=--:::::::::::::::::::::::::==+++++=--%@=+%###%%*=====+#%%%@+%%@==+**==++=:::::::::::::::::::-=::--:\n" +
                            "=---::=:::::::::::::::::::::*@*=::::::%@:-=**#*##@@@@#%###**=##@======+#%%-:::::::::::::::::-:-::---\n" +
                            "=--=::::::::::::::::::::::::=+*=.::::-%@:::::-:==+++*=*+=+===##@-::::-=#*+::::::::::::::::::--=:::--\n" +
                            "+==-:-+::::::::::::::::::::::-%%+-.::-%@:::::::::::::.::::::-*#@:::::+#@+::::::::::::::::::::-++:+=+\n" +
                            "+==-=-:-::::::::::::::::::::=**#@*=.:=%@:::::::::::::::::::--*#%:::=+@%#*+::::::::::::::::::=::=:---\n" +
                            "=--:---:::::::::::::::::::-+*+:=%@%+=-%@:::::::::::::::::::--*#%:-+%@#+=#%-:::::::::::::::::=---:=--\n" +
                            "=-=-=:=-::::::::::::::::::=@+::-+#%@%+@@:::::::::::::::::::--+#@+*@#*---***=::::::::::::::::+-==:=-:\n" +
                            "+=+-+===-:::::::::::::::::=%#++*#*****@@:.::::::::::::::::::=#*%@%+##**##%@*::::::::::::::::*::=:===\n" +
                            "+=-=-=+::::::::::::::::=#%@@@@@@@@@%*-%@..:+*%@@@@@@@@@@@@*:=**#@@%@@@@@@@@@%=:::::::::::::::-=+::--\n" +
                            "::::::-::::::::::::::::+@=....:-=+*@*@@@--:%@@*-=+++===*@@#---+#@#::::::-=#+@=:::::::::::::::::=:--:\n" +
                            "=--=::=::::::::::::::::+@=....::-=+@*@@@@@@@#:::::::-:=+=%@@@@@@@#::::::--++@=::::::::::::::-::=:=-=\n" +
                            "+--*==:::::::::::::::::+@=::.:::==+@*@-:#%@-..::::::::-==**@@%**@#..::::-=*+@+-::::::::::::::-+::+:-\n" +
                            "*=-:::+::-=::::::::::::+@=::::::=++@*@=:..:#@@@@@@@@@@@@@@%*=#**@#....::==*+@*+::::::::::::::::+::=+\n" +
                            "----::=:--:::::::::::::+@=::::::===@*@@@-:..:::::-===:=--+=*=###@#....::--*+@%%:::::::::::::::-=:=::\n" +
                            "-----:-:-::::::::::::::+@=::::::===@*@%#@#=:.:::::::::---*+**@%%@#...:::--*+%%@::::::::::::::-:-:--:\n" +
                            "+---=:::-::::::::::::::+@+=====+***%#@*=#%@+=::::::::::=-*+**#**@%======*+#*%%@:::::::::::::-:=:::=+\n" +
                            "+--=-::::::::::::::::::-+%@*+**#@@@##+%@***%#::::::::::=-%#***%@+*#%%@#*##%%@++:::::::::::::::::-=-:\n" +
                            "=--=--::=::::::::::::::::%@%=::+@@@@#:%@=*%%%%+:::::-:*++@@%*+%%:=%@@@+.++*#@=::::::::::::::::-+:+==\n" +
                            "=---::::::::::::::::::::::=@@@*+=--:::%@@@@@@@@@@@*+%+%@@@@@@@@%:::==-+*@@@*::::::::::::::::-::::=--\n" +
                            "=--:-:-:::-::::::::::::::::::-%@@*::::%@:::=+#*@*+@@@@@@@:::-=##@*::-%@@+::::::::::::::::::::---:---\n" +
                            "=--:=---:-+::::::::::::::::::::::::::-%@::::-*+@+::::=%#+::::-#*@#::::::::::::::::::::::::::--=+:+==\n" +
                            "*=+-:-:--::=::::::::::::::::::::::::--++:::::=+@+::::+@+::::-=**@#::::::::::::::::::::::::::--=+-+=+\n" +
                            "*===-:=:-::--::::::::::::::::::::::--=%@@@@@@@@@+::::-=%@@@@@@@@@#-::::::::::::::::::::::-::-:--:=--\n" +
                            "-:::--=----:-=::::::::::::::::::::-@@@-:::::-++@@@:::*@+::::::::%+@=:::::::::::::::::::::--:=---:=--\n" +
                            "*=+-+=-=::--+::+:::::::::::::::::#%::::::::::*-%#@:=@#::::::::::#=*%@-::::::::::::::::::::-=:-=::+==\n" +
                            "+=+-+-+:--+:-=:::::::::::::-+%=+@*--::---+++*#=%#%@@@#::------:-*+*%*@@@=+@#*%+::::::::::*::+-++:+=+\n" +
                            "+==:+-:-=-+--:::::::+@:=%:-%+%#*@@@@@@@@@@@@@@@@@@@@%@@@@@@@@@@@@@@@@@%%%#%*#%*%##*+::=-::--:::=:+=+\n" +
                            "=--:-:=:-:=-===*:::::::::::::::::::*++=====-=+-===-======+==-===+====-----:::::::::::::::*==+-==:--:\n" +
                            "=-=---=--:=-=-=-=-=--::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::==-:--:::-:--:=--";
            CustomMessageBox.Show(message);
        }

        private void allEnergies_CheckedChanged(object sender, EventArgs e)
        {
            Energy.SelectedIndex = -1;
            // Clear previous entries in Energy ComboBox to prepare for new items
            Energy.Items.Clear();
            if (allEnergies.Checked)
            {
                allowAllEnergies = true;
            }
            else
            {
                allowAllEnergies = false;
            }
            if (TestType.SelectedItem != null) // Ensure there is a selected item to avoid null reference
            {
                string selectedTestType = TestType.SelectedItem.ToString();

                // Determine which type of energy to add based on the selected test type
                string energyType = selectedTestType == "Electron Linearity" ? "e" : "x";

                // List of all energies to be used when allEnergies is checked
                List<string> allEnergiesList = new List<string>
                {
                    "4x", "6x", "6xfff", "8x", "10x", "10xfff",
                    "15x", "16x", "18x", "20x", "23x", "6e", "9e",
                    "12e", "15e", "16e", "18e", "20e", "22e"
                };

                // Choose the correct energy list based on the allEnergies checkbox
                List<string> energiesToUse = allEnergies.Checked ? allEnergiesList : machine.Energies;

                // Add appropriate energy types to Energy ComboBox
                foreach (string energy in energiesToUse)
                {
                    if (energy.Contains(energyType))
                    {
                        Energy.Items.Add(energy);
                    }
                }

                Energy.SelectedIndex = -1;
            }

        }

        private async void Start_Click(object sender, EventArgs e)
        {
            if (doseXConnected)
            {
                requested = true;


                bool isControlAvailable = await RequestControlTokenAsync();  // Check control status before proceeding

                if (!isControlAvailable)
                {
                    MessageBox.Show("Please free control from the electrometer by pressing the 'cloud' at the top.");
                }
                else
                {
                    await StartMeasurementAsync();
                    await ReleaseControlAsync();
                }
            }
            else
            {
                MessageBox.Show("Connect Electrometer");
            }


        }

        private void Delete_Click(object sender, EventArgs e)
        {
            if (Results.SelectedIndex != -1)
            {

                Results.Items.RemoveAt(Results.SelectedIndex);

            }
        }

        private void Insert_Click(object sender, EventArgs e)
        {
            if (Results.SelectedIndex != -1)
            {

                Results.Items.Insert(Results.SelectedIndex, 0.00);

            }
        }
        private async Task startBackgroundAsync()
        {

            try
            {
                var request = new
                {
                    cmd = "measurement",
                    value = "start_background"
                };
                string jsonRequest = JsonConvert.SerializeObject(request);
                var buffer = Encoding.UTF8.GetBytes(jsonRequest);
                var segment = new ArraySegment<byte>(buffer);

                await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
                Console.WriteLine("Measurement start requested.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error starting measurement: " + ex.Message);
                Console.WriteLine("Error starting measurement: " + ex.Message);
            }
        }

        async private void Background_Click(object sender, EventArgs e)
        {
            bool isControlAvailable = await RequestControlTokenAsync();  // Check control status before proceeding

            if (!isControlAvailable)
            {
                MessageBox.Show("Please free control from the electrometer by pressing the 'cloud' at the top.");
            }
            else
            {
                await startBackgroundAsync();
                await ReleaseControlAsync();
            }

        }

        async private void Admin_Click(object sender, EventArgs e)
        {
            await AuthenticateAdminAsync(PW);
        }
    }
    public class LogEntry
    {
        public string Path { get; set; }
        public string Test { get; set; }
        public int RangeNumber { get; set; }
        public DateTime Date { get; set; }
        public double Charge { get; set; } // Updated to include charge
    }

    public class Logger
    {
        private Dictionary<(string Path, string Test, int RangeNumber), LogEntry> logEntries;
        private string logFilePath;
        private List<LogEntry> logHistory; // To keep history of all entries

        public Logger(string machinePath)
        {
            logEntries = new Dictionary<(string Path, string Test, int RangeNumber), LogEntry>();
            logHistory = new List<LogEntry>();
            logFilePath = Path.Combine(Path.GetDirectoryName(machinePath), "DoseX_logfile.xml");

            // Load existing log entries from file if it exists
            LoadLogEntries();
        }

        public void AddOrUpdateLog(string test, int rangeNumber, string machinePath, double charge)
        {
            var key = (Path: machinePath, Test: test, RangeNumber: rangeNumber);

            LogEntry newLogEntry = new LogEntry
            {
                Path = machinePath,
                Test = test,
                RangeNumber = rangeNumber,
                Date = DateTime.Now,
                Charge = charge
            };

            logHistory.Add(newLogEntry); // Add to history

            if (logEntries.TryGetValue(key, out LogEntry existingLogEntry))
            {
                // Overwrite the existing entry only if both the charge value and timestamp are different
                if (existingLogEntry.Charge != newLogEntry.Charge && existingLogEntry.Date != newLogEntry.Date)
                {
                    logEntries[key] = newLogEntry;
                }
            }
            else
            {
                logEntries[key] = newLogEntry; // Add the new log entry
            }
        }

        public void SaveLogEntries()
        {
            List<LogEntry> entriesToSave = new List<LogEntry>(logEntries.Values);

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<LogEntry>));
            using (FileStream fs = new FileStream(logFilePath, FileMode.Create))
            {
                xmlSerializer.Serialize(fs, entriesToSave);
            }

            // Optionally save the log history
            /*string historyFilePath = Path.Combine(Path.GetDirectoryName(logFilePath), "DoseX_logfile_history.xml");
            using (FileStream fs = new FileStream(historyFilePath, FileMode.Create))
            {
                xmlSerializer.Serialize(fs, logHistory);
            }*/
        }

        private void LoadLogEntries()
        {
            if (File.Exists(logFilePath))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(List<LogEntry>));
                using (FileStream fs = new FileStream(logFilePath, FileMode.Open))
                {
                    var entries = (List<LogEntry>)serializer.Deserialize(fs);
                    foreach (var entry in entries)
                    {
                        var key = (entry.Path, entry.Test, entry.RangeNumber);
                        logEntries[key] = entry;
                    }
                }
            }
        }
    }

    public class CustomMessageBox : Form
    {
        private RichTextBox richTextBox;

        public CustomMessageBox(string message)
        {
            this.richTextBox = new RichTextBox();
            this.richTextBox.ReadOnly = true;
            this.richTextBox.Dock = DockStyle.Fill;
            this.richTextBox.Font = new System.Drawing.Font("Consolas", 10); // Monospaced font
            this.richTextBox.Text = message;
            this.richTextBox.WordWrap = false; // Prevents wrapping
            this.richTextBox.ScrollBars = RichTextBoxScrollBars.Both; // Enable both scrollbars

            this.Controls.Add(this.richTextBox);

            this.Text = "A Morose Message from Melvin";
            this.Size = new Size(765, 1000); // Set an appropriate window size
            this.FormBorderStyle = FormBorderStyle.FixedDialog; // Disable resizing
            this.MaximizeBox = false; // Disable maximize button
            this.StartPosition = FormStartPosition.CenterScreen; // Center the form on the screen
        }

        public static void Show(string message)
        {
            CustomMessageBox form = new CustomMessageBox(message);
            form.ShowDialog();
        }
    }

}