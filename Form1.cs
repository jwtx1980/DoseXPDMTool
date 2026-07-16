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
    using System.Xml.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
    using System.Globalization;


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
        private const int SecureDoseXWebSocketPort = 8083;
        private const int PlainDoseXWebSocketPort = 8081;
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
        private bool suppressNextMeasurementCommit = false;
        private DateTime lastMeasurementCommitAt = DateTime.MinValue;
        private static readonly TimeSpan MeasurementCommitDeadTime = TimeSpan.FromMilliseconds(100);
        private bool globalHighVoltageEnabled;
        private int globalBiasVoltage;
        private bool doseXConnected = false;
        private bool excelConnected = false;
        private bool allowAllEnergies = false;
        private Logger logger;
        private TaskCompletionSource<bool> loginTcs;
        private string PW = "VarianAOS1!";
        private const int Tg51ReadingsPerBias = 3;
        private const int Tg51BiasTransitionDelayMs = 5000;
        private const int Tg51ReadingsPerPoint = 9;
        private readonly int[] tg51BiasSequence = { 300, -300, 150 };
        private bool tg51TransitionInProgress = false;
        private int expectedBiasVoltage = 300;
        private Tg51Run tg51Run;
        private Tg51RunLogger tg51Logger;
        private int tg51CurrentPointIndex = -1;
        private bool tg51RunActive = false;
        private bool tg51ChangingSelection = false;
        private bool tg51EnergyChangeInProgress = false;
        private bool tg51BiasSelectionInProgress = false;
        private bool tg51MoveInProgress = false;
        private string tg51LastHandledEnergy = string.Empty;
        private string tg51DepthEnergy = string.Empty;
        private Tg51Point pendingTg51OverwritePoint = null;
        private Tg51Reading pendingTg51OverwriteReading = null;
        private int pendingTg51OverwriteBias = 0;
        private int pendingTg51OverwriteIndex = -1;
        private const int Tg51ManualBiasSettleDelayMs = 3000;
        private const string Tg51BridgeDefaultUrl = "http://127.0.0.1:8770";
        private const string Tg51BridgeCapabilitiesSchema = "collector-capabilities-v1";
        private const string Tg51BridgeFirewallRuleName = "TG-51 Tank Bridge Inbound";
        private const string Tg51BridgeInstallFolderName = "TG51TankBridge";
        private const string Tg51BridgeRuntimeFolderName = "TankBridgeRuntime";
        private static readonly TimeSpan Tg51LiveCallbackFreshness = TimeSpan.FromSeconds(5);
        private readonly Dictionary<Tg51Point, int> tg51ActiveBiasByPoint = new Dictionary<Tg51Point, int>();
        private readonly CheckedListBox Tg51EnergyList = new CheckedListBox();
        private readonly DataGridView Tg51PointGrid = new DataGridView();
        private readonly DataGridView Tg51ReadingGrid = new DataGridView();
        private readonly System.Windows.Forms.CheckBox Tg51UseBridge = new System.Windows.Forms.CheckBox();
        private readonly HttpClient tg51TankClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        private readonly System.Windows.Forms.Timer tg51TankStatusTimer = new System.Windows.Forms.Timer();
        private bool tg51TankStatusRefreshInProgress = false;

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
            DiscoverServices(false);
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
            InitializeTg51Workspace();
            ExecuteWithDelay();
        }
        public async Task ExecuteWithDelay()
        {
            for (int attempt = 0; attempt < 6; attempt++)
            {
                await Task.Delay(1000);

                if (doseXConnected)
                {
                    return;
                }

                if (!string.IsNullOrWhiteSpace(IPAddress) && Discover.Text == "Connect")
                {
                    ConnectToWebSocketAsync(false);
                    return;
                }

                if (Discover.Text == "Discover")
                {
                    DiscoverServices(false);
                }
            }
        }

        private async void ConnectToWebSocketAsync(bool showErrors = true)
        {
            if (string.IsNullOrWhiteSpace(IPAddress))
            {
                Discover.Text = "Discover";
                if (showErrors)
                {
                    MessageBox.Show("No DoseX IP address is available. Discover the electrometer first.");
                }

                return;
            }

            Discover.Text = "Connecting";
            Exception lastError = null;
            var endpoints = new[]
            {
                new Uri($"wss://{IPAddress}:{SecureDoseXWebSocketPort}"),
                new Uri($"ws://{IPAddress}:{PlainDoseXWebSocketPort}")
            };

            foreach (Uri serverUri in endpoints)
            {
                ClientWebSocket candidate = new ClientWebSocket();
                CancellationTokenSource connectTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(8));

                try
                {
                    candidate.Options.RemoteCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
                    await candidate.ConnectAsync(serverUri, connectTimeout.Token);

                    wsClient = candidate;
                    cancellationTokenSource = new CancellationTokenSource();
                    StartReceiving();
                    await ReadHighVoltageAndBiasVoltageAsync();
                    Discover.Text = "Connected";
                    doseXConnected = true;
                    return;
                }
                catch (Exception ex)
                {
                    lastError = ex;
                    Console.WriteLine($"DoseX websocket connection failed for {serverUri}: {ex.Message}");
                    candidate.Dispose();
                }
                finally
                {
                    connectTimeout.Dispose();
                }
            }

            Discover.Text = "Discover";
            doseXConnected = false;
            if (showErrors)
            {
                MessageBox.Show($"WebSocket Error: {FormatDoseXConnectionError(lastError)}");
            }
        }

        private string FormatDoseXConnectionError(Exception error)
        {
            if (error is OperationCanceledException || error is TimeoutException)
            {
                return "DoseX connection timed out. Confirm the electrometer is powered on, the USB/Ethernet adapter is recognized, and the DoseX IP address is reachable, then try Discover again.";
            }

            if (error is WebSocketException)
            {
                return error.Message;
            }

            return error?.Message ?? "Unable to connect to the remote server";
        }






        private async void StartReceiving()
        {
            ClientWebSocket receivingClient = wsClient;
            CancellationTokenSource receivingTokenSource = cancellationTokenSource;
            if (receivingClient == null || receivingTokenSource == null)
            {
                return;
            }

            var buffer = new byte[1024];
            while (receivingClient.State == WebSocketState.Open)
            {
                try
                {
                    var result = await receivingClient.ReceiveAsync(new ArraySegment<byte>(buffer), receivingTokenSource.Token);
                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        await receivingClient.CloseAsync(WebSocketCloseStatus.NormalClosure, string.Empty, CancellationToken.None);
                    }
                    else
                    {
                        var message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                        Console.WriteLine($"Received message: {message}");
                        HandleWebSocketMessage(message);
                    }
                }
                catch (OperationCanceledException ex)
                {
                    Console.WriteLine("WebSocket receive canceled: " + ex.Message);
                    break;
                }
                catch (ObjectDisposedException ex)
                {
                    Console.WriteLine("WebSocket receive stopped after dispose: " + ex.Message);
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error receiving message: " + ex.Message);
                    break;
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
                MessageBox.Show($"Bias Set to {expectedBiasVoltage}V\n Aquisition timers set to:\n\t Beam on time = {beamConfig.BeamOnTime}\n\t Beam off time = {beamConfig.BeamOffTime}\n\t Pretrigger time = {beamConfig.PreTriggerTime}\n\t Posttrigger time = {beamConfig.PostTriggerTime}");
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

                armed = globalHighVoltageEnabled && globalBiasVoltage == expectedBiasVoltage;

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
                SetLiveMeasurementText($"{lastKnownCharge:0.0000}");
            }

        }

        private void SetLiveMeasurementText(string value)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            if (InvokeRequired)
            {
                if (!IsHandleCreated)
                {
                    return;
                }

                BeginInvoke(new Action<string>(SetLiveMeasurementText), value);
                return;
            }

            Measurement.Text = value;
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

        private void UpdateStatusLabel(bool isCollecting)
        {
            try
            {
                if (IsDisposed || Disposing || statusLabel == null || Measurement == null)
                {
                    return;
                }

                // Ensure this method runs on the UI thread.
                if (InvokeRequired)
                {
                    if (!IsHandleCreated)
                    {
                        return;
                    }

                    BeginInvoke(new Action<bool>(UpdateStatusLabel), isCollecting);
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
                            if (suppressNextMeasurementCommit)
                            {
                                suppressNextMeasurementCommit = false;
                                statusLabel.Text = "Background complete.";
                                statusLabel.BackColor = this.BackColor;
                                statusLabel.ForeColor = Color.Black;
                            }
                            else
                            {
                                AddOrUpdateCurrentMeasurement(enforceDeadTime: true);
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
                        statusLabel.Text = $"DoseX not armed for {expectedBiasVoltage} V.";
                        statusLabel.BackColor = this.BackColor;
                        statusLabel.ForeColor = Color.Black;
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
            catch (Exception ex)
            {
                LogApplicationException("UpdateStatusLabel", ex);
            }
        }




        private void LogRunningState(bool running)
        {
            // Debug.WriteLine($"Running state updated to: {running}");
        }

        private static void LogApplicationException(string context, Exception ex)
        {
            try
            {
                string directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "DoseXPDMTool");
                Directory.CreateDirectory(directory);
                string logPath = Path.Combine(directory, "DoseXPDMTool-errors.log");
                string message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}{Environment.NewLine}{ex}{Environment.NewLine}";
                File.AppendAllText(logPath, message);
            }
            catch
            {
            }
        }

        private async void DiscoverServices(bool showErrors = true)
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
                    if (showErrors)
                    {
                        MessageBox.Show("No DoseX services found.  Please ensure the electrometer is on and check the ethernet conections at both ends.  Then wait 30 seconds before attempting to discover again.");
                    }

                }
            }
            catch (Exception ex)
            {
                if (showErrors)
                {
                    MessageBox.Show($"Error discovering services: {ex.Message}");
                }

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

        private bool EnsureExcelWorkbookReady(bool showErrors = true)
        {
            try
            {
                if (excelApp == null)
                {
                    excelApp = new Excel.Application();
                    excelApp.Visible = true;
                }

                if (machine == null || string.IsNullOrWhiteSpace(machine.Path))
                {
                    statusLabel.Text = "No Excel workbook path is configured.";
                    if (showErrors)
                    {
                        MessageBox.Show("The file path is not specified.");
                    }

                    return false;
                }

                if (workbook == null)
                {
                    workbook = excelApp.Workbooks.Open(machine.Path);
                }

                excelConnected = true;
                return true;
            }
            catch (Exception ex)
            {
                statusLabel.Text = "Excel workbook unavailable.";
                if (showErrors)
                {
                    MessageBox.Show("Failed to open Excel workbook: " + ex.Message);
                }

                return false;
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

        private bool AddOrUpdateCurrentMeasurement(bool enforceDeadTime = false)
        {
            if (enforceDeadTime && !CanAcceptAutomaticMeasurementCommit())
            {
                statusLabel.Text = "Duplicate measurement ignored.";
                statusLabel.BackColor = this.BackColor;
                statusLabel.ForeColor = Color.Black;
                return false;
            }

            if (IsTg51Mode())
            {
                return RecordTg51Charge(Measurement.Text);
            }

            if (!CanCommitPointDoseMeasurement())
            {
                statusLabel.Text = "Choose a point dose test and energy before recording.";
                statusLabel.BackColor = this.BackColor;
                statusLabel.ForeColor = Color.Black;
                return false;
            }

            int rangeSize = GetCurrentRangeCapacity();
            if (rangeSize <= 0)
            {
                statusLabel.Text = "No Excel range is available for this point dose.";
                statusLabel.BackColor = this.BackColor;
                statusLabel.ForeColor = Color.Black;
                return false;
            }

            int indexToUpdate = Results.SelectedIndex;

            if (indexToUpdate >= 0)
            {
                Results.Items[indexToUpdate] = Measurement.Text;
                statusLabel.Text = "";
                statusLabel.BackColor = this.BackColor;
                statusLabel.ForeColor = Color.Black;
                Results.SelectedIndex = -1;

                if (AutoAccept.Checked) { UpdateExcelRange(); }
                MarkMeasurementCommitted();
                return true;
            }

            if (Results.Items.Count < rangeSize)
            {
                Results.Items.Add(Measurement.Text);
                statusLabel.Text = "Data collection complete.";
                statusLabel.BackColor = this.BackColor;
                statusLabel.ForeColor = Color.Black;

                if (AutoAccept.Checked) { UpdateExcelRange(); }
                MarkMeasurementCommitted();
                return true;
            }

            MessageBox.Show("Cannot add more items than the available cells in the named range.");
            return false;
        }

        private bool CanAcceptAutomaticMeasurementCommit()
        {
            return DateTime.Now - lastMeasurementCommitAt >= MeasurementCommitDeadTime;
        }

        private void MarkMeasurementCommitted()
        {
            lastMeasurementCommitAt = DateTime.Now;
        }

        private int GetCurrentRangeCapacity()
        {
            return GetNamedRangeSize(currentNamedRange);
        }

        private bool CanCommitPointDoseMeasurement()
        {
            return TestType?.SelectedIndex >= 0
                && Energy?.SelectedIndex >= 0
                && !string.IsNullOrWhiteSpace(currentNamedRange)
                && Results != null;
        }

        private bool IsTg51Mode()
        {
            return MainTabs?.SelectedTab == Tg51Tab && !string.IsNullOrWhiteSpace(GetSelectedTg51Energy());
        }

        private bool NamedRangeExists(string namedRange)
        {
            if (workbook == null || string.IsNullOrEmpty(namedRange))
            {
                return false;
            }

            try
            {
                Excel.Name excelNamedRange = workbook.Names.Item(namedRange);
                return excelNamedRange?.RefersToRange != null;
            }
            catch
            {
                return false;
            }
        }

        private async Task RunTg51WaitModalAsync(string message, Func<Task> action, int holdMs = 3000)
        {
            using Tg51WaitDialog waitDialog = new Tg51WaitDialog(message);
            waitDialog.Show(this);
            waitDialog.Update();
            Enabled = false;
            Task actionTask = Task.CompletedTask;

            try
            {
                actionTask = action();
                await Task.Delay(Math.Max(0, holdMs));
            }
            finally
            {
                Enabled = true;
                if (!waitDialog.IsDisposed)
                {
                    waitDialog.Close();
                }

                Activate();
            }

            await actionTask;
        }

        private async Task TransitionTg51BiasAsync(int nextBias)
        {
            tg51TransitionInProgress = true;
            try
            {
                string biasLabel = FormatTg51BiasLabel(nextBias);
                statusLabel.Text = $"Changing bias to {biasLabel}...";
                Tg51Status.Text = $"Changing bias to {biasLabel}...";
                statusLabel.BackColor = Color.LightYellow;
                await SetBiasVoltageAsync(nextBias);
                await Task.Delay(Tg51BiasTransitionDelayMs);
                statusLabel.Text = "Beam on";
                Tg51Status.Text = $"Beam on: {biasLabel}.";
                statusLabel.BackColor = Color.LightGreen;
            }
            finally
            {
                tg51TransitionInProgress = false;
            }
        }

        private async Task SetBiasVoltageAsync(int biasVoltage)
        {
            expectedBiasVoltage = biasVoltage;

            if (!doseXConnected)
            {
                MessageBox.Show("Connect Electrometer");
                return;
            }

            bool isControlAvailable = await RequestControlTokenAsync();
            if (!isControlAvailable)
            {
                MessageBox.Show("Please free control from the electrometer by pressing the 'cloud' at the top.");
                return;
            }

            try
            {
                await SendBiasVoltageAsync(biasVoltage);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error setting bias voltage: " + ex.Message);
                Console.WriteLine("Error setting bias voltage: " + ex.Message);
            }
            finally
            {
                await ReleaseControlAsync();
            }
        }

        private async Task SendBiasVoltageAsync(int biasVoltage)
        {
            expectedBiasVoltage = biasVoltage;
            await AuthenticateAdminAsync(PW);

            var config = new
            {
                cmd = "change_values",
                values = new Dictionary<string, object>
                {
                    { "biasVoltage", new { value = biasVoltage, unit = "V" } },
                    { "highVoltageEnabled", true }
                }
            };

            string jsonConfig = JsonConvert.SerializeObject(config);
            var buffer = Encoding.UTF8.GetBytes(jsonConfig);
            var segment = new ArraySegment<byte>(buffer);

            await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
            await ReadHighVoltageAndBiasVoltageAsync();
        }

        private void InitializeTg51Workspace()
        {
            if (Tg51EnergyCombo == null)
            {
                return;
            }

            tg51ChangingSelection = true;
            Tg51EnergyCombo.Items.Clear();
            Tg51EnergyList.Items.Clear();
            foreach (string energy in machine.Energies.Any() ? machine.Energies : GetDefaultEnergyList())
            {
                Tg51EnergyCombo.Items.Add(energy);
                Tg51EnergyList.Items.Add(energy, energy.Contains("x", StringComparison.OrdinalIgnoreCase));
            }
            tg51ChangingSelection = false;

            ConfigureTg51PointGrid();
            ConfigureTg51ReadingGrid();
            Tg51BridgeUrl.Text = Tg51BridgeDefaultUrl;
            Tg51DepthValue.Text = "Depth: --";
            Tg51DepthCmText.Text = "10";
            Tg51MoveDepthAuto.Checked = true;
            Tg51MoveDepthAuto.Enabled = true;
            Tg51Status.Text = "Choose an energy.";
            Tg51BridgeStatus.Text = "Tank not connected";
            tg51TankStatusTimer.Interval = 1000;
            tg51TankStatusTimer.Tick += async (sender, e) => await RefreshTg51TankStatusTimerAsync();
            LoadTg51RunState();
        }

        private List<string> GetDefaultEnergyList()
        {
            return new List<string>
            {
                "4x", "6x", "6xfff", "8x", "10x", "10xfff",
                "15x", "16x", "18x", "20x", "23x", "6e", "9e",
                "12e", "15e", "16e", "18e", "20e", "22e"
            };
        }

        private void ConfigureTg51PointGrid()
        {
            Tg51PointGrid.AutoGenerateColumns = false;
            Tg51PointGrid.AllowUserToAddRows = true;
            Tg51PointGrid.Columns.Clear();
            Tg51PointGrid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "UsePoint", HeaderText = "Use", Width = 40 });
            Tg51PointGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Energy", HeaderText = "Energy", Width = 70 });
            Tg51PointGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Modality", HeaderText = "Mode", Width = 70 });
            Tg51PointGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DepthLabel", HeaderText = "Depth", Width = 90 });
            Tg51PointGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DepthCm", HeaderText = "Depth cm", Width = 75 });
            Tg51PointGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "RawCrossline", HeaderText = "Raw X", Width = 70 });
            Tg51PointGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "RawInline", HeaderText = "Raw Y", Width = 70 });
            Tg51PointGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "RawDepth", HeaderText = "Raw Z", Width = 70 });
            Tg51PointGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "Status", Width = 110, ReadOnly = true });
        }

        private void ConfigureTg51ReadingGrid()
        {
            Tg51ReadingGrid.AutoGenerateColumns = false;
            Tg51ReadingGrid.AllowUserToAddRows = false;
            Tg51ReadingGrid.ReadOnly = true;
            Tg51ReadingGrid.Columns.Clear();
            Tg51ReadingGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Point", HeaderText = "Point", Width = 150 });
            Tg51ReadingGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Bias", HeaderText = "Bias", Width = 70 });
            Tg51ReadingGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Repeat", HeaderText = "#", Width = 40 });
            Tg51ReadingGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Charge", HeaderText = "Charge nC", Width = 90 });
            Tg51ReadingGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Time", HeaderText = "Time", Width = 135 });
        }

        private void Tg51CreateRun_Click(object sender, EventArgs e)
        {
            tg51Run = BuildTg51RunFromGrid();
            tg51Logger = new Tg51RunLogger(machine.Path);
            tg51Logger.Save(tg51Run);
            tg51CurrentPointIndex = tg51Run.Points.Count > 0 ? 0 : -1;
            tg51RunActive = false;
            RefreshTg51Grids();
            Tg51Status.Text = tg51Run.Points.Count == 0
                ? "Add at least one energy/depth point."
                : $"Run ready: {tg51Run.Points.Count} point(s).";
        }

        private Tg51Run BuildTg51RunFromGrid()
        {
            var run = new Tg51Run
            {
                RunId = $"TG51_{DateTime.Now:yyyyMMdd_HHmmss}",
                MachineName = machine.Name ?? string.Empty,
                MachineType = machine.Type ?? string.Empty,
                MachinePath = machine.Path ?? string.Empty,
                StartedAt = DateTime.Now,
                BridgeUrl = Tg51BridgeUrl.Text.Trim(),
                BridgeEnabled = Tg51UseBridge.Checked
            };

            foreach (DataGridViewRow row in Tg51PointGrid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                if (!Convert.ToBoolean(row.Cells["UsePoint"].Value ?? true))
                {
                    continue;
                }

                string energy = row.Cells["Energy"].Value?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(energy))
                {
                    continue;
                }

                string modality = row.Cells["Modality"].Value?.ToString()?.Trim() ?? GuessModality(energy);
                string depthLabel = row.Cells["DepthLabel"].Value?.ToString()?.Trim() ?? "Reference";
                var point = new Tg51Point
                {
                    PointId = $"{run.Points.Count + 1:00}_{energy}_{SanitizeXmlToken(depthLabel)}",
                    Energy = energy,
                    Modality = modality,
                    DepthLabel = depthLabel,
                    ClinicalDepthCm = ParseNullableDouble(row.Cells["DepthCm"].Value),
                    RawCrosslineMm = ParseNullableDouble(row.Cells["RawCrossline"].Value),
                    RawInlineMm = ParseNullableDouble(row.Cells["RawInline"].Value),
                    RawDepthMm = ParseNullableDouble(row.Cells["RawDepth"].Value),
                    Status = "Pending"
                };
                run.Points.Add(point);
            }

            return run;
        }

        private string SanitizeXmlToken(string value)
        {
            return Regex.Replace(value, "[^A-Za-z0-9]+", "_").Trim('_');
        }

        private double? ParseNullableDouble(object value)
        {
            string text = value?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out double result) ||
                   double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
                ? result
                : null;
        }

        private async void Tg51StartPoint_Click(object sender, EventArgs e)
        {
            if (tg51Run == null)
            {
                Tg51CreateRun_Click(sender, e);
            }

            if (tg51Run == null || tg51Run.Points.Count == 0)
            {
                MessageBox.Show("Create a TG-51 run with at least one point first.");
                return;
            }

            if (tg51CurrentPointIndex < 0 || tg51CurrentPointIndex >= tg51Run.Points.Count)
            {
                tg51CurrentPointIndex = Math.Max(0, tg51Run.Points.FindIndex(p => p.Status != "Complete"));
            }

            await PrepareTg51PointAsync(tg51CurrentPointIndex);
        }

        private async Task PrepareTg51PointAsync(int pointIndex)
        {
            if (pointIndex < 0 || tg51Run == null || pointIndex >= tg51Run.Points.Count)
            {
                Tg51Status.Text = "TG-51 run complete.";
                tg51RunActive = false;
                return;
            }

            var point = tg51Run.Points[pointIndex];
            tg51CurrentPointIndex = pointIndex;
            point.Status = "Preparing";
            RefreshTg51Grids();

            if (Tg51UseBridge.Checked)
            {
                bool moved = await MoveTankForTg51PointAsync(point);
                if (!moved)
                {
                    point.Status = "Tank needed";
                    tg51RunActive = false;
                    RefreshTg51Grids();
                    SaveTg51Run();
                    return;
                }
            }

            int nextBias = GetActiveTg51Bias(point);
            expectedBiasVoltage = nextBias;
            await SetBiasVoltageAsync(nextBias);
            await SetMeasurementConfigurationAsync();
            await StartMeasurementAsync();

            point.Status = "Active";
            tg51RunActive = true;
            RefreshTg51Grids();
            SaveTg51Run();
            int nextRepeat = GetTg51BiasReadingCount(point, nextBias) + 1;
            Tg51Status.Text = $"Beam on: {point.Energy} {point.DepthLabel}, {FormatTg51BiasLabel(nextBias)}, reading {nextRepeat}.";
        }

        private async Task<bool> MoveTankForTg51PointAsync(Tg51Point point)
        {
            if (!point.RawCrosslineMm.HasValue || !point.RawInlineMm.HasValue || !point.RawDepthMm.HasValue)
            {
                MessageBox.Show("Bridge movement requires Raw X, Raw Y, and Raw Z values for the selected TG-51 point.");
                return false;
            }

            string baseUrl = Tg51BridgeUrl.Text.Trim().TrimEnd('/');
            if (string.IsNullOrEmpty(baseUrl))
            {
                MessageBox.Show("Enter the tank bridge URL before enabling bridge movement.");
                return false;
            }

            try
            {
                Tg51Status.Text = "Checking tank bridge...";
                await EnsureTg51TankConnectedAsync(baseUrl, TimeSpan.FromSeconds(30));

                var moveBody = JsonConvert.SerializeObject(new
                {
                    crossline = point.RawCrosslineMm.Value,
                    inline = point.RawInlineMm.Value,
                    depth = point.RawDepthMm.Value,
                    speed = 20.0
                });
                Tg51Status.Text = $"Moving tank to {point.RawDepthMm:0.###} mm...";
                using (HttpResponseMessage moveResponse = await tg51TankClient.PostAsync($"{baseUrl}/api/move", new StringContent(moveBody, Encoding.UTF8, "application/json")))
                {
                    await EnsureBridgeSuccessAsync(moveResponse, "move tank");
                }

                JObject finalState = await WaitForTankPositionAsync(baseUrl, point);
                point.LastBridgeSnapshotJson = finalState.ToString(Formatting.None);
                Tg51BridgeStatus.Text = "Tank ready";
                return true;
            }
            catch (Exception ex)
            {
                Tg51BridgeStatus.Text = "Bridge error";
                MessageBox.Show("Tank bridge movement failed: " + ex.Message);
                return false;
            }
        }

        private async Task<bool> ConnectTg51TankBridgeAsync()
        {
            string baseUrl = GetTg51BridgeBaseUrl();
            if (string.IsNullOrEmpty(baseUrl))
            {
                MessageBox.Show("Enter the tank bridge URL first.");
                return false;
            }

            try
            {
                Tg51BridgeStatus.Text = "Checking tank...";
                JObject state = await EnsureTg51TankConnectedAsync(baseUrl, TimeSpan.FromSeconds(30));
                bool connected = state["connected"]?.ToObject<bool>() ?? false;
                JToken status = state["latestStatus"];
                double? x = ReadBridgeNumber(status, "x");
                double? y = ReadBridgeNumber(status, "y");
                double? z = ReadBridgeNumber(status, "z");
                Tg51BridgeStatus.Text = connected && x.HasValue && y.HasValue && z.HasValue
                    ? FormatTankPositionStatus(state)
                    : connected ? "Tank connected" : "Bridge reachable";
                if (connected)
                {
                    tg51TankStatusTimer.Start();
                }

                return connected;
            }
            catch (TimeoutException)
            {
                Tg51BridgeStatus.Text = "Connected; waiting for position";
                return true;
            }
            catch (Exception ex)
            {
                Tg51BridgeStatus.Text = "Tank unavailable";
                MessageBox.Show("Tank bridge connection failed: " + ex.Message);
                return false;
            }
        }

        private async Task MoveSelectedTg51DepthAsync()
        {
            var point = GetOrCreateSelectedTg51Point();
            if (point == null)
            {
                Tg51Status.Text = "Choose an energy before moving depth.";
                return;
            }

            if (!UpdateTg51DepthFromText(point))
            {
                return;
            }

            string baseUrl = GetTg51BridgeBaseUrl();
            if (string.IsNullOrEmpty(baseUrl))
            {
                MessageBox.Show("Enter the tank bridge URL first.");
                return;
            }

            if (tg51MoveInProgress)
            {
                return;
            }

            try
            {
                tg51MoveInProgress = true;
                Tg51MoveDepth.Enabled = false;
                Tg51BridgeStatus.Text = "Reading tank position...";
                JObject state = await EnsureTg51TankConnectedAsync(baseUrl, TimeSpan.FromSeconds(5));

                JToken isocenter = state["isocenter"];
                double? isoX = ReadBridgeNumber(isocenter, "x");
                double? isoY = ReadBridgeNumber(isocenter, "y");
                double? isoZ = ReadBridgeNumber(isocenter, "z");
                if (!isoX.HasValue || !isoY.HasValue || !isoZ.HasValue)
                {
                    Tg51BridgeStatus.Text = "No tank isocenter";
                    MessageBox.Show("The tank bridge does not have a saved isocenter yet. Set/check isocenter in the tank workflow before moving TG-51 depth.");
                    return;
                }

                double depthMm = point.ClinicalDepthCm.Value * 10.0;
                double x = isoX.Value;
                double y = isoY.Value;
                double targetZ = isoZ.Value - depthMm;

                point.RawCrosslineMm = x;
                point.RawInlineMm = y;
                point.RawDepthMm = targetZ;

                var moveBody = JsonConvert.SerializeObject(new
                {
                    crossline = x,
                    inline = y,
                    depth = targetZ,
                    speed = 20.0
                });

                Tg51BridgeStatus.Text = $"Sending move to {point.ClinicalDepthCm:0.##} cm below iso...";
                using (HttpResponseMessage moveResponse = await tg51TankClient.PostAsync($"{baseUrl}/api/move", new StringContent(moveBody, Encoding.UTF8, "application/json")))
                {
                    await EnsureBridgeSuccessAsync(moveResponse, "move tank");
                }

                JObject latestState = await GetTankBridgeStateAsync(baseUrl);
                point.LastBridgeSnapshotJson = latestState.ToString(Formatting.None);
                SaveTg51Run();
                RefreshTg51View();
                Tg51BridgeStatus.Text = FormatTankPositionStatus(latestState);
                Tg51Status.Text = $"Move sent: {point.Energy}, {point.ClinicalDepthCm:0.##} cm.";
            }
            catch (Exception ex)
            {
                Tg51BridgeStatus.Text = "Move failed";
                MessageBox.Show("Tank depth movement failed: " + ex.Message);
            }
            finally
            {
                tg51MoveInProgress = false;
                Tg51MoveDepth.Enabled = true;
            }
        }

        private async Task RefreshSelectedTg51TankStateAsync()
        {
            string baseUrl = GetTg51BridgeBaseUrl();
            if (string.IsNullOrEmpty(baseUrl))
            {
                return;
            }

            try
            {
                JObject latestState = await GetTankBridgeStateAsync(baseUrl);
                if (GetCurrentTg51Point() is { } point)
                {
                    point.LastBridgeSnapshotJson = latestState.ToString(Formatting.None);
                    SaveTg51Run();
                    RefreshTg51View();
                }

                Tg51BridgeStatus.Text = FormatTankPositionStatus(latestState, GetCurrentTg51Point());
            }
            catch (Exception ex)
            {
                Tg51BridgeStatus.Text = "Tank status refresh failed";
                Console.WriteLine("Tank status refresh failed: " + ex.Message);
            }
        }

        private async Task RefreshTg51TankStatusTimerAsync()
        {
            if (tg51TankStatusRefreshInProgress || MainTabs?.SelectedTab != Tg51Tab)
            {
                return;
            }

            string baseUrl = GetTg51BridgeBaseUrl();
            if (string.IsNullOrEmpty(baseUrl))
            {
                return;
            }

            try
            {
                tg51TankStatusRefreshInProgress = true;
                JObject latestState = await GetTankBridgeStateAsync(baseUrl);
                Tg51BridgeStatus.Text = FormatTankPositionStatus(latestState, GetCurrentTg51Point());
            }
            catch (Exception ex)
            {
                Console.WriteLine("Tank status timer refresh failed: " + ex.Message);
            }
            finally
            {
                tg51TankStatusRefreshInProgress = false;
            }
        }

        private string GetTg51BridgeBaseUrl()
        {
            string baseUrl = Tg51BridgeUrl.Text.Trim().TrimEnd('/');
            return string.IsNullOrWhiteSpace(baseUrl) ? Tg51BridgeDefaultUrl : baseUrl;
        }

        private bool UpdateTg51DepthFromText(Tg51Point point)
        {
            if (!TryGetTg51DepthCmFromText(out double depthCm))
            {
                Tg51Status.Text = "Enter a TG-51 depth in cm before moving depth.";
                return false;
            }

            if (depthCm < 0 || depthCm > 50)
            {
                Tg51Status.Text = "TG-51 depth must be between 0 and 50 cm.";
                return false;
            }

            point.ClinicalDepthCm = depthCm;
            point.DepthLabel = $"{depthCm:0.##} cm";
            return true;
        }

        private bool TryGetTg51DepthCmFromText(out double depthCm)
        {
            string depthText = Tg51DepthCmText.Text.Trim();
            return double.TryParse(depthText, NumberStyles.Float, CultureInfo.CurrentCulture, out depthCm) ||
                   double.TryParse(depthText, NumberStyles.Float, CultureInfo.InvariantCulture, out depthCm);
        }

        private void SetTg51DepthTextToDefaultForEnergy(string energy)
        {
            double? defaultDepth = GetDefaultTg51DepthCm(energy);
            if (!defaultDepth.HasValue)
            {
                Tg51DepthCmText.Text = string.Empty;
                Tg51DepthValue.Text = "Depth: --";
                tg51DepthEnergy = energy;
                return;
            }

            Tg51DepthCmText.Text = defaultDepth.Value.ToString("0.##", CultureInfo.CurrentCulture);
            Tg51DepthValue.Text = $"Depth: {defaultDepth.Value:0.##} cm";
            tg51DepthEnergy = energy;
        }

        private async Task<JObject> GetTankBridgeStateAsync(string baseUrl)
        {
            string stateJson = await tg51TankClient.GetStringAsync($"{baseUrl}/api/state");
            return JObject.Parse(stateJson);
        }

        private async Task<JObject> EnsureTg51TankConnectedAsync(string baseUrl, TimeSpan livePositionTimeout)
        {
            await EnsureTg51BridgeRuntimeReadyAsync(baseUrl);

            JObject state = await GetTankBridgeStateAsync(baseUrl);
            bool connected = state["connected"]?.ToObject<bool>() ?? false;

            if (!connected)
            {
                Tg51BridgeStatus.Text = "Connecting tank...";
                using HttpResponseMessage response = await tg51TankClient.PostAsync($"{baseUrl}/api/connect", new StringContent("", Encoding.UTF8, "application/json"));
                await EnsureBridgeSuccessAsync(response, "connect tank");
            }
            else
            {
                Tg51BridgeStatus.Text = "Tank already connected";
            }

            return await WaitForLiveTankStateAsync(baseUrl, livePositionTimeout);
        }

        private async Task EnsureTg51BridgeRuntimeReadyAsync(string baseUrl)
        {
            Tg51BridgeStatus.Text = "Checking bridge...";

            if (!await IsExpectedTg51BridgeAsync(baseUrl))
            {
                await StartOrInstallTg51BridgeAsync(baseUrl);
            }

            if (!await IsExpectedTg51BridgeAsync(baseUrl))
            {
                throw new InvalidOperationException("The tank bridge is running, but it did not identify as the expected TG-51 tank bridge.");
            }

            Tg51BridgeStatus.Text = "Checking firewall...";
            var firewall = await CheckTg51BridgeFirewallAsync();
            if (!firewall.IsOk)
            {
                bool repaired = await OfferTg51FirewallRepairAsync(firewall.Message);
                if (!repaired)
                {
                    throw new InvalidOperationException("The TG-51 tank bridge firewall rule is not ready. " + firewall.Message);
                }

                firewall = await CheckTg51BridgeFirewallAsync();
                if (!firewall.IsOk)
                {
                    throw new InvalidOperationException("The TG-51 tank bridge firewall rule still is not ready after the repair attempt. " + firewall.Message);
                }
            }
        }

        private async Task<bool> IsExpectedTg51BridgeAsync(string baseUrl)
        {
            try
            {
                using HttpResponseMessage health = await tg51TankClient.GetAsync($"{baseUrl}/health");
                if (!health.IsSuccessStatusCode)
                {
                    return false;
                }

                string capabilitiesJson = await tg51TankClient.GetStringAsync($"{baseUrl}/capabilities");
                JObject capabilities = JObject.Parse(capabilitiesJson);
                string schemaVersion = capabilities["schemaVersion"]?.ToString() ?? string.Empty;
                if (!string.Equals(schemaVersion, Tg51BridgeCapabilitiesSchema, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                var endpoints = capabilities["implementedControlEndpoints"] as JArray;
                return endpoints != null
                    && endpoints.Any(e => string.Equals(e?.ToString(), "POST /api/connect", StringComparison.OrdinalIgnoreCase))
                    && endpoints.Any(e => string.Equals(e?.ToString(), "POST /api/move", StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return false;
            }
        }

        private async Task StartOrInstallTg51BridgeAsync(string baseUrl)
        {
            string installedLauncher = GetInstalledTg51BridgeLauncherPath();
            if (File.Exists(installedLauncher))
            {
                Tg51BridgeStatus.Text = "Starting bridge...";
                await RunProcessAsync(installedLauncher, string.Empty, false);
                if (await WaitForExpectedTg51BridgeAsync(baseUrl, TimeSpan.FromSeconds(15)))
                {
                    return;
                }
            }

            string bundledInstaller = GetBundledTg51BridgeInstallerPath();
            if (!File.Exists(bundledInstaller))
            {
                throw new FileNotFoundException("The bundled Tank Bridge Runtime installer was not found beside DoseXPDMTool.", bundledInstaller);
            }

            DialogResult result = MessageBox.Show(
                "The TG-51 Tank Bridge Runtime needs to be installed or repaired. This will ask for Administrator approval so it can install the bridge and create the Windows Firewall rule.",
                "Install TG-51 Tank Bridge",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);
            if (result != DialogResult.OK)
            {
                throw new InvalidOperationException("TG-51 Tank Bridge setup was cancelled.");
            }

            Tg51BridgeStatus.Text = "Running bridge installer...";
            await RunProcessAsync(bundledInstaller, string.Empty, true);

            installedLauncher = GetInstalledTg51BridgeLauncherPath();
            if (!File.Exists(installedLauncher))
            {
                throw new FileNotFoundException("The TG-51 Tank Bridge installer completed, but the installed launcher was not found.", installedLauncher);
            }

            Tg51BridgeStatus.Text = "Starting bridge...";
            await RunProcessAsync(installedLauncher, string.Empty, false);
            if (!await WaitForExpectedTg51BridgeAsync(baseUrl, TimeSpan.FromSeconds(20)))
            {
                throw new TimeoutException("The TG-51 Tank Bridge did not answer after installation.");
            }
        }

        private async Task<bool> WaitForExpectedTg51BridgeAsync(string baseUrl, TimeSpan timeout)
        {
            DateTime deadline = DateTime.Now.Add(timeout);
            while (DateTime.Now < deadline)
            {
                if (await IsExpectedTg51BridgeAsync(baseUrl))
                {
                    return true;
                }

                await Task.Delay(500);
            }

            return false;
        }

        private string GetInstalledTg51BridgeRoot()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), Tg51BridgeInstallFolderName);
        }

        private string GetInstalledTg51BridgeExePath()
        {
            return Path.Combine(GetInstalledTg51BridgeRoot(), "bridge", "TankControllerBridge.exe");
        }

        private string GetInstalledTg51BridgeLauncherPath()
        {
            return Path.Combine(GetInstalledTg51BridgeRoot(), "Start-TankBridgeRuntime.cmd");
        }

        private string GetBundledTg51BridgeRuntimeRoot()
        {
            return Path.Combine(AppContext.BaseDirectory, Tg51BridgeRuntimeFolderName);
        }

        private string GetBundledTg51BridgeInstallerPath()
        {
            return Path.Combine(GetBundledTg51BridgeRuntimeRoot(), "Install-TankBridgeRuntime.cmd");
        }

        private string GetBundledTg51BridgeFirewallHelperPath()
        {
            return Path.Combine(GetBundledTg51BridgeRuntimeRoot(), "Allow-TankBridgeRuntimeFirewall.cmd");
        }

        private string GetInstalledTg51BridgeFirewallHelperPath()
        {
            return Path.Combine(GetInstalledTg51BridgeRoot(), "Allow-TankBridgeRuntimeFirewall.cmd");
        }

        private async Task<(bool IsOk, string Message)> CheckTg51BridgeFirewallAsync()
        {
            string bridgeExe = GetInstalledTg51BridgeExePath();
            if (!File.Exists(bridgeExe))
            {
                return (false, $"Installed bridge executable was not found at {bridgeExe}.");
            }

            string command = BuildFirewallCheckPowerShell(bridgeExe);
            var result = await RunProcessCaptureAsync("powershell.exe", $"-NoProfile -ExecutionPolicy Bypass -Command \"{command}\"", false);
            string detail = string.IsNullOrWhiteSpace(result.Output) ? result.Error : result.Output;
            return result.ExitCode == 0
                ? (true, "Firewall rule is ready.")
                : (false, string.IsNullOrWhiteSpace(detail) ? "Firewall rule was not found or did not match the installed bridge." : detail.Trim());
        }

        private string BuildFirewallCheckPowerShell(string bridgeExe)
        {
            string escapedExe = bridgeExe.Replace("'", "''");
            return "$ErrorActionPreference='SilentlyContinue';" +
                   $"$exe='{escapedExe}';" +
                   $"$rules=@(Get-NetFirewallRule -DisplayName '{Tg51BridgeFirewallRuleName}' | Where-Object {{$_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' -and $_.Action -eq 'Allow'}});" +
                   "foreach($rule in $rules){" +
                   "$apps=@(Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $rule);" +
                   "$ports=@(Get-NetFirewallPortFilter -AssociatedNetFirewallRule $rule);" +
                   "$addrs=@(Get-NetFirewallAddressFilter -AssociatedNetFirewallRule $rule);" +
                   "$programOk=$false; foreach($app in $apps){if($app.Program -and ([System.IO.Path]::GetFullPath($app.Program) -ieq [System.IO.Path]::GetFullPath($exe))){$programOk=$true}};" +
                   "$tcpOk=$false; foreach($port in $ports){if([string]$port.Protocol -eq 'TCP'){$tcpOk=$true}};" +
                   "$remoteOk=$false; foreach($addr in $addrs){" +
                   "$remoteText=($addr.RemoteAddress -join ',');" +
                   "if($remoteText -match '(^|,)169\\.254\\.0\\.0/16(,|$)' -or $remoteText -match '(^|,)169\\.254\\.0\\.0/255\\.255\\.0\\.0(,|$)' -or $remoteText -match '(^|,)169\\.254\\.0\\.0-169\\.254\\.255\\.255(,|$)'){$remoteOk=$true}" +
                   "};" +
                   "if($programOk -and $tcpOk -and $remoteOk){Write-Output 'OK'; exit 0}" +
                   "}" +
                   $"Write-Output 'No enabled inbound TCP firewall rule named {Tg51BridgeFirewallRuleName} points to the installed bridge executable and allows 169.254.0.0/16.'; exit 2";
        }

        private async Task<bool> OfferTg51FirewallRepairAsync(string reason)
        {
            string helper = GetInstalledTg51BridgeFirewallHelperPath();
            string installer = GetBundledTg51BridgeInstallerPath();
            string repairPath = File.Exists(helper) ? helper : installer;
            if (!File.Exists(repairPath))
            {
                MessageBox.Show("The tank bridge firewall rule is not ready, and the bundled firewall repair tool was not found.\n\n" + reason);
                return false;
            }

            DialogResult result = MessageBox.Show(
                "The tank bridge is installed, but the Windows Firewall rule for the tank callback is missing or does not match the installed bridge.\n\n" +
                reason + "\n\nRun the bundled firewall repair as Administrator now?",
                "Repair TG-51 Tank Bridge Firewall",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning);
            if (result != DialogResult.OK)
            {
                return false;
            }

            Tg51BridgeStatus.Text = "Repairing firewall...";
            await RunProcessAsync(repairPath, string.Empty, true);
            return true;
        }

        private async Task RunProcessAsync(string fileName, string arguments, bool elevate)
        {
            await Task.Run(() =>
            {
                bool useShellExecute = elevate || string.Equals(Path.GetExtension(fileName), ".cmd", StringComparison.OrdinalIgnoreCase);
                using Process process = new Process();
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = useShellExecute,
                    Verb = elevate ? "runas" : string.Empty,
                    CreateNoWindow = !elevate,
                    WindowStyle = elevate ? ProcessWindowStyle.Normal : ProcessWindowStyle.Hidden,
                    WorkingDirectory = Path.GetDirectoryName(fileName) ?? AppContext.BaseDirectory
                };
                process.Start();
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException($"{Path.GetFileName(fileName)} exited with code {process.ExitCode}.");
                }
            });
        }

        private async Task<(int ExitCode, string Output, string Error)> RunProcessCaptureAsync(string fileName, string arguments, bool elevate)
        {
            return await Task.Run(() =>
            {
                using Process process = new Process();
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = elevate,
                    RedirectStandardOutput = !elevate,
                    RedirectStandardError = !elevate,
                    CreateNoWindow = !elevate,
                    WindowStyle = elevate ? ProcessWindowStyle.Normal : ProcessWindowStyle.Hidden
                };
                process.Start();
                string output = elevate ? string.Empty : process.StandardOutput.ReadToEnd();
                string error = elevate ? string.Empty : process.StandardError.ReadToEnd();
                process.WaitForExit();
                return (process.ExitCode, output, error);
            });
        }

        private async Task EnsureBridgeSuccessAsync(HttpResponseMessage response, string action)
        {
            if (response.IsSuccessStatusCode)
            {
                return;
            }

            string body = await response.Content.ReadAsStringAsync();
            string detail = string.IsNullOrWhiteSpace(body)
                ? response.ReasonPhrase
                : body.Trim();
            throw new InvalidOperationException($"Bridge failed to {action}: {(int)response.StatusCode} {response.ReasonPhrase}. {detail}");
        }

        private async Task DisconnectTg51TankBridgeAsync()
        {
            string baseUrl = GetTg51BridgeBaseUrl();
            if (string.IsNullOrEmpty(baseUrl))
            {
                MessageBox.Show("Enter the tank bridge URL first.");
                return;
            }

            try
            {
                Tg51BridgeStatus.Text = "Disconnecting tank...";
                using HttpResponseMessage response = await tg51TankClient.PostAsync($"{baseUrl}/api/disconnect", new StringContent("", Encoding.UTF8, "application/json"));
                await EnsureBridgeSuccessAsync(response, "disconnect tank");
                tg51TankStatusTimer.Stop();
                Tg51BridgeStatus.Text = "Tank disconnected";
            }
            catch (Exception ex)
            {
                Tg51BridgeStatus.Text = "Disconnect failed";
                MessageBox.Show("Tank bridge disconnect failed: " + ex.Message);
            }
        }

        private async Task SetSelectedTg51TargetFromCurrentPositionAsync()
        {
            var point = GetOrCreateSelectedTg51Point();
            if (point == null)
            {
                Tg51Status.Text = "Choose an energy before setting a tank target.";
                return;
            }

            string baseUrl = GetTg51BridgeBaseUrl();
            if (string.IsNullOrEmpty(baseUrl))
            {
                MessageBox.Show("Enter the tank bridge URL first.");
                return;
            }

            try
            {
                Tg51BridgeStatus.Text = "Reading tank target...";
                JObject state = await EnsureTg51TankConnectedAsync(baseUrl, TimeSpan.FromSeconds(30));
                JToken status = state["latestStatus"];
                double? x = ReadBridgeNumber(status, "x");
                double? y = ReadBridgeNumber(status, "y");
                double? z = ReadBridgeNumber(status, "z");

                if (!x.HasValue || !y.HasValue || !z.HasValue)
                {
                    Tg51BridgeStatus.Text = "No live tank position";
                    MessageBox.Show("No live tank position is available to save as the TG-51 target.");
                    return;
                }

                point.RawCrosslineMm = x.Value;
                point.RawInlineMm = y.Value;
                point.RawDepthMm = z.Value;
                point.LastBridgeSnapshotJson = state.ToString(Formatting.None);
                point.Status = HasTg51MinimumReadings(point) ? "Complete" : "Active";

                SaveTg51Run();
                RefreshTg51View();
                Tg51BridgeStatus.Text = $"Target saved X {x:0.0}, Y {y:0.0}, Z {z:0.0}";
            }
            catch (Exception ex)
            {
                Tg51BridgeStatus.Text = "Set target failed";
                MessageBox.Show("Tank target save failed: " + ex.Message);
            }
        }

        private async Task<JObject> WaitForLiveTankStateAsync(string baseUrl, TimeSpan timeout)
        {
            DateTime deadline = DateTime.Now.Add(timeout);
            JObject latest = new JObject();

            while (DateTime.Now < deadline)
            {
                latest = await GetTankBridgeStateAsync(baseUrl);
                JToken status = latest["latestStatus"];
                double? x = ReadBridgeNumber(status, "x");
                double? y = ReadBridgeNumber(status, "y");
                double? z = ReadBridgeNumber(status, "z");

                if (x.HasValue && y.HasValue && z.HasValue && HasRecentTg51Callback(latest))
                {
                    return latest;
                }

                bool connected = latest["connected"]?.ToObject<bool>() ?? false;
                Tg51BridgeStatus.Text = connected
                    ? x.HasValue && y.HasValue && z.HasValue ? "Waiting for live tank callback" : "Waiting for tank position"
                    : "Tank not connected";
                await Task.Delay(500);
            }

            throw new TimeoutException("The tank bridge connected, but live tank callback telemetry did not arrive. Windows Firewall may be blocking the CCU callback to the TG-51 Tank Bridge. Run the bundled Tank Bridge Runtime installer or firewall repair as Administrator, then restart the bridge and connect again.");
        }

        private bool HasRecentTg51Callback(JObject state)
        {
            string lastCallbackText = state["lastCallbackAt"]?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(lastCallbackText) ||
                !DateTimeOffset.TryParse(lastCallbackText, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTimeOffset lastCallbackAt))
            {
                return false;
            }

            return DateTimeOffset.Now - lastCallbackAt <= Tg51LiveCallbackFreshness;
        }

        private double? ReadBridgeNumber(JToken token, string name)
        {
            if (token is JObject obj && obj.TryGetValue(name, out JToken value))
            {
                return value.Type == JTokenType.Null ? null : value.ToObject<double?>();
            }

            return null;
        }

        private string FormatTankPositionStatus(JObject state)
        {
            return FormatTankPositionStatus(state, GetCurrentTg51Point());
        }

        private string FormatTankPositionStatus(JObject state, Tg51Point point)
        {
            JToken status = state["latestStatus"];
            JToken isocenter = state["isocenter"];
            double? liveX = ReadBridgeNumber(status, "x");
            double? liveY = ReadBridgeNumber(status, "y");
            double? liveZ = ReadBridgeNumber(status, "z");
            double? isoX = ReadBridgeNumber(isocenter, "x");
            double? isoY = ReadBridgeNumber(isocenter, "y");
            double? isoZ = ReadBridgeNumber(isocenter, "z");
            double? targetX = null;
            double? targetY = null;
            double? targetZ = null;
            if (isoX.HasValue && isoY.HasValue && isoZ.HasValue
                && TryGetTg51DepthCmFromText(out double depthCm)
                && depthCm >= 0
                && depthCm <= 50)
            {
                targetX = isoX.Value;
                targetY = isoY.Value;
                targetZ = isoZ.Value - (depthCm * 10.0);
            }

            string liveText = liveX.HasValue && liveY.HasValue && liveZ.HasValue
                ? $"Live X {liveX:0.0}, Y {liveY:0.0}, Z {liveZ:0.0}"
                : "Live --";
            string isoText = isoX.HasValue && isoY.HasValue && isoZ.HasValue
                ? $"Iso X {isoX:0.0}, Y {isoY:0.0}, Z {isoZ:0.0}"
                : "Iso --";
            string targetText = targetX.HasValue && targetY.HasValue && targetZ.HasValue
                    ? $"Target X {targetX:0.0}, Y {targetY:0.0}, Z {targetZ:0.0}"
                    : "Target --";
            return $"{liveText}{Environment.NewLine}{targetText}{Environment.NewLine}{isoText}";
        }

        private async Task<JObject> WaitForTankPositionAsync(string baseUrl, Tg51Point point)
        {
            var deadline = DateTime.Now.AddSeconds(90);
            JObject latest = new JObject();
            int stableCount = 0;

            while (DateTime.Now < deadline)
            {
                string stateJson = await tg51TankClient.GetStringAsync($"{baseUrl}/api/state");
                latest = JObject.Parse(stateJson);
                var status = latest["latestStatus"];
                double? x = ReadBridgeNumber(status, "x");
                double? y = ReadBridgeNumber(status, "y");
                double? z = ReadBridgeNumber(status, "z");
                bool atTarget = IsTankAtTg51Target(latest, point, requireNotBusy: true);

                stableCount = atTarget ? stableCount + 1 : 0;
                Tg51BridgeStatus.Text = x.HasValue
                    ? FormatTankPositionStatus(latest)
                    : "Waiting for tank position";

                if (stableCount >= 6)
                {
                    return latest;
                }

                await Task.Delay(500);
            }

            if (IsTankAtTg51Target(latest, point, requireNotBusy: false))
            {
                Tg51BridgeStatus.Text = FormatTankPositionStatus(latest) + " (reached; settling not confirmed)";
                return latest;
            }

            throw new TimeoutException("Tank did not reach and hold the requested position. Avoid sending a new TG-51 move until the current move finishes settling.");
        }

        private bool IsTankAtTg51Target(JObject state, Tg51Point point, bool requireNotBusy)
        {
            var status = state["latestStatus"];
            bool busy = state["busy"]?.ToObject<bool>() ?? false;

            return
                (!requireNotBusy || !busy) &&
                IsTankPositionAtTarget(state, point);
        }

        private bool IsTankPositionAtTarget(JObject state, Tg51Point point)
        {
            var status = state["latestStatus"];
            double? x = ReadBridgeNumber(status, "x");
            double? y = ReadBridgeNumber(status, "y");
            double? z = ReadBridgeNumber(status, "z");

            return
                x.HasValue && y.HasValue && z.HasValue &&
                Math.Abs(x.Value - point.RawCrosslineMm.GetValueOrDefault()) <= 0.5 &&
                Math.Abs(y.Value - point.RawInlineMm.GetValueOrDefault()) <= 0.5 &&
                Math.Abs(z.Value - point.RawDepthMm.GetValueOrDefault()) <= 0.5;
        }

        private bool RecordTg51Charge(string chargeText)
        {
            var point = GetOrCreateSelectedTg51Point();
            if (point == null)
            {
                Tg51Status.Text = "Choose an energy before accepting a TG-51 reading.";
                return false;
            }

            if (!double.TryParse(chargeText, NumberStyles.Float, CultureInfo.CurrentCulture, out double charge) &&
                !double.TryParse(chargeText, NumberStyles.Float, CultureInfo.InvariantCulture, out charge))
            {
                return false;
            }

            if (!UpdateTg51DepthFromText(point))
            {
                return false;
            }

            if (pendingTg51OverwritePoint == point &&
                pendingTg51OverwriteReading != null &&
                pendingTg51OverwriteIndex >= 0)
            {
                int overwriteBias = pendingTg51OverwriteBias;
                int overwriteIndex = pendingTg51OverwriteIndex;
                Tg51Reading overwriteReading = pendingTg51OverwriteReading;

                pendingTg51OverwritePoint = null;
                pendingTg51OverwriteReading = null;
                pendingTg51OverwriteBias = 0;
                pendingTg51OverwriteIndex = -1;

                overwriteReading.ChargeNc = charge;
                overwriteReading.RecordedAt = DateTime.Now;
                overwriteReading.DoseXHighVoltageEnabled = globalHighVoltageEnabled;
                overwriteReading.DoseXBiasVoltage = globalBiasVoltage;
                overwriteReading.BridgeSnapshotJson = point.LastBridgeSnapshotJson ?? string.Empty;
                RenumberTg51BiasReadings(point, overwriteBias);
                point.Status = HasTg51MinimumReadings(point) ? "Complete" : "Active";
                tg51RunActive = true;
                SaveTg51Run();
                RefreshTg51View();
                RestoreTg51Selection(overwriteBias, overwriteIndex);
                MarkMeasurementCommitted();
                Tg51Status.Text = $"TG-51 reading overwritten from next measurement: {FormatTg51BiasLabel(overwriteBias)} #{overwriteIndex + 1}.";
                return true;
            }

            int bias = GetActiveTg51Bias(point);
            int repeat = GetTg51BiasReadingCount(point, bias) + 1;

            point.Readings.Add(new Tg51Reading
            {
                BiasVoltage = bias,
                RepeatNumber = repeat,
                ChargeNc = charge,
                RecordedAt = DateTime.Now,
                DoseXHighVoltageEnabled = globalHighVoltageEnabled,
                DoseXBiasVoltage = globalBiasVoltage,
                BridgeSnapshotJson = point.LastBridgeSnapshotJson ?? string.Empty
            });

            point.Status = HasTg51MinimumReadings(point) ? "Complete" : "Active";
            tg51RunActive = true;
            SaveTg51Run();
            RefreshTg51View();
            MarkMeasurementCommitted();
            _ = AdvanceTg51AfterReadingAsync(point, bias);
            return true;
        }

        private async Task AdvanceTg51AfterReadingAsync(Tg51Point point, int recordedBias)
        {
            int bucketCount = GetTg51BiasReadingCount(point, recordedBias);
            if (bucketCount == Tg51ReadingsPerBias)
            {
                int nextBias = GetNextTg51Bias(recordedBias);
                if (nextBias != recordedBias)
                {
                    await Task.Delay(500);
                    await RunTg51WaitModalAsync(
                        $"Wait while bias stabilizes at {FormatTg51BiasLabel(nextBias)}.",
                        async () => await SetActiveTg51BiasAsync(point, nextBias, transition: false),
                        Tg51ManualBiasSettleDelayMs);
                    return;
                }
            }

            RefreshTg51View();
        }

        private Tg51Point GetCurrentTg51Point()
        {
            return tg51Run != null && tg51CurrentPointIndex >= 0 && tg51CurrentPointIndex < tg51Run.Points.Count
                ? tg51Run.Points[tg51CurrentPointIndex]
                : null;
        }

        private string GetSelectedTg51Energy()
        {
            string energy = Tg51EnergyCombo?.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(energy))
            {
                energy = Tg51EnergyCombo?.SelectedItem?.ToString()?.Trim() ?? string.Empty;
            }

            return energy;
        }

        private Tg51Point GetOrCreateSelectedTg51Point()
        {
            string energy = GetSelectedTg51Energy();
            if (string.IsNullOrEmpty(energy))
            {
                return null;
            }

            if (tg51Run == null)
            {
                tg51Run = new Tg51Run
                {
                    RunId = $"TG51_{DateTime.Now:yyyyMMdd_HHmmss}",
                    MachineName = machine.Name ?? string.Empty,
                    MachineType = machine.Type ?? string.Empty,
                    MachinePath = machine.Path ?? string.Empty,
                    StartedAt = DateTime.Now,
                    BridgeUrl = Tg51BridgeUrl.Text.Trim(),
                    BridgeEnabled = false
                };
                tg51Logger = new Tg51RunLogger(machine.Path);
            }

            var point = tg51Run.Points.FirstOrDefault(p => string.Equals(p.Energy, energy, StringComparison.OrdinalIgnoreCase));
            if (point == null)
            {
                string modality = GuessModality(energy);
                double? depth = GetDefaultTg51DepthCm(energy);
                point = new Tg51Point
                {
                    PointId = $"{tg51Run.Points.Count + 1:00}_{energy}_Reference",
                    Energy = energy,
                    Modality = modality,
                    DepthLabel = modality == "Electron" ? "dref" : "Reference",
                    ClinicalDepthCm = depth,
                    Status = "Active"
                };
                tg51Run.Points.Add(point);
            }

            tg51CurrentPointIndex = tg51Run.Points.IndexOf(point);
            return point;
        }

        private int GetActiveTg51Bias(Tg51Point point)
        {
            if (tg51ActiveBiasByPoint.TryGetValue(point, out int activeBias))
            {
                return activeBias;
            }

            foreach (int bias in tg51BiasSequence)
            {
                if (GetTg51BiasReadingCount(point, bias) < Tg51ReadingsPerBias)
                {
                    tg51ActiveBiasByPoint[point] = bias;
                    return bias;
                }
            }

            tg51ActiveBiasByPoint[point] = tg51BiasSequence[0];
            return tg51BiasSequence[0];
        }

        private int GetTg51BiasReadingCount(Tg51Point point, int bias)
        {
            return point.Readings.Count(reading => reading.BiasVoltage == bias);
        }

        private int GetNextTg51Bias(int currentBias)
        {
            int index = Array.IndexOf(tg51BiasSequence, currentBias);
            return index >= 0 && index < tg51BiasSequence.Length - 1
                ? tg51BiasSequence[index + 1]
                : currentBias;
        }

        private bool HasTg51MinimumReadings(Tg51Point point)
        {
            return tg51BiasSequence.All(bias => GetTg51BiasReadingCount(point, bias) >= Tg51ReadingsPerBias);
        }

        private async Task SetActiveTg51BiasAsync(Tg51Point point, int bias, bool transition)
        {
            tg51ActiveBiasByPoint[point] = bias;
            expectedBiasVoltage = bias;
            RefreshTg51View();

            if (!doseXConnected || wsClient?.State != WebSocketState.Open)
            {
                return;
            }

            if (transition)
            {
                await TransitionTg51BiasAsync(bias);
            }
            else
            {
                await SetBiasVoltageAsync(bias);
                await SetMeasurementConfigurationAsync();
                await StartMeasurementAsync();
            }
        }

        private void RefreshTg51View()
        {
            Tg51Bias300List.Items.Clear();
            Tg51BiasNeg300List.Items.Clear();
            Tg51Bias150List.Items.Clear();

            var point = GetOrCreateSelectedTg51Point();
            if (point == null)
            {
                Tg51DepthValue.Text = "Depth: --";
                Tg51DepthCmText.Text = string.Empty;
                RefreshTg51Preview(null);
                return;
            }

            if (!string.Equals(tg51DepthEnergy, point.Energy, StringComparison.OrdinalIgnoreCase))
            {
                SetTg51DepthTextToDefaultForEnergy(point.Energy);
            }

            Tg51DepthValue.Text = TryGetTg51DepthCmFromText(out double displayDepthCm)
                ? $"Depth: {displayDepthCm:0.##} cm"
                : "Depth: --";

            foreach (var reading in point.Readings)
            {
                string item = $"{reading.RepeatNumber}: {reading.ChargeNc:0.0000}";
                if (reading.BiasVoltage == 300)
                {
                    Tg51Bias300List.Items.Add(item);
                }
                else if (reading.BiasVoltage == -300)
                {
                    Tg51BiasNeg300List.Items.Add(item);
                }
                else if (reading.BiasVoltage == 150)
                {
                    Tg51Bias150List.Items.Add(item);
                }
            }

            Color activeColor = Color.FromArgb(255, 255, 180);
            Tg51Bias300List.BackColor = SystemColors.Window;
            Tg51BiasNeg300List.BackColor = SystemColors.Window;
            Tg51Bias150List.BackColor = SystemColors.Window;

            if (!HasTg51MinimumReadings(point))
            {
                int activeBias = GetActiveTg51Bias(point);
                if (activeBias == 300) Tg51Bias300List.BackColor = activeColor;
                if (activeBias == -300) Tg51BiasNeg300List.BackColor = activeColor;
                if (activeBias == 150) Tg51Bias150List.BackColor = activeColor;
                int repeat = GetTg51BiasReadingCount(point, activeBias) + 1;
                Tg51Status.Text = $"Beam on: {point.Energy}, {FormatTg51BiasLabel(activeBias)}, reading {repeat}.";
            }
            else
            {
                int activeBias = GetActiveTg51Bias(point);
                if (activeBias == 300) Tg51Bias300List.BackColor = activeColor;
                if (activeBias == -300) Tg51BiasNeg300List.BackColor = activeColor;
                if (activeBias == 150) Tg51Bias150List.BackColor = activeColor;
                int repeat = GetTg51BiasReadingCount(point, activeBias) + 1;
                Tg51Status.Text = $"{point.Energy} minimum set complete. Active: {FormatTg51BiasLabel(activeBias)}, reading {repeat}.";
            }

            RefreshTg51Grids();
            RefreshTg51Preview(point);
        }

        private void RefreshTg51Preview(Tg51Point point)
        {
            if (Tg51Preview == null)
            {
                return;
            }

            if (point == null)
            {
                Tg51Preview.Text = "Preview only - not written to XML\r\nChoose an energy to preview TG-51 checks.";
                return;
            }

            var high = GetTg51Charges(point, 300);
            var low = GetTg51Charges(point, 150);
            var opposite = GetTg51Charges(point, -300);
            Tg51TemplatePreviewContext context = LoadTg51TemplatePreviewContext(point.Energy, point.Modality);

            var lines = new List<string>
            {
                "Preview only - not written to XML"
            };

            if (high.Count == 0)
            {
                lines.Add("+300 V: waiting for readings");
            }
            else
            {
                double meanHigh = high.Average();
                lines.Add($"+300 V mean {meanHigh:0.####} nC ({high.Count} reading{(high.Count == 1 ? string.Empty : "s")})");

                if (high.Count >= 2)
                {
                    double absMean = Math.Abs(meanHigh);
                    double rangePercent = absMean > 0 ? (high.Max() - high.Min()) / absMean * 100.0 : 0.0;
                    lines.Add($"Repeatability range {rangePercent:0.###}%{(rangePercent > 0.2 ? " - check spread" : string.Empty)}");
                }
            }

            double? measuredPpol = CalculateTg51Ppol(high, opposite);
            AddFactorPreviewLine(lines, "Ppol", measuredPpol, context.ExpectedPpol, 0.995, 1.005, 0.990, 1.010,
                opposite.Count == 0 ? "waiting for -300 V" : null);

            double? measuredPion = CalculateTg51Pion(high, low);
            AddFactorPreviewLine(lines, "Pion", measuredPion, context.ExpectedPion, 1.000, 1.020, 0.995, 1.030,
                low.Count == 0 ? "waiting for 50%/+150 V" : null);

            double? outputPpol = SelectPreviewCorrectionFactor(measuredPpol, context.ExpectedPpol, context.UseAssignedPionPpol);
            double? outputPion = SelectPreviewCorrectionFactor(measuredPion, context.ExpectedPion, context.UseAssignedPionPpol);
            if ((outputPpol != measuredPpol || outputPion != measuredPion) && (outputPpol.HasValue || outputPion.HasValue))
            {
                lines.Add($"Output corrections use Ppol {FormatNullableFactor(outputPpol)}, Pion {FormatNullableFactor(outputPion)}");
            }

            double? roughOutput = CalculateRoughTg51OutputPerMu(high, outputPpol, outputPion, context);
            if (roughOutput.HasValue)
            {
                lines.Add($"{context.SetupLabel}: measured point {roughOutput.Value:0.###} cGy/MU{context.QualityCorrectionLabel}");
                double? referenceOutput = Tg51PreviewMath.CalculateReferenceOutputPerMu(roughOutput, context.ClinicalPdd);
                if (referenceOutput.HasValue)
                {
                    string factorLabel = context.IsTmr.GetValueOrDefault(false) ? "TMR" : "PDD";
                    lines.Add($"{context.ReferenceOutputLabel} {referenceOutput.Value:0.###} cGy/MU using {factorLabel} {context.ClinicalPdd.Value:0.####}");
                }
                else if (context.HasTemplate)
                {
                    lines.Add("Reference output unavailable - missing ClinicalPDD/TMR factor");
                }
            }
            else
            {
                lines.Add(context.HasTemplate
                    ? "Rough output unavailable - missing template factor or +300 V readings"
                    : "Rough output unavailable - import TG-51 Session XML");
            }

            if (context.MeasuredPdd10.HasValue)
            {
                lines.Add($"Template PDD10 {context.MeasuredPdd10.Value:0.###}%");
            }
            else if (context.R50.HasValue)
            {
                lines.Add($"Template R50 {context.R50.Value:0.###} cm");
            }
            else if (context.HasTemplate)
            {
                lines.Add("Template loaded; no PDD10/R50 found for this energy");
            }

            Tg51Preview.Text = string.Join("\r\n", lines);
        }

        private List<double> GetTg51Charges(Tg51Point point, int bias)
        {
            return point.Readings
                .Where(reading => reading.BiasVoltage == bias)
                .Select(reading => reading.ChargeNc)
                .ToList();
        }

        private double? CalculateTg51Ppol(List<double> high, List<double> opposite)
        {
            return Tg51PreviewMath.CalculatePpol(high, opposite);
        }

        private double? CalculateTg51Pion(List<double> high, List<double> low)
        {
            return Tg51PreviewMath.CalculatePion(high, low);
        }

        private void AddFactorPreviewLine(
            List<string> lines,
            string label,
            double? measured,
            double? expected,
            double normalLow,
            double normalHigh,
            double warningLow,
            double warningHigh,
            string waitingMessage)
        {
            if (!measured.HasValue)
            {
                lines.Add($"{label}: {waitingMessage ?? "unavailable"}");
                return;
            }

            string message = $"{label} {measured.Value:0.####}";
            if (expected.HasValue)
            {
                double delta = measured.Value - expected.Value;
                message += $" vs template {expected.Value:0.####} ({delta:+0.####;-0.####;0})";
            }
            else
            {
                message += " vs sanity range";
            }

            if (measured.Value < warningLow || measured.Value > warningHigh)
            {
                message += " - CHECK";
            }
            else if (measured.Value < normalLow || measured.Value > normalHigh)
            {
                message += " - watch";
            }

            lines.Add(message);
        }

        private static double? SelectPreviewCorrectionFactor(double? measured, double? expected, bool useAssigned)
        {
            if (measured.HasValue)
            {
                return measured;
            }

            if (useAssigned && expected.HasValue)
            {
                return expected;
            }

            return expected;
        }

        private static string FormatNullableFactor(double? value)
        {
            return value.HasValue ? value.Value.ToString("0.####", CultureInfo.InvariantCulture) : "n/a";
        }

        private double? CalculateRoughTg51OutputPerMu(
            List<double> high,
            double? ppol,
            double? pion,
            Tg51TemplatePreviewContext context)
        {
            return Tg51PreviewMath.CalculateRoughOutputPerMu(
                high,
                ppol,
                pion,
                context.DetectorCalibrationFactor,
                context.TemperatureC,
                context.PressureMmHg,
                context.DeliveredMu,
                context.Prp,
                context.DoseToTissueCorrection,
                context.QualityCorrectionFactor);
        }

        private Tg51TemplatePreviewContext LoadTg51TemplatePreviewContext(string energy, string modality)
        {
            var context = new Tg51TemplatePreviewContext();
            string templatePath = tg51Run?.TemplateXmlPath ?? string.Empty;
            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                return context;
            }

            try
            {
                XDocument document = XDocument.Load(templatePath);
                XElement root = document.Root;
                if (root == null)
                {
                    return context;
                }

                context.HasTemplate = true;
                context.DetectorCalibrationFactor = ReadXmlDouble(root.Descendants().FirstOrDefault(element => element.Name.LocalName == "Detector")?
                    .Elements().FirstOrDefault(element => element.Name.LocalName == "CalibrationFactor"));

                XElement property = root.Descendants()
                    .FirstOrDefault(element =>
                        element.Name.LocalName == "P_Property" &&
                        EnergyMatches(GetChildValue(element, "NRG"), energy));
                if (property != null)
                {
                    context.ExpectedPion = ReadXmlDouble(GetChild(property, "Pion"));
                    context.ExpectedPpol = ReadXmlDouble(GetChild(property, "Ppol"));
                }

                bool isElectron = string.Equals(modality, "Electron", StringComparison.OrdinalIgnoreCase) ||
                                  energy.Contains("e", StringComparison.OrdinalIgnoreCase);
                context.IsElectron = isElectron;
                XElement tg51Factors = root.Descendants().FirstOrDefault(element => element.Name.LocalName == "TG51Factors");
                context.PhotonKqA = ReadXmlDouble(tg51Factors?.Elements().FirstOrDefault(element => element.Name.LocalName == "A"));
                context.PhotonKqB = ReadXmlDouble(tg51Factors?.Elements().FirstOrDefault(element => element.Name.LocalName == "B"));
                context.PhotonKqC = ReadXmlDouble(tg51Factors?.Elements().FirstOrDefault(element => element.Name.LocalName == "C"));
                context.ElectronKecal = ReadXmlDouble(tg51Factors?.Elements().FirstOrDefault(element => element.Name.LocalName == "Kecal"));
                string calibrationElementName = isElectron ? "TG51_Electron_Full" : "TG51_Photon_Full";
                XElement calibration = root.Descendants()
                    .FirstOrDefault(element =>
                        element.Name.LocalName == calibrationElementName &&
                        EnergyMatches(GetChildValue(element, "NRG"), energy));

                if (calibration != null)
                {
                    context.TemperatureC = ReadXmlDouble(GetChild(calibration, "Temperature"));
                    context.PressureMmHg = ReadXmlDouble(GetChild(calibration, "Pressure"));
                    context.DeliveredMu = ReadXmlDouble(GetChild(calibration, "DeliveredMU"));
                    context.UseAssignedPionPpol = ReadXmlBool(GetChild(calibration, "UseAssignedPionPpol")).GetValueOrDefault(false);
                    context.IsTmr = ReadXmlBool(GetChild(calibration, "IsTMR"));
                    context.SsdCm = ReadXmlDouble(GetChild(calibration, "SSD"));
                    context.DepthCm = ReadXmlDouble(GetChild(calibration, isElectron ? "MeasurementDepth" : "Depth"));
                    context.MeasuredPdd10 = ReadXmlDouble(GetChild(calibration, "MeasuredPDD10"));
                    context.ClinicalPdd = ReadXmlDouble(GetChild(calibration, "ClinicalPDD"));
                    context.R50 = ReadXmlDouble(GetChild(calibration, "R50"));
                    context.Prp = ReadXmlDouble(GetChild(calibration, "Prp"));
                    context.DoseToTissueCorrection = ReadXmlDouble(GetChild(calibration, "DoseToTissueCorrection"));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("TG-51 preview template read failed: " + ex.Message);
            }

            return context;
        }

        private static bool EnergyMatches(string left, string right)
        {
            return string.Equals(NormalizeEnergy(left), NormalizeEnergy(right), StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeEnergy(string value)
        {
            return Regex.Replace(value ?? string.Empty, @"[\s\-_]", string.Empty).ToUpperInvariant();
        }

        private static double? ReadXmlDouble(XElement element)
        {
            if (element == null)
            {
                return null;
            }

            string text = element.Value?.Trim() ?? string.Empty;
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                   double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value)
                ? value
                : null;
        }

        private static bool? ReadXmlBool(XElement element)
        {
            if (element == null)
            {
                return null;
            }

            string text = element.Value?.Trim() ?? string.Empty;
            return bool.TryParse(text, out bool value) ? value : null;
        }

        private static XElement GetChild(XElement parent, string localName)
        {
            return parent?.Elements().FirstOrDefault(element => string.Equals(element.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetChildValue(XElement parent, string localName)
        {
            return GetChild(parent, localName)?.Value?.Trim() ?? string.Empty;
        }

        private class Tg51TemplatePreviewContext
        {
            public bool HasTemplate { get; set; }
            public double? DetectorCalibrationFactor { get; set; }
            public double? ExpectedPion { get; set; }
            public double? ExpectedPpol { get; set; }
            public double? TemperatureC { get; set; }
            public double? PressureMmHg { get; set; }
            public double? DeliveredMu { get; set; }
            public bool UseAssignedPionPpol { get; set; }
            public bool? IsTmr { get; set; }
            public double? SsdCm { get; set; }
            public double? DepthCm { get; set; }
            public double? MeasuredPdd10 { get; set; }
            public double? ClinicalPdd { get; set; }
            public double? R50 { get; set; }
            public double? Prp { get; set; }
            public double? DoseToTissueCorrection { get; set; }
            public bool IsElectron { get; set; }
            public double? ElectronKecal { get; set; }
            public double? PhotonKqA { get; set; }
            public double? PhotonKqB { get; set; }
            public double? PhotonKqC { get; set; }

            public double? QualityCorrectionFactor => IsElectron
                ? ElectronKecal
                : Tg51PreviewMath.CalculatePhotonKq(MeasuredPdd10, PhotonKqA, PhotonKqB, PhotonKqC);

            public string QualityCorrectionLabel
            {
                get
                {
                    if (IsElectron && ElectronKecal.HasValue)
                    {
                        return $" with Kecal {ElectronKecal.Value:0.####}";
                    }

                    double? photonKq = Tg51PreviewMath.CalculatePhotonKq(MeasuredPdd10, PhotonKqA, PhotonKqB, PhotonKqC);
                    if (!IsElectron && photonKq.HasValue)
                    {
                        return $" with kQ {photonKq.Value:0.####}";
                    }

                    return IsElectron ? " (Kecal unavailable)" : " (photon kQ unavailable)";
                }
            }

            public string SetupLabel
            {
                get
                {
                    string geometry = IsTmr.GetValueOrDefault(false) ? "SAD/TMR" : "SSD/PDD";
                    string ssd = SsdCm.HasValue ? $"{SsdCm.Value:0.###} cm SSD" : "SSD unknown";
                    string depth = DepthCm.HasValue ? $"d={DepthCm.Value:0.###} cm" : "depth unknown";
                    return $"{geometry} {ssd}, {depth}";
                }
            }

            public string ReferenceOutputLabel => IsTmr.GetValueOrDefault(false)
                ? "SAD/TMR reference output"
                : "SSD/PDD reference output";
        }

        private async Task SetTg51BiasForCurrentBucketAsync(Tg51Point point)
        {
            if (!doseXConnected || wsClient?.State != WebSocketState.Open)
            {
                return;
            }

            int bias = GetActiveTg51Bias(point);
            expectedBiasVoltage = bias;
            await SetBiasVoltageAsync(bias);
            await SetMeasurementConfigurationAsync();
            await StartMeasurementAsync();
        }

        private string FormatTg51BiasLabel(int bias)
        {
            return bias == 150 ? "50% (+150 V)" : $"{bias:+#;-#;0} V";
        }

        private async Task SelectTg51BiasBucketAsync(int bias)
        {
            if (tg51BiasSelectionInProgress)
            {
                return;
            }

            var point = GetOrCreateSelectedTg51Point();
            if (point == null)
            {
                return;
            }

            try
            {
                tg51BiasSelectionInProgress = true;
                int previousBias = GetActiveTg51Bias(point);
                bool changingBias = previousBias != bias;
                point.Status = HasTg51MinimumReadings(point) ? "Complete" : "Active";
                tg51RunActive = true;
                SaveTg51Run();

                if (!changingBias)
                {
                    tg51ActiveBiasByPoint[point] = bias;
                    return;
                }

                if (changingBias)
                {
                    await RunTg51WaitModalAsync(
                        $"Wait while bias stabilizes at {FormatTg51BiasLabel(bias)}.",
                        async () => await SetActiveTg51BiasAsync(point, bias, transition: false),
                        Tg51ManualBiasSettleDelayMs);
                }
            }
            finally
            {
                tg51BiasSelectionInProgress = false;
            }
        }

        private async void Tg51Bias300Bucket_Selected(object sender, EventArgs e)
        {
            await SelectTg51BiasBucketAsync(300);
        }

        private async void Tg51BiasNeg300Bucket_Selected(object sender, EventArgs e)
        {
            await SelectTg51BiasBucketAsync(-300);
        }

        private async void Tg51Bias150Bucket_Selected(object sender, EventArgs e)
        {
            await SelectTg51BiasBucketAsync(150);
        }

        private bool TryGetSelectedTg51Reading(out Tg51Point point, out int bias, out int selectedIndex, out Tg51Reading reading)
        {
            reading = null;
            if (!TryGetSelectedTg51Bucket(out point, out bias, out selectedIndex, out System.Windows.Forms.ListBox _))
            {
                return false;
            }

            int selectedBias = bias;
            var bucketReadings = point.Readings.Where(item => item.BiasVoltage == selectedBias).ToList();
            if (selectedIndex < 0 || selectedIndex >= bucketReadings.Count)
            {
                return false;
            }

            reading = bucketReadings[selectedIndex];
            return true;
        }

        private bool TryGetSelectedTg51Bucket(out Tg51Point point, out int bias, out int selectedIndex, out System.Windows.Forms.ListBox selectedList)
        {
            point = GetOrCreateSelectedTg51Point();
            bias = 0;
            selectedIndex = -1;
            selectedList = null;

            if (point == null)
            {
                return false;
            }

            if (Tg51Bias300List.SelectedIndex >= 0)
            {
                bias = 300;
                selectedIndex = Tg51Bias300List.SelectedIndex;
                selectedList = Tg51Bias300List;
                return true;
            }

            if (Tg51BiasNeg300List.SelectedIndex >= 0)
            {
                bias = -300;
                selectedIndex = Tg51BiasNeg300List.SelectedIndex;
                selectedList = Tg51BiasNeg300List;
                return true;
            }

            if (Tg51Bias150List.SelectedIndex >= 0)
            {
                bias = 150;
                selectedIndex = Tg51Bias150List.SelectedIndex;
                selectedList = Tg51Bias150List;
                return true;
            }

            return false;
        }

        private Tg51Reading CreateTg51Reading(int bias, double charge)
        {
            return new Tg51Reading
            {
                BiasVoltage = bias,
                RepeatNumber = 0,
                ChargeNc = charge,
                RecordedAt = DateTime.Now,
                DoseXHighVoltageEnabled = globalHighVoltageEnabled,
                DoseXBiasVoltage = globalBiasVoltage,
                BridgeSnapshotJson = GetCurrentTg51Point()?.LastBridgeSnapshotJson ?? string.Empty
            };
        }

        private double GetCurrentTg51CorrectionCharge()
        {
            if (double.TryParse(Measurement.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out double charge) ||
                double.TryParse(Measurement.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out charge))
            {
                return charge;
            }

            return lastKnownCharge;
        }

        private void RenumberTg51BiasReadings(Tg51Point point, int bias)
        {
            int repeat = 1;
            foreach (var reading in point.Readings.Where(item => item.BiasVoltage == bias))
            {
                reading.RepeatNumber = repeat++;
            }
        }

        private void RestoreTg51Selection(int bias, int selectedIndex)
        {
            System.Windows.Forms.ListBox list = bias == 300
                ? Tg51Bias300List
                : bias == -300
                    ? Tg51BiasNeg300List
                    : Tg51Bias150List;

            if (list.Items.Count == 0)
            {
                return;
            }

            list.SelectedIndex = Math.Min(Math.Max(selectedIndex, 0), list.Items.Count - 1);
        }

        private void ClearPendingTg51Overwrite()
        {
            pendingTg51OverwritePoint = null;
            pendingTg51OverwriteReading = null;
            pendingTg51OverwriteBias = 0;
            pendingTg51OverwriteIndex = -1;
        }

        private async void Tg51EnergyCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tg51ChangingSelection || tg51EnergyChangeInProgress)
            {
                return;
            }

            try
            {
                tg51EnergyChangeInProgress = true;

                string selectedEnergy = GetSelectedTg51Energy();
                if (string.IsNullOrWhiteSpace(selectedEnergy))
                {
                    return;
                }

                Tg51Point currentPoint = GetCurrentTg51Point();
                if (currentPoint != null &&
                    string.Equals(currentPoint.Energy, selectedEnergy, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(tg51LastHandledEnergy, selectedEnergy, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                tg51LastHandledEnergy = selectedEnergy;
                ClearPendingTg51Overwrite();
                var point = GetOrCreateSelectedTg51Point();
                if (point == null)
                {
                    return;
                }

                if (point.Readings.Count == 0)
                {
                    tg51ActiveBiasByPoint[point] = 300;
                }

                point.Status = HasTg51MinimumReadings(point) ? "Complete" : "Active";
                tg51RunActive = point.Status != "Complete";
                RefreshTg51View();
                SaveTg51Run();

                if (Tg51MoveDepthAuto.Checked)
                {
                    await RunTg51WaitModalAsync(
                        "Wait while chamber move is sent.",
                        async () => await MoveSelectedTg51DepthAsync(),
                        3000);
                    await RefreshSelectedTg51TankStateAsync();
                }
            }
            finally
            {
                tg51EnergyChangeInProgress = false;
            }
        }

        private void Tg51ClearEnergy_Click(object sender, EventArgs e)
        {
            var point = GetOrCreateSelectedTg51Point();
            if (point == null)
            {
                return;
            }

            point.Readings.Clear();
            point.Status = "Active";
            tg51RunActive = true;
            tg51ActiveBiasByPoint[point] = 300;
            ClearPendingTg51Overwrite();
            SaveTg51Run();
            RefreshTg51View();
        }

        private void Tg51DeleteReading_Click(object sender, EventArgs e)
        {
            if (!TryGetSelectedTg51Reading(out Tg51Point point, out int bias, out int selectedIndex, out Tg51Reading reading))
            {
                Tg51Status.Text = "Select a TG-51 reading to delete.";
                return;
            }

            point.Readings.Remove(reading);
            if (pendingTg51OverwriteReading == reading)
            {
                ClearPendingTg51Overwrite();
            }

            RenumberTg51BiasReadings(point, bias);
            point.Status = HasTg51MinimumReadings(point) ? "Complete" : "Active";
            tg51RunActive = true;
            SaveTg51Run();
            RefreshTg51View();
            RestoreTg51Selection(bias, selectedIndex);
            Tg51Status.Text = "TG-51 reading deleted.";
        }

        private void Tg51InsertReading_Click(object sender, EventArgs e)
        {
            if (!TryGetSelectedTg51Bucket(out Tg51Point point, out int bias, out int selectedIndex, out _))
            {
                Tg51Status.Text = "Select a TG-51 bucket position before inserting.";
                return;
            }

            double charge = GetCurrentTg51CorrectionCharge();
            var bucketReadings = point.Readings.Where(reading => reading.BiasVoltage == bias).ToList();
            int globalIndex = point.Readings.Count;
            if (selectedIndex >= 0 && selectedIndex < bucketReadings.Count)
            {
                globalIndex = point.Readings.IndexOf(bucketReadings[selectedIndex]);
            }

            point.Readings.Insert(globalIndex, CreateTg51Reading(bias, charge));
            RenumberTg51BiasReadings(point, bias);
            point.Status = HasTg51MinimumReadings(point) ? "Complete" : "Active";
            tg51RunActive = true;
            SaveTg51Run();
            RefreshTg51View();
            RestoreTg51Selection(bias, Math.Max(0, selectedIndex));
            Tg51Status.Text = "TG-51 reading inserted.";
        }

        private void Tg51OverwriteReading_Click(object sender, EventArgs e)
        {
            if (!TryGetSelectedTg51Reading(out Tg51Point point, out int bias, out int selectedIndex, out Tg51Reading reading))
            {
                Tg51Status.Text = "Select a TG-51 reading to overwrite.";
                return;
            }

            pendingTg51OverwritePoint = point;
            pendingTg51OverwriteReading = reading;
            pendingTg51OverwriteBias = bias;
            pendingTg51OverwriteIndex = selectedIndex;
            tg51ActiveBiasByPoint[point] = bias;
            RestoreTg51Selection(bias, selectedIndex);
            Tg51Status.Text = $"Next TG-51 measurement will overwrite {FormatTg51BiasLabel(bias)} #{selectedIndex + 1}.";
        }

        private async void Tg51ConnectTank_Click(object sender, EventArgs e)
        {
            await ConnectTg51TankBridgeAsync();
        }

        private async void Tg51MoveDepth_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(GetSelectedTg51Energy()))
            {
                Tg51Status.Text = "Choose an energy before moving depth.";
                return;
            }

            if (!TryGetTg51DepthCmFromText(out double depthCm))
            {
                Tg51Status.Text = "Enter a TG-51 depth in cm before moving depth.";
                return;
            }

            if (depthCm < 0 || depthCm > 50)
            {
                Tg51Status.Text = "TG-51 depth must be between 0 and 50 cm.";
                return;
            }

            await RunTg51WaitModalAsync(
                "Wait while chamber move is sent.",
                async () => await MoveSelectedTg51DepthAsync(),
                3000);
            await RefreshSelectedTg51TankStateAsync();
        }

        private async void Tg51DisconnectTank_Click(object sender, EventArgs e)
        {
            await DisconnectTg51TankBridgeAsync();
        }

        private async void Tg51SetTarget_Click(object sender, EventArgs e)
        {
            await SetSelectedTg51TargetFromCurrentPositionAsync();
        }

        private async void Tg51TestStart_Click(object sender, EventArgs e)
        {
            await StartTg51TestMeasurementAsync();
        }

        private async void Tg51TestStop_Click(object sender, EventArgs e)
        {
            await StopTg51TestMeasurementAsync();
        }

        private async void Tg51Background_Click(object sender, EventArgs e)
        {
            await RunBackgroundMeasurementAsync();
        }

        private async Task StartTg51TestMeasurementAsync()
        {
            if (!doseXConnected || wsClient?.State != WebSocketState.Open)
            {
                MessageBox.Show("Connect Electrometer");
                return;
            }

            var point = GetOrCreateSelectedTg51Point();
            if (point == null)
            {
                Tg51Status.Text = "Choose an energy before starting a TG-51 test measurement.";
                return;
            }

            bool hasControl = false;
            try
            {
                Tg51TestStart.Enabled = false;
                int bias = GetActiveTg51Bias(point);
                expectedBiasVoltage = bias;

                bool isControlAvailable = await RequestControlTokenAsync();
                if (!isControlAvailable)
                {
                    MessageBox.Show("Please free control from the electrometer by pressing the 'cloud' at the top.");
                    return;
                }
                hasControl = true;

                await SendBiasVoltageAsync(bias);
                await SetMeasurementConfigurationAsync("manual");
                await StartMeasurementAsync();
                Tg51Status.Text = $"Test collecting: {point.Energy}, {FormatTg51BiasLabel(GetActiveTg51Bias(point))}.";
            }
            finally
            {
                if (hasControl)
                {
                    await ReleaseControlAsync();
                }
                Tg51TestStart.Enabled = true;
            }
        }

        private async Task StopTg51TestMeasurementAsync()
        {
            if (!doseXConnected || wsClient?.State != WebSocketState.Open)
            {
                MessageBox.Show("Connect Electrometer");
                return;
            }

            bool isControlAvailable = await RequestControlTokenAsync();
            if (!isControlAvailable)
            {
                MessageBox.Show("Please free control from the electrometer by pressing the 'cloud' at the top.");
                return;
            }

            try
            {
                Tg51TestStop.Enabled = false;
                await StopMeasurementAsync();
                bool suppressAutomaticCommit = hasSeenTrue;
                bool recorded = AddOrUpdateCurrentMeasurement();
                if (!recorded && IsTg51Mode())
                {
                    recorded = RecordTg51Charge(lastKnownCharge.ToString("0.0000", CultureInfo.InvariantCulture));
                }

                if (recorded)
                {
                    if (suppressAutomaticCommit)
                    {
                        suppressNextMeasurementCommit = true;
                    }

                    Tg51Status.Text = "Test reading recorded.";
                }
                else
                {
                    Tg51Status.Text = "Test stop requested. Waiting for DoseX reading...";
                }
            }
            finally
            {
                await ReleaseControlAsync();
                Tg51TestStop.Enabled = true;
            }
        }

        private void RefreshTg51Grids()
        {
            RefreshTg51PointGridStatuses();
            RefreshTg51ReadingGrid();
        }

        private void RefreshTg51PointGridStatuses()
        {
            if (tg51Run == null)
            {
                return;
            }

            foreach (DataGridViewRow row in Tg51PointGrid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                string energy = row.Cells["Energy"].Value?.ToString() ?? string.Empty;
                string depthLabel = row.Cells["DepthLabel"].Value?.ToString() ?? string.Empty;
                var point = tg51Run.Points.FirstOrDefault(p => p.Energy == energy && p.DepthLabel == depthLabel);
                if (point != null)
                {
                    row.Cells["Status"].Value = $"{point.Status} ({point.Readings.Count}/9)";
                }
            }
        }

        private void RefreshTg51ReadingGrid()
        {
            Tg51ReadingGrid.Rows.Clear();
            if (tg51Run == null)
            {
                return;
            }

            foreach (var point in tg51Run.Points)
            {
                foreach (var reading in point.Readings)
                {
                    Tg51ReadingGrid.Rows.Add(
                        $"{point.Energy} {point.DepthLabel}",
                        reading.BiasVoltage,
                        reading.RepeatNumber,
                        reading.ChargeNc.ToString("0.0000", CultureInfo.CurrentCulture),
                        reading.RecordedAt.ToString("yyyy-MM-dd HH:mm:ss"));
                }
            }
        }

        private void SaveTg51Run()
        {
            if (tg51Run == null)
            {
                return;
            }

            tg51Logger ??= new Tg51RunLogger(machine.Path);
            tg51Run.BridgeEnabled = false;
            tg51Run.BridgeUrl = GetTg51BridgeBaseUrl();
            tg51Logger.Save(tg51Run);
            Tg51XmlPath.Text = tg51Logger.FilePath;
        }

        private void Tg51ImportXml_Click(object sender, EventArgs e)
        {
            using OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Import TG-51 Session XML",
                Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            try
            {
                XDocument document = XDocument.Load(dialog.FileName);
                if (!string.Equals(document.Root?.Name.LocalName, "Session", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("That file does not look like a TG-51 Session XML.");
                    return;
                }

                tg51Run ??= new Tg51Run
                {
                    RunId = $"TG51_{DateTime.Now:yyyyMMdd_HHmmss}",
                    MachineName = machine.Name ?? string.Empty,
                    MachineType = machine.Type ?? string.Empty,
                    MachinePath = machine.Path ?? string.Empty,
                    StartedAt = DateTime.Now
                };

                tg51Run.TemplateXmlPath = dialog.FileName;
                SaveTg51Run();
                Tg51Status.Text = "TG-51 Session XML template loaded.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to import TG-51 XML: " + ex.Message);
            }
        }

        private void LoadTg51RunState()
        {
            try
            {
                tg51Logger = new Tg51RunLogger(machine.Path);
                Tg51XmlPath.Text = tg51Logger.FilePath;

                Tg51Run loadedRun = tg51Logger.Load();
                if (loadedRun == null)
                {
                    return;
                }

                tg51Run = loadedRun;
                foreach (var point in tg51Run.Points)
                {
                    foreach (int bias in tg51BiasSequence)
                    {
                        RenumberTg51BiasReadings(point, bias);
                    }
                }

                tg51CurrentPointIndex = -1;
                tg51RunActive = tg51Run.Points.Any(point => point.Status != "Complete");
                RefreshTg51Grids();
                Tg51Status.Text = "TG-51 state loaded.";
            }
            catch (Exception ex)
            {
                Tg51Status.Text = "TG-51 state load failed.";
                Console.WriteLine("Failed to load TG-51 state: " + ex.Message);
            }
        }

        private void Tg51SaveXml_Click(object sender, EventArgs e)
        {
            SaveTg51Run();
            MessageBox.Show(string.IsNullOrEmpty(Tg51XmlPath.Text) ? "No TG-51 run has been created." : "TG-51 XML saved.");
        }

        private void Tg51Pause_Click(object sender, EventArgs e)
        {
            tg51RunActive = false;
            if (GetCurrentTg51Point() is { } point && point.Status == "Active")
            {
                point.Status = "Paused";
            }
            SaveTg51Run();
            RefreshTg51Grids();
            Tg51Status.Text = "TG-51 paused.";
        }

        private async void Tg51SkipPoint_Click(object sender, EventArgs e)
        {
            if (GetCurrentTg51Point() is { } point)
            {
                point.Status = "Skipped";
            }

            tg51RunActive = false;
            SaveTg51Run();
            RefreshTg51Grids();

            if (tg51Run != null)
            {
                int nextIndex = tg51Run.Points.FindIndex(tg51CurrentPointIndex + 1, p => p.Status != "Complete" && p.Status != "Skipped");
                await PrepareTg51PointAsync(nextIndex);
            }
        }

        private void Tg51BuildPoints_Click(object sender, EventArgs e)
        {
            Tg51PointGrid.Rows.Clear();
            foreach (object item in Tg51EnergyList.CheckedItems)
            {
                string energy = item.ToString();
                string modality = GuessModality(energy);
                double? defaultDepth = GetDefaultTg51DepthCm(energy);
                Tg51PointGrid.Rows.Add(true, energy, modality, modality == "Electron" ? "dref" : "Reference", defaultDepth, 0, 0, string.Empty, "Pending");
            }
            Tg51Status.Text = "Review depths and raw tank targets, then create the run.";
        }

        private string GuessModality(string energy)
        {
            return energy.Contains("e", StringComparison.OrdinalIgnoreCase) ? "Electron" : "Photon";
        }

        private double? GetDefaultTg51DepthCm(string energy)
        {
            string normalized = energy.ToLowerInvariant();
            if (normalized.Contains("x"))
            {
                return 10.0;
            }

            return ExtractNumber(energy) switch
            {
                <= 6 => 1.3,
                <= 9 => 2.1,
                <= 12 => 2.8,
                <= 15 => 3.5,
                <= 18 => 4.3,
                _ => 5.0
            };
        }

        private bool RepairTg51PointDepthIfNeeded(Tg51Point point)
        {
            if (point == null || !string.Equals(GuessModality(point.Energy), "Electron", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            double? defaultDepth = GetDefaultTg51DepthCm(point.Energy);
            if (!defaultDepth.HasValue)
            {
                return false;
            }

            if (!point.ClinicalDepthCm.HasValue)
            {
                point.ClinicalDepthCm = defaultDepth;
                point.DepthLabel = "dref";
                return true;
            }

            bool looksLikePhotonDefault =
                Math.Abs(point.ClinicalDepthCm.Value - 10.0) < 0.001 &&
                Math.Abs(defaultDepth.Value - 10.0) > 0.001;

            bool hasPhotonDepthLabel =
                string.IsNullOrWhiteSpace(point.DepthLabel) ||
                string.Equals(point.DepthLabel, "Reference", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(point.DepthLabel, "10 cm", StringComparison.OrdinalIgnoreCase);

            if (looksLikePhotonDefault && hasPhotonDepthLabel)
            {
                point.ClinicalDepthCm = defaultDepth;
                point.DepthLabel = "dref";
                point.RawCrosslineMm = null;
                point.RawInlineMm = null;
                point.RawDepthMm = null;
                return true;
            }

            return false;
        }

        private async void Tg51CheckBridge_Click(object sender, EventArgs e)
        {
            string baseUrl = Tg51BridgeUrl.Text.Trim().TrimEnd('/');
            if (string.IsNullOrEmpty(baseUrl))
            {
                MessageBox.Show("Enter a tank bridge URL first.");
                return;
            }

            try
            {
                string stateJson = await tg51TankClient.GetStringAsync($"{baseUrl}/api/state");
                JObject state = JObject.Parse(stateJson);
                bool connected = state["connected"]?.ToObject<bool>() ?? false;
                Tg51BridgeStatus.Text = connected ? "Bridge connected" : "Bridge reachable";
            }
            catch (Exception ex)
            {
                Tg51BridgeStatus.Text = "Bridge unavailable";
                MessageBox.Show("Tank bridge check failed: " + ex.Message);
            }
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

                AddOrUpdateCurrentMeasurement();

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
            if (string.IsNullOrWhiteSpace(namedRange))
            {
                return 0;
            }

            if (workbook == null)
            {
                if (excelApp == null)
                {
                    excelApp = new Excel.Application();
                }

                if (machine == null || string.IsNullOrWhiteSpace(machine.Path))
                {
                    return 0;
                }

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
            expectedBiasVoltage = 300;
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
            Excel.Range range = null;
            try
            {
                if (!EnsureExcelWorkbookReady(showErrors: false))
                {
                    return null;
                }

                excelApp.Visible = true;

                Excel.Name excelNamedRange = workbook.Names.Item(namedRange);
                range = excelNamedRange.RefersToRange;
                range.Worksheet.Activate();
                range.Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error accessing named range.\nPlease verify that a valid energy and test is selected in the dropdowns. " + ex.Message);
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
            tg51TankStatusTimer.Stop();

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

            // Handle WebSocket cleanup without surfacing benign close races to the user.
            if (wsClient != null)
            {
                try
                {
                    WebSocketState state = wsClient.State;
                    if (state == WebSocketState.Open ||
                        state == WebSocketState.CloseReceived ||
                        state == WebSocketState.CloseSent)
                    {
                        using CancellationTokenSource closeTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(2));
                        wsClient.CloseAsync(WebSocketCloseStatus.NormalClosure, "Closing", closeTimeout.Token).GetAwaiter().GetResult();
                    }
                }
                catch (OperationCanceledException ex)
                {
                    Console.WriteLine($"WebSocket close timed out during shutdown: {ex.Message}");
                }
                catch (ObjectDisposedException ex)
                {
                    Console.WriteLine($"WebSocket was already disposed during shutdown: {ex.Message}");
                }
                catch (WebSocketException ex)
                {
                    Console.WriteLine($"WebSocket was already closed or aborted during shutdown: {ex.Message}");
                }
                catch (InvalidOperationException ex)
                {
                    Console.WriteLine($"WebSocket was not in a closable state during shutdown: {ex.Message}");
                }
                finally
                {
                    try
                    {
                        cancellationTokenSource?.Cancel();
                    }
                    catch (ObjectDisposedException)
                    {
                    }

                    try
                    {
                        wsClient.Dispose();
                    }
                    catch (ObjectDisposedException)
                    {
                    }

                    cancellationTokenSource?.Dispose();
                    cancellationTokenSource = null;
                    wsClient = null;
                    doseXConnected = false;
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
                int energyValue = ExtractNumber(selectedEnergy);
                               
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
                        Results.Items.Clear();

                        if (!EnsureExcelWorkbookReady(showErrors: false))
                        {
                            return;
                        }

                        if (!NamedRangeExists(currentNamedRange))
                        {
                            worksheet = null;
                            statusLabel.Text = $"Named range not found: {currentNamedRange}";
                            return;
                        }

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
            if (!EnsureExcelWorkbookReady(showErrors: false))
            {
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
            if (Discover.Text == "Discover") { DiscoverServices(true); }
            else { ConnectToWebSocketAsync(true); }
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

                // Wait for the response, but do not let a missing DoseX reply strand the UI.
                Task completedTask = await Task.WhenAny(controlTcs.Task, Task.Delay(TimeSpan.FromSeconds(5)));
                if (completedTask != controlTcs.Task)
                {
                    Console.WriteLine("Timed out waiting for control token response.");
                    return false;
                }

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

        private async Task SetMeasurementConfigurationAsync(string measurementTimerMode = "trigger")
        {
            try
            {
                var config = new
                {
                    cmd = "change_values",
                    values = new Dictionary<string, object>
            {
                { "biasVoltage", new { value = expectedBiasVoltage, unit = "V" } },  // Set bias voltage
                { "highVoltageEnabled", true },  // Enable high voltage
                { "measurementMode", "charge" },  // Set measurement mode to charge
                { "autoReset", true },
                { "sensitivity", "mid" },
                { "measurementTimerMode", measurementTimerMode },
                { "measurementStartType", measurementTimerMode },
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

                Task completedTask = await Task.WhenAny(loginTcs.Task, Task.Delay(TimeSpan.FromSeconds(5)));
                if (completedTask != loginTcs.Task)
                {
                    Console.WriteLine("Timed out waiting for admin login response.");
                    return false;
                }

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
                suppressNextMeasurementCommit = false;
                MessageBox.Show("Error starting measurement: " + ex.Message);
                Console.WriteLine("Error starting measurement: " + ex.Message);
            }
        }

        private async Task StopMeasurementAsync()
        {
            try
            {
                var request = new
                {
                    cmd = "measurement",
                    value = "stop"
                };
                string jsonRequest = JsonConvert.SerializeObject(request);
                var buffer = Encoding.UTF8.GetBytes(jsonRequest);
                var segment = new ArraySegment<byte>(buffer);

                await wsClient.SendAsync(segment, WebSocketMessageType.Text, true, cancellationTokenSource.Token);
                Console.WriteLine("Measurement stop requested.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error stopping measurement: " + ex.Message);
                Console.WriteLine("Error stopping measurement: " + ex.Message);
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
                expectedBiasVoltage = 300;
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
            await RunBackgroundMeasurementAsync();
        }

        private async Task RunBackgroundMeasurementAsync()
        {
            bool isControlAvailable = await RequestControlTokenAsync();  // Check control status before proceeding

            if (!isControlAvailable)
            {
                MessageBox.Show("Please free control from the electrometer by pressing the 'cloud' at the top.");
            }
            else
            {
                suppressNextMeasurementCommit = true;
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

    public class Tg51RunLogger
    {
        public string FilePath { get; private set; }
        public string StateFilePath { get; private set; }
        private readonly string machinePath;

        public Tg51RunLogger(string machinePath)
        {
            this.machinePath = machinePath ?? string.Empty;
            string directory = Path.GetDirectoryName(machinePath);
            if (string.IsNullOrEmpty(directory))
            {
                directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            Directory.CreateDirectory(directory);
            FilePath = Path.Combine(directory, "TG-51 DoseX.xml");

            string stateDirectory = Path.Combine(Path.GetTempPath(), "DoseXPDMTool", "TG51");
            Directory.CreateDirectory(stateDirectory);
            string stateName = string.IsNullOrWhiteSpace(machinePath)
                ? "TG-51 DoseX.state.xml"
                : Regex.Replace(Path.GetFileNameWithoutExtension(machinePath), @"[^\w\-. ]", "_") + ".TG-51 DoseX.state.xml";
            StateFilePath = Path.Combine(stateDirectory, stateName);
        }

        public Tg51Run Load()
        {
            Tg51Run stateRun = TryLoadRun(StateFilePath);
            if (stateRun != null)
            {
                return stateRun;
            }

            Tg51Run legacyRun = TryLoadRun(FilePath);
            if (legacyRun != null)
            {
                return legacyRun;
            }

            Tg51DoseXExport export = TryLoadExport(FilePath);
            if (export != null)
            {
                return export.ToRun(machinePath, FilePath);
            }

            return null;
        }

        private Tg51Run TryLoadRun(string path)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            try
            {
                var serializer = new XmlSerializer(typeof(Tg51Run));
                using FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return serializer.Deserialize(fs) as Tg51Run;
            }
            catch
            {
                return null;
            }
        }

        private Tg51DoseXExport TryLoadExport(string path)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            try
            {
                var serializer = new XmlSerializer(typeof(Tg51DoseXExport));
                using FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return serializer.Deserialize(fs) as Tg51DoseXExport;
            }
            catch
            {
                return null;
            }
        }

        public void Save(Tg51Run run)
        {
            if (run == null)
            {
                return;
            }

            run.XmlPath = FilePath;

            SaveFullState(run);

            Tg51DoseXExport export = Tg51DoseXExport.FromRun(run);
            if (!string.IsNullOrWhiteSpace(run.TemplateXmlPath) && File.Exists(run.TemplateXmlPath))
            {
                try
                {
                    SaveMergedTemplate(run.TemplateXmlPath, export);
                    return;
                }
                catch
                {
                    // Fall back to the stripped DoseX session export if the uploaded template cannot be merged.
                }
            }

            var exportSerializer = new XmlSerializer(typeof(Tg51DoseXExport));
            using FileStream fs = new FileStream(FilePath, FileMode.Create);
            exportSerializer.Serialize(fs, export);
        }

        private void SaveFullState(Tg51Run run)
        {
            var serializer = new XmlSerializer(typeof(Tg51Run));
            using FileStream fs = new FileStream(StateFilePath, FileMode.Create);
            serializer.Serialize(fs, run);
        }

        private void SaveMergedTemplate(string templatePath, Tg51DoseXExport export)
        {
            XDocument document = XDocument.Load(templatePath, LoadOptions.PreserveWhitespace);
            XElement root = document.Root ?? throw new InvalidOperationException("Template XML has no root element.");

            XElement photonParent = GetOrCreateChild(root, "PhotonCalibrations");
            foreach (var calibration in export.PhotonCalibrations ?? new List<Tg51DoseXCalibration>())
            {
                MergeCalibration(photonParent, "TG51_Photon_Full", calibration);
            }

            XElement electronParent = GetOrCreateChild(root, "ElectronCalibrations");
            foreach (var calibration in export.ElectronCalibrations ?? new List<Tg51DoseXCalibration>())
            {
                MergeCalibration(electronParent, "TG51_Electron_Full", calibration);
            }

            document.Save(FilePath);
        }

        private static void MergeCalibration(XElement parent, string elementName, Tg51DoseXCalibration source)
        {
            XElement calibration = parent
                .Elements()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, elementName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(GetChildValue(element, "NRG"), source.NRG, StringComparison.OrdinalIgnoreCase));

            if (calibration == null)
            {
                calibration = BuildCalibrationElement(parent.Name.Namespace, elementName, source);
                parent.Add(calibration);
                return;
            }

            SetChildValue(calibration, "IsEnabled", source.IsEnabled ? "true" : "false", createIfMissing: false);
            SetChildValue(calibration, "HighVoltage", source.HighVoltage.ToString(CultureInfo.InvariantCulture), createIfMissing: false);
            SetChildValue(calibration, "LowVoltage", source.LowVoltage.ToString(CultureInfo.InvariantCulture), createIfMissing: false);
            SetChildValue(calibration, "OppositeVoltage", source.OppositeVoltage.ToString(CultureInfo.InvariantCulture), createIfMissing: false);

            if (source.Depth.HasValue)
            {
                SetChildValue(calibration, "Depth", source.Depth.Value.ToString("0.###", CultureInfo.InvariantCulture), createIfMissing: true);
            }

            if (source.MeasurementDepth.HasValue)
            {
                SetChildValue(calibration, "MeasurementDepth", source.MeasurementDepth.Value.ToString("0.###", CultureInfo.InvariantCulture), createIfMissing: true);
            }

            ReplaceMeasurementArray(calibration, "HighMeasurements", source.HighMeasurements);
            ReplaceMeasurementArray(calibration, "LowMeasurements", source.LowMeasurements);
            ReplaceMeasurementArray(calibration, "OppositeMeasurements", source.OppositeMeasurements);
        }

        private static XElement BuildCalibrationElement(XNamespace ns, string elementName, Tg51DoseXCalibration source)
        {
            var element = new XElement(ns + elementName,
                new XElement(ns + "IsEnabled", source.IsEnabled ? "true" : "false"),
                new XElement(ns + "NRG", source.NRG ?? string.Empty));

            bool isElectron = string.Equals(source.Modality, "Electron", StringComparison.OrdinalIgnoreCase);
            if (!isElectron)
            {
                element.Add(new XElement(ns + "IsTMR", source.IsTMR ? "true" : "false"));
            }

            element.Add(
                new XElement(ns + "UseAssignedPionPpol", source.UseAssignedPionPpol ? "true" : "false"),
                new XElement(ns + "SSD", source.SSD),
                new XElement(ns + "X", source.X),
                new XElement(ns + "Y", source.Y),
                new XElement(ns + "DeliveredMU", source.DeliveredMU));

            if (isElectron)
            {
                element.Add(
                    new XElement(ns + "MeasureAtShift", source.MeasureAtShift ? "true" : "false"),
                    new XElement(ns + "MeasurementDepth", source.MeasurementDepth?.ToString("0.###", CultureInfo.InvariantCulture) ?? string.Empty));
            }
            else
            {
                element.Add(new XElement(ns + "Depth", source.Depth?.ToString("0.###", CultureInfo.InvariantCulture) ?? string.Empty));
            }

            element.Add(
                new XElement(ns + "HighVoltage", source.HighVoltage),
                new XElement(ns + "LowVoltage", source.LowVoltage),
                new XElement(ns + "OppositeVoltage", source.OppositeVoltage),
                BuildMeasurementArray(ns, "HighMeasurements", source.HighMeasurements),
                BuildMeasurementArray(ns, "LowMeasurements", source.LowMeasurements),
                BuildMeasurementArray(ns, "OppositeMeasurements", source.OppositeMeasurements));

            if (isElectron)
            {
                element.Add(new XElement(ns + "PQgrMeasurements"));
            }

            element.Add(
                new XElement(ns + "AdjustmentNeeded", "false"),
                new XElement(ns + "AdjustedMeasurements"));

            return element;
        }

        private static void ReplaceMeasurementArray(XElement calibration, string arrayName, List<double> values)
        {
            XElement replacement = BuildMeasurementArray(calibration.Name.Namespace, arrayName, values);
            XElement existing = GetChild(calibration, arrayName);
            if (existing != null)
            {
                existing.ReplaceWith(replacement);
                return;
            }

            XElement insertionPoint = GetChild(calibration, "AdjustmentNeeded") ?? GetChild(calibration, "AdjustedMeasurements");
            if (insertionPoint != null)
            {
                insertionPoint.AddBeforeSelf(replacement);
            }
            else
            {
                calibration.Add(replacement);
            }
        }

        private static XElement BuildMeasurementArray(XNamespace ns, string arrayName, List<double> values)
        {
            var element = new XElement(ns + arrayName);
            foreach (double value in values ?? new List<double>())
            {
                element.Add(new XElement(ns + "double", value.ToString("G17", CultureInfo.InvariantCulture)));
            }

            return element;
        }

        private static XElement GetOrCreateChild(XElement parent, string localName)
        {
            XElement child = GetChild(parent, localName);
            if (child != null)
            {
                return child;
            }

            child = new XElement(parent.Name.Namespace + localName);
            parent.Add(child);
            return child;
        }

        private static XElement GetChild(XElement parent, string localName)
        {
            return parent.Elements().FirstOrDefault(element => string.Equals(element.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetChildValue(XElement parent, string localName)
        {
            return GetChild(parent, localName)?.Value?.Trim() ?? string.Empty;
        }

        private static void SetChildValue(XElement parent, string localName, string value, bool createIfMissing)
        {
            XElement child = GetChild(parent, localName);
            if (child == null)
            {
                if (!createIfMissing)
                {
                    return;
                }

                child = new XElement(parent.Name.Namespace + localName);
                parent.Add(child);
            }

            child.Value = value ?? string.Empty;
        }
    }

    public static class Tg51PreviewMath
    {
        public static double? CalculatePpol(IEnumerable<double> highReadings, IEnumerable<double> oppositeReadings)
        {
            var high = highReadings?.ToList() ?? new List<double>();
            var opposite = oppositeReadings?.ToList() ?? new List<double>();
            if (high.Count == 0 || opposite.Count == 0)
            {
                return null;
            }

            double reference = Math.Abs(high.Average());
            if (reference <= 0)
            {
                return null;
            }

            return (Math.Abs(high.Average()) + Math.Abs(opposite.Average())) / (2.0 * reference);
        }

        public static double? CalculatePion(IEnumerable<double> highReadings, IEnumerable<double> lowReadings)
        {
            var high = highReadings?.ToList() ?? new List<double>();
            var low = lowReadings?.ToList() ?? new List<double>();
            if (high.Count == 0 || low.Count == 0)
            {
                return null;
            }

            double highAbs = Math.Abs(high.Average());
            double lowAbs = Math.Abs(low.Average());
            if (highAbs <= 0 || lowAbs <= 0)
            {
                return null;
            }

            double voltageRatioSquared = Math.Pow(300.0 / 150.0, 2);
            double denominator = highAbs / lowAbs - voltageRatioSquared;
            if (Math.Abs(denominator) < 0.000001)
            {
                return null;
            }

            return (1.0 - voltageRatioSquared) / denominator;
        }

        public static double? CalculateRoughOutputPerMu(
            IEnumerable<double> highReadings,
            double? ppol,
            double? pion,
            double? detectorCalibrationFactor,
            double? temperatureC,
            double? pressureMmHg,
            double? deliveredMu,
            double? prp,
            double? doseToTissueCorrection,
            double? qualityCorrectionFactor = null)
        {
            var high = highReadings?.ToList() ?? new List<double>();
            if (high.Count == 0 ||
                !ppol.HasValue ||
                !pion.HasValue ||
                !detectorCalibrationFactor.HasValue ||
                !temperatureC.HasValue ||
                !pressureMmHg.HasValue ||
                !deliveredMu.HasValue ||
                deliveredMu.Value <= 0)
            {
                return null;
            }

            double ptp = CalculatePtp(temperatureC.Value, pressureMmHg.Value);
            double correction = ptp * ppol.Value * pion.Value * prp.GetValueOrDefault(1.0) * doseToTissueCorrection.GetValueOrDefault(1.0) * qualityCorrectionFactor.GetValueOrDefault(1.0);
            double doseGy = Math.Abs(high.Average()) * 1e-9 * detectorCalibrationFactor.Value * correction;
            return doseGy * 100.0 / deliveredMu.Value;
        }

        public static double? CalculateReferenceOutputPerMu(double? measuredPointOutputPerMu, double? clinicalPddOrTmr)
        {
            if (!measuredPointOutputPerMu.HasValue || !clinicalPddOrTmr.HasValue)
            {
                return null;
            }

            double factor = clinicalPddOrTmr.Value;
            if (factor > 2.0)
            {
                factor /= 100.0;
            }

            if (factor <= 0)
            {
                return null;
            }

            return measuredPointOutputPerMu.Value / factor;
        }

        public static double? CalculatePhotonKq(double? measuredPdd10, double? a, double? b, double? c)
        {
            if (!measuredPdd10.HasValue || !a.HasValue || !b.HasValue || !c.HasValue)
            {
                return null;
            }

            double pdd10 = measuredPdd10.Value;
            return a.Value + (b.Value * pdd10 / 1000.0) + (c.Value * pdd10 * pdd10 / 100000.0);
        }

        public static double CalculatePtp(double temperatureC, double pressureMmHg)
        {
            return ((273.2 + temperatureC) / 295.2) * (760.0 / pressureMmHg);
        }
    }

    [XmlRoot("Session")]
    public class Tg51DoseXExport
    {
        [XmlIgnore]
        public string RunId { get; set; }

        public string MachineName { get; set; }
        public string MachineType { get; set; }

        [XmlIgnore]
        public DateTime StartedAt { get; set; }

        [XmlIgnore]
        public DateTime? CompletedAt { get; set; }

        public int HighVoltage { get; set; } = 300;
        public int LowVoltage { get; set; } = 150;
        public int OppositeVoltage { get; set; } = -300;

        [XmlArrayItem("TG51_Photon_Full")]
        public List<Tg51DoseXCalibration> PhotonCalibrations { get; set; } = new List<Tg51DoseXCalibration>();

        [XmlArrayItem("TG51_Electron_Full")]
        public List<Tg51DoseXCalibration> ElectronCalibrations { get; set; } = new List<Tg51DoseXCalibration>();

        public static Tg51DoseXExport FromRun(Tg51Run run)
        {
            var export = new Tg51DoseXExport
            {
                RunId = run.RunId ?? string.Empty,
                MachineName = run.MachineName ?? string.Empty,
                MachineType = run.MachineType ?? string.Empty,
                StartedAt = run.StartedAt,
                CompletedAt = run.CompletedAt
            };

            foreach (var point in run.Points ?? new List<Tg51Point>())
            {
                var calibration = Tg51DoseXCalibration.FromPoint(point);
                if (string.Equals(point.Modality, "Electron", StringComparison.OrdinalIgnoreCase))
                {
                    export.ElectronCalibrations.Add(calibration);
                }
                else
                {
                    export.PhotonCalibrations.Add(calibration);
                }
            }

            return export;
        }

        public Tg51Run ToRun(string machinePath, string xmlPath)
        {
            var run = new Tg51Run
            {
                RunId = string.IsNullOrWhiteSpace(RunId) ? $"TG51_{DateTime.Now:yyyyMMdd_HHmmss}" : RunId,
                MachineName = MachineName ?? string.Empty,
                MachineType = MachineType ?? string.Empty,
                MachinePath = machinePath ?? string.Empty,
                StartedAt = StartedAt == default ? DateTime.Now : StartedAt,
                CompletedAt = CompletedAt,
                XmlPath = xmlPath ?? string.Empty,
                BridgeUrl = string.Empty,
                BridgeEnabled = false
            };

            foreach (var calibration in PhotonCalibrations ?? new List<Tg51DoseXCalibration>())
            {
                run.Points.Add(calibration.ToPoint("Photon"));
            }

            foreach (var calibration in ElectronCalibrations ?? new List<Tg51DoseXCalibration>())
            {
                run.Points.Add(calibration.ToPoint("Electron"));
            }

            return run;
        }
    }

    public class Tg51DoseXCalibration
    {
        public bool IsEnabled { get; set; } = true;
        public string NRG { get; set; }
        [XmlIgnore]
        public string Modality { get; set; }
        [XmlIgnore]
        public string DepthLabel { get; set; }
        public bool IsTMR { get; set; } = true;
        public bool UseAssignedPionPpol { get; set; } = true;
        public int SSD { get; set; } = 100;
        public int X { get; set; } = 10;
        public int Y { get; set; } = 10;
        public int DeliveredMU { get; set; } = 100;
        public double? Depth { get; set; }
        public double? MeasurementDepth { get; set; }
        public bool MeasureAtShift { get; set; } = true;

        [XmlIgnore]
        public double? RawXmm { get; set; }
        [XmlIgnore]
        public double? RawYmm { get; set; }
        [XmlIgnore]
        public double? RawZmm { get; set; }
        public int HighVoltage { get; set; } = 300;
        public int LowVoltage { get; set; } = 150;
        public int OppositeVoltage { get; set; } = -300;

        [XmlArrayItem("double")]
        public List<double> HighMeasurements { get; set; } = new List<double>();

        [XmlArrayItem("double")]
        public List<double> LowMeasurements { get; set; } = new List<double>();

        [XmlArrayItem("double")]
        public List<double> OppositeMeasurements { get; set; } = new List<double>();

        [XmlIgnore]
        [XmlArrayItem("Measurement")]
        public List<Tg51DoseXMeasurement> Measurements { get; set; } = new List<Tg51DoseXMeasurement>();

        public bool AdjustmentNeeded { get; set; } = false;

        [XmlArrayItem("double")]
        public List<double> AdjustedMeasurements { get; set; } = new List<double>();

        public bool ShouldSerializeDepth()
        {
            return Depth.HasValue;
        }

        public bool ShouldSerializeMeasurementDepth()
        {
            return MeasurementDepth.HasValue;
        }

        public bool ShouldSerializeIsTMR()
        {
            return !string.Equals(Modality, "Electron", StringComparison.OrdinalIgnoreCase);
        }

        public bool ShouldSerializeMeasureAtShift()
        {
            return string.Equals(Modality, "Electron", StringComparison.OrdinalIgnoreCase);
        }

        public static Tg51DoseXCalibration FromPoint(Tg51Point point)
        {
            bool isElectron = string.Equals(point.Modality, "Electron", StringComparison.OrdinalIgnoreCase);
            var calibration = new Tg51DoseXCalibration
            {
                NRG = point.Energy ?? string.Empty,
                Modality = point.Modality ?? string.Empty,
                DepthLabel = point.DepthLabel ?? string.Empty,
                Depth = isElectron ? null : point.ClinicalDepthCm,
                MeasurementDepth = isElectron ? point.ClinicalDepthCm : null,
                RawXmm = point.RawCrosslineMm,
                RawYmm = point.RawInlineMm,
                RawZmm = point.RawDepthMm
            };

            foreach (var reading in point.Readings ?? new List<Tg51Reading>())
            {
                if (reading.BiasVoltage == 300)
                {
                    calibration.HighMeasurements.Add(reading.ChargeNc);
                }
                else if (reading.BiasVoltage == 150)
                {
                    calibration.LowMeasurements.Add(reading.ChargeNc);
                }
                else if (reading.BiasVoltage == -300)
                {
                    calibration.OppositeMeasurements.Add(reading.ChargeNc);
                }

                calibration.Measurements.Add(Tg51DoseXMeasurement.FromReading(reading));
            }

            return calibration;
        }

        public Tg51Point ToPoint(string fallbackModality)
        {
            string modality = string.IsNullOrWhiteSpace(Modality) ? fallbackModality : Modality;
            var point = new Tg51Point
            {
                PointId = $"{NRG}_{(string.Equals(modality, "Electron", StringComparison.OrdinalIgnoreCase) ? MeasurementDepth : Depth):0.###}",
                Energy = NRG ?? string.Empty,
                Modality = modality ?? string.Empty,
                DepthLabel = DepthLabel ?? string.Empty,
                ClinicalDepthCm = string.Equals(modality, "Electron", StringComparison.OrdinalIgnoreCase) ? MeasurementDepth : Depth,
                RawCrosslineMm = RawXmm,
                RawInlineMm = RawYmm,
                RawDepthMm = RawZmm,
                LastBridgeSnapshotJson = string.Empty
            };

            if (Measurements != null && Measurements.Count > 0)
            {
                foreach (var measurement in Measurements)
                {
                    point.Readings.Add(measurement.ToReading());
                }
            }
            else
            {
                AddBucketReadings(point, 300, HighMeasurements);
                AddBucketReadings(point, 150, LowMeasurements);
                AddBucketReadings(point, -300, OppositeMeasurements);
            }

            point.Status =
                point.Readings.Count(reading => reading.BiasVoltage == 300) >= 3 &&
                point.Readings.Count(reading => reading.BiasVoltage == -300) >= 3 &&
                point.Readings.Count(reading => reading.BiasVoltage == 150) >= 3
                    ? "Complete"
                    : "Active";

            return point;
        }

        private static void AddBucketReadings(Tg51Point point, int biasVoltage, List<double> charges)
        {
            int repeat = 1;
            foreach (double charge in charges ?? new List<double>())
            {
                point.Readings.Add(new Tg51Reading
                {
                    BiasVoltage = biasVoltage,
                    RepeatNumber = repeat++,
                    ChargeNc = charge,
                    RecordedAt = DateTime.MinValue,
                    DoseXHighVoltageEnabled = true,
                    DoseXBiasVoltage = biasVoltage,
                    BridgeSnapshotJson = string.Empty
                });
            }
        }
    }

    public class Tg51DoseXMeasurement
    {
        public int BiasVoltage { get; set; }
        public string BiasLabel { get; set; }
        public int RepeatNumber { get; set; }
        public double ChargeNc { get; set; }
        public DateTime RecordedAt { get; set; }

        public static Tg51DoseXMeasurement FromReading(Tg51Reading reading)
        {
            return new Tg51DoseXMeasurement
            {
                BiasVoltage = reading.BiasVoltage,
                BiasLabel = reading.BiasVoltage == 150 ? "50%" : $"{reading.BiasVoltage:+0;-0;0} V",
                RepeatNumber = reading.RepeatNumber,
                ChargeNc = reading.ChargeNc,
                RecordedAt = reading.RecordedAt
            };
        }

        public Tg51Reading ToReading()
        {
            return new Tg51Reading
            {
                BiasVoltage = BiasVoltage,
                RepeatNumber = RepeatNumber,
                ChargeNc = ChargeNc,
                RecordedAt = RecordedAt,
                DoseXHighVoltageEnabled = true,
                DoseXBiasVoltage = BiasVoltage,
                BridgeSnapshotJson = string.Empty
            };
        }
    }

    public class Tg51Run
    {
        public string RunId { get; set; }
        public string MachineName { get; set; }
        public string MachineType { get; set; }
        public string MachinePath { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime? CompletedAt { get; set; }
        public string BridgeUrl { get; set; }
        public bool BridgeEnabled { get; set; }
        public string XmlPath { get; set; }
        public string TemplateXmlPath { get; set; }
        public List<Tg51Point> Points { get; set; } = new List<Tg51Point>();
    }

    public class Tg51Point
    {
        public string PointId { get; set; }
        public string Energy { get; set; }
        public string Modality { get; set; }
        public string DepthLabel { get; set; }
        public double? ClinicalDepthCm { get; set; }
        public double? RawCrosslineMm { get; set; }
        public double? RawInlineMm { get; set; }
        public double? RawDepthMm { get; set; }
        public string Status { get; set; }
        public string LastBridgeSnapshotJson { get; set; }
        public List<Tg51Reading> Readings { get; set; } = new List<Tg51Reading>();
    }

    public class Tg51Reading
    {
        public int BiasVoltage { get; set; }
        public int RepeatNumber { get; set; }
        public double ChargeNc { get; set; }
        public DateTime RecordedAt { get; set; }
        public bool DoseXHighVoltageEnabled { get; set; }
        public int DoseXBiasVoltage { get; set; }
        public string BridgeSnapshotJson { get; set; }
    }

    public class Tg51WaitDialog : Form
    {
        public Tg51WaitDialog(string message)
        {
            Text = "TG-51 Wait";
            Size = new Size(360, 135);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ControlBox = false;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            System.Windows.Forms.Label messageLabel = new System.Windows.Forms.Label
            {
                AutoSize = false,
                Text = message,
                Location = new System.Drawing.Point(18, 16),
                Size = new Size(310, 38),
                TextAlign = ContentAlignment.MiddleLeft
            };

            ProgressBar progress = new ProgressBar
            {
                Location = new System.Drawing.Point(18, 65),
                Size = new Size(310, 18),
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };

            Controls.Add(messageLabel);
            Controls.Add(progress);
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
