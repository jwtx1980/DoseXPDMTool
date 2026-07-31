(function () {
  const readiness = window.TankWorkflowReadiness;

  const presets = {
    center: { label: "Iso", x: 0, y: 0, z: 0 },
    left: { label: "X -5 cm", x: -50, y: 0, z: 0 },
    right: { label: "X +5 cm", x: 50, y: 0, z: 0 },
    up: { label: "Z +5 cm", x: 0, y: 0, z: 50 },
    down: { label: "Z -5 cm", x: 0, y: 0, z: -50 }
  };

  let snapshot = null;
  let pendingMessage = "";
  let opabQueue = null;
  let activeQueueRunId = null;
  let activeQueueRun = null;
  let selectedScanResult = null;
  let selectedScanJobId = null;
  let liveScanJobId = null;
  let liveScanResult = null;
  let coordinateInputsTouched = false;
  let biasWorkflowComplete = readStoredBiasWorkflowComplete();
  let centeringWorkflowComplete = false;
  let centeringStatusText = "Not centered";
  let centeringStatusClass = "badge neutral";
  let workflowProgress = null;
  let workflowProgressTimer = null;
  let recoveryPromptOpen = false;
  let recoveryStartupCheckArmed = true;
  let apiWasOffline = false;

  const DEFAULT_TIMED_ACQUISITION_SECONDS = 10;
  const NORMALIZATION_DEPTH_FROM_ISO_MM = 15;
  const NORMALIZATION_POSITION_TOLERANCE_MM = 2;

  const elements = {
    connectButton: document.getElementById("connectButton"),
    disconnectButton: document.getElementById("disconnectButton"),
    sendMoveButton: document.getElementById("sendMoveButton"),
    goIsocenterButton: document.getElementById("goIsocenterButton"),
    goSurfaceButton: document.getElementById("goSurfaceButton"),
    biasButton: document.getElementById("biasButton"),
    biasOffButton: document.getElementById("biasOffButton"),
    backgroundButton: document.getElementById("backgroundButton"),
    normalizeButton: document.getElementById("normalizeButton"),
    prepareButton: document.getElementById("prepareButton"),
    runCenteringScanButton: document.getElementById("runCenteringScanButton"),
    applySensitivityButton: document.getElementById("applySensitivityButton"),
    clearLogButton: document.getElementById("clearLogButton"),
    crosslineInput: document.getElementById("crosslineInput"),
    inlineInput: document.getElementById("inlineInput"),
    depthInput: document.getElementById("depthInput"),
    speedInput: document.getElementById("speedInput"),
    centeringWidthInput: document.getElementById("centeringWidthInput"),
    turnAngleSelect: document.getElementById("turnAngleSelect"),
    detectorModeSelect: document.getElementById("detectorModeSelect"),
    biasModeSelect: document.getElementById("biasModeSelect"),
    fieldSensitivitySelect: document.getElementById("fieldSensitivitySelect"),
    referenceSensitivitySelect: document.getElementById("referenceSensitivitySelect"),
    statusPill: document.getElementById("statusPill"),
    endpointText: document.getElementById("endpointText"),
    targetState: document.getElementById("targetState"),
    callbackText: document.getElementById("callbackText"),
    positionText: document.getElementById("positionText"),
    temperatureText: document.getElementById("temperatureText"),
    pressureText: document.getElementById("pressureText"),
    biasText: document.getElementById("biasText"),
    measurementText: document.getElementById("measurementText"),
    signalText: document.getElementById("signalText"),
    controllerErrorText: document.getElementById("controllerErrorText"),
    landmarkText: document.getElementById("landmarkText"),
    preparationText: document.getElementById("preparationText"),
    doubleText: document.getElementById("doubleText"),
    floatText: document.getElementById("floatText"),
    spigotText: document.getElementById("spigotText"),
    centeringStatus: document.getElementById("centeringStatus"),
    workflowHint: document.getElementById("workflowHint"),
    workflowSteps: document.getElementById("workflowSteps"),
    logOutput: document.getElementById("logOutput"),
    chamberCanvas: document.getElementById("chamberCanvas"),
    fieldMeterFill: document.getElementById("fieldMeterFill"),
    fieldMeterText: document.getElementById("fieldMeterText"),
    referenceMeter: document.getElementById("referenceMeter"),
    referenceMeterFill: document.getElementById("referenceMeterFill"),
    referenceMeterText: document.getElementById("referenceMeterText")
  };

  Object.assign(elements, {
    opabPathInput: document.getElementById("opabPathInput"),
    opabFileInput: document.getElementById("opabFileInput"),
    loadOpabButton: document.getElementById("loadOpabButton"),
    queueCheckAllInput: document.getElementById("queueCheckAllInput"),
    runQueueButton: document.getElementById("runQueueButton"),
    abortQueueButton: document.getElementById("abortQueueButton"),
    queueState: document.getElementById("queueState"),
    queueItemsBody: document.getElementById("queueItemsBody"),
    queueCenterXInput: document.getElementById("queueCenterXInput"),
    queueCenterYInput: document.getElementById("queueCenterYInput"),
    queueSurfaceZInput: document.getElementById("queueSurfaceZInput"),
    queueCrosslineYInput: document.getElementById("queueCrosslineYInput"),
    queueNoBeamInput: document.getElementById("queueNoBeamInput"),
    scanResultState: document.getElementById("scanResultState"),
    scanResultSummary: document.getElementById("scanResultSummary"),
    scanResultCanvas: document.getElementById("scanResultCanvas"),
    scanResultJsonLink: document.getElementById("scanResultJsonLink"),
    scanResultAscLink: document.getElementById("scanResultAscLink")
  });

  restoreUserPreferences();

  async function api(path, options) {
    const response = await fetch(path, {
      headers: { "content-type": "application/json" },
      ...options
    });
    if (!response.ok) {
      let message = `${response.status} ${response.statusText}`;
      try {
        const body = await response.json();
        message = body.message || body.detail || message;
      } catch {
        message = await response.text() || message;
      }
      throw new Error(message);
    }
    return response.json();
  }

  async function post(path, body) {
    setPending("Command in progress");
    try {
      snapshot = await api(path, {
        method: "POST",
        body: body === undefined ? undefined : JSON.stringify(body)
      });
      pendingMessage = "";
      render();
    } catch (error) {
      pendingMessage = `Error: ${error.message}`;
      await pollState();
    }
  }

  async function pollState() {
    try {
      const wasConnected = !!snapshot?.connected;
      snapshot = await api("/api/state");
      if (apiWasOffline) {
        recoveryStartupCheckArmed = true;
      }
      apiWasOffline = false;
      if (wasConnected && !snapshot?.connected) {
        setBiasWorkflowComplete(false);
        setCenteringWorkflowComplete(false);
      }
      if (activeQueueRunId) {
        await pollQueueRun();
      }
      await pollLiveScan();
      render();
      await maybePromptForAcquisitionRecovery();
    } catch (error) {
      apiWasOffline = true;
      snapshot = null;
      pendingMessage = `API offline: ${error.message}`;
      render();
    }
  }

  function setPending(label) {
    pendingMessage = label;
    renderControls();
  }

  function readTargetInputs() {
    const relative = readRelativeTargetInputs();
    const iso = snapshot?.isocenter;
    const hasIso = iso && isFiniteNumber(iso.x) && isFiniteNumber(iso.y) && isFiniteNumber(iso.z);
    return {
      crossline: hasIso ? iso.x + relative.crossline : relative.crossline,
      inline: hasIso ? iso.y + relative.inline : relative.inline,
      depth: hasIso ? iso.z + relative.depth : relative.depth,
      speed: relative.speed
    };
  }

  function readRelativeTargetInputs() {
    return {
      crossline: clampInput(elements.crosslineInput),
      inline: clampInput(elements.inlineInput),
      depth: clampInput(elements.depthInput),
      speed: clampInput(elements.speedInput)
    };
  }

  function clampInput(input) {
    const min = Number(input.min);
    const max = Number(input.max);
    const fallback = Number(input.defaultValue) || 0;
    const value = Number(input.value);
    const clamped = Math.max(min, Math.min(max, Number.isFinite(value) ? value : fallback));
    input.value = clamped.toFixed(3);
    return clamped;
  }

  function setTarget(x, y, z) {
    elements.crosslineInput.value = x.toFixed(3);
    elements.inlineInput.value = y.toFixed(3);
    elements.depthInput.value = z.toFixed(3);
    readRelativeTargetInputs();
  }

  function render() {
    syncBiasWorkflowState();
    syncCoordinateInputsFromSnapshot();
    renderControls();
    renderWorkflowReadiness();
    renderTelemetry();
    renderQueue();
    renderSelectedScanResult();
    renderLog();
    drawGraph();
  }

  function syncBiasWorkflowState() {
    if (!snapshot?.connected && biasWorkflowComplete) {
      setBiasWorkflowComplete(false);
    }
    if (snapshot?.latestCentering && !centeringWorkflowComplete) {
      restoreCenteringFromSnapshot();
    } else if (!snapshot?.connected && centeringWorkflowComplete && !snapshot?.latestCentering) {
      setCenteringWorkflowComplete(false);
    }
  }

  function syncCoordinateInputsFromSnapshot() {
    if (coordinateInputsTouched || !snapshot?.connected) {
      return;
    }

    setInputIfFinite(elements.queueCenterXInput, snapshot?.isocenter?.x);
    setInputIfFinite(elements.queueCenterYInput, snapshot?.isocenter?.y);
    setInputIfFinite(elements.queueSurfaceZInput, snapshot?.surface?.z);
  }

  function setInputIfFinite(input, value) {
    const numeric = numberOrNull(value);
    if (input && isFiniteNumber(numeric)) {
      input.value = numeric.toFixed(3);
    }
  }

  function renderControls() {
    const connected = !!snapshot?.connected;
    const busy = !!snapshot?.busy || isCommandPending();
    const missingCoordinates = snapshot?.coordinateState?.missing || [];
    const detectorMode = readDetectorMode();
    const modeLabel = detectorModeLabel(detectorMode);
    const fieldOnly = detectorMode === "fieldOnly";
    const currentModeBackgroundReady = hasCurrentModeBackground();
    elements.connectButton.disabled = connected || busy;
    elements.disconnectButton.disabled = !connected || busy;
    elements.sendMoveButton.disabled = !connected || busy || !snapshot?.isocenter;
    elements.sendMoveButton.title = snapshot?.isocenter
      ? "Move to the entered position relative to isocenter"
      : "Set isocenter before using relative target moves";
    if (elements.goIsocenterButton) {
      elements.goIsocenterButton.disabled = !connected || busy || !snapshot?.isocenter;
      elements.goIsocenterButton.title = snapshot?.isocenter
        ? "Move to the stored tank isocenter"
        : "No tank isocenter is available";
    }
    elements.goSurfaceButton.disabled = !connected || busy || !snapshot?.surface;
    elements.goSurfaceButton.title = snapshot?.surface
      ? "Move to the stored tank surface"
      : "No tank surface is available";
    elements.biasButton.disabled = !connected || busy || !snapshot?.biasHvEnabled;
    elements.biasButton.title = snapshot?.biasHvEnabled
      ? "Apply configured detector bias"
      : "Bias/HV disabled for this run";
    elements.biasOffButton.disabled = !connected || busy;
    elements.biasModeSelect.disabled = !connected || busy;
    elements.applySensitivityButton.disabled = !connected || busy;
    elements.backgroundButton.disabled = !connected || busy || !biasWorkflowComplete;
    elements.backgroundButton.title = `Take ${modeLabel} background with beam off`;
    elements.normalizeButton.disabled = !connected || busy || !currentModeBackgroundReady;
    elements.normalizeButton.title = currentModeBackgroundReady
      ? `Take ${modeLabel} normalization with beam on`
      : `Take ${modeLabel} background first with beam off`;
    elements.prepareButton.disabled = !connected || busy || missingCoordinates.length > 0;
    elements.prepareButton.title = missingCoordinateMessage(missingCoordinates) || `Guide ${modeLabel} background, normalization, and scan readiness`;
    if (elements.runCenteringScanButton) {
      elements.runCenteringScanButton.disabled = !connected || busy || !!centeringReadinessIssue();
      elements.runCenteringScanButton.title = centeringReadinessIssue() || "Run a setup profile and optionally write the measured center to tank isocenter";
      elements.runCenteringScanButton.textContent = centeringWorkflowComplete ? "Rerun Centering" : "Run Centering Scan";
    }
    renderCenteringStatus();
    elements.referenceSensitivitySelect.disabled = fieldOnly;
    elements.referenceSensitivitySelect.title = fieldOnly
      ? "Reference channel is skipped in Field only mode"
      : "Reference chamber sensitivity";
    elements.referenceSensitivitySelect.closest("label")?.classList.toggle("disabled-control", fieldOnly);
    elements.loadOpabButton.disabled = busy;
    if (elements.queueCheckAllInput) {
      elements.queueCheckAllInput.disabled = !opabQueue?.items?.length || busy;
    }
    elements.runQueueButton.disabled = !opabQueue?.items?.length || busy;
    elements.abortQueueButton.disabled = !activeQueueRunId || activeQueueRun?.status === "completed" || activeQueueRun?.status === "aborted";

    const label = pendingMessage || missingCoordinateMessage(missingCoordinates) || snapshot?.busyLabel || "Ready";
    elements.targetState.textContent = label;
    elements.targetState.className = `badge ${busy || pendingMessage || missingCoordinates.length ? "warning" : "neutral"}`;

    if (!snapshot) {
      elements.statusPill.className = "status-pill disconnected";
      elements.statusPill.textContent = "API offline";
      return;
    }

    elements.statusPill.className = `status-pill ${snapshot.busy ? "busy" : snapshot.connected ? "connected" : "disconnected"}`;
    elements.statusPill.textContent = snapshot.busy ? snapshot.busyLabel : snapshot.connected ? "Connected" : "Disconnected";
    elements.endpointText.textContent = snapshot.localEndpoint || "169.254.40.1:1222";
  }

  function isCommandPending() {
    if (!pendingMessage || pendingMessage.startsWith("Error") || pendingMessage.startsWith("API offline")) {
      return false;
    }

    return [
      "Command in progress",
      "Loading ",
      "Starting queue run",
      "Aborting queue",
      "Prepare ",
      "Taking background",
      "Taking normalization",
      "Running centering scan",
      "Centering scan:",
      "Writing measured beam center"
    ].some((prefix) => pendingMessage.startsWith(prefix));
  }

  function missingCoordinateMessage(missing) {
    return readiness.missingCoordinateMessage(missing);
  }

  function renderWorkflowReadiness() {
    const workflowOptions = { biasReady: biasWorkflowComplete, centeringReady: centeringWorkflowComplete };
    if (elements.workflowHint) {
      const hint = workflowProgress
        ? { text: workflowProgressLabel(), className: "workflow-hint warning" }
        : readiness.workflowHint(snapshot, readDetectorMode(), workflowOptions);
      elements.workflowHint.textContent = hint.text;
      elements.workflowHint.className = hint.className;
    }

    if (!elements.workflowSteps) {
      return;
    }

    elements.workflowSteps.innerHTML = readiness.workflowSteps(snapshot, readDetectorMode(), workflowOptions)
      .map((step) => {
        const progress = workflowStepProgress(step.key);
        const enabled = !progress && workflowStepEnabled(step.key);
        const actionClass = progress ? "collecting" : enabled ? "active" : "inactive";
        const title = progress ? progress.label : workflowStepTitle(step);
        const detail = progress ? progress.detail : step.detail;
        return `
        <button type="button" class="workflow-step ${step.status} ${actionClass}" data-workflow-step="${escapeHtml(step.key)}" title="${escapeHtml(title)}" ${enabled ? "" : "disabled"}>
          <strong>${escapeHtml(step.label)}</strong>
          <span>${escapeHtml(detail)}</span>
          ${progress ? `<div class="workflow-progress" aria-hidden="true"><span style="width: ${progress.percent.toFixed(1)}%"></span></div>` : ""}
        </button>`;
      })
      .join("");
  }

  function startWorkflowProgress(key, label, durationSeconds) {
    clearWorkflowProgress();
    workflowProgress = {
      key,
      label,
      startedAt: Date.now(),
      durationMs: Math.max(1000, durationSeconds * 1000),
      percent: 0
    };
    updateWorkflowProgress();
    workflowProgressTimer = window.setInterval(() => {
      updateWorkflowProgress();
      renderControls();
      renderWorkflowReadiness();
    }, 250);
    renderControls();
    renderWorkflowReadiness();
  }

  function updateWorkflowProgress() {
    if (!workflowProgress) {
      return;
    }

    const elapsed = Date.now() - workflowProgress.startedAt;
    workflowProgress.percent = Math.min(95, Math.max(3, elapsed / workflowProgress.durationMs * 100));
  }

  function clearWorkflowProgress() {
    if (workflowProgressTimer) {
      window.clearInterval(workflowProgressTimer);
      workflowProgressTimer = null;
    }
    workflowProgress = null;
  }

  function workflowProgressLabel() {
    if (!workflowProgress) {
      return "";
    }

    const remainingMs = Math.max(0, workflowProgress.durationMs - (Date.now() - workflowProgress.startedAt));
    const remainingSeconds = Math.ceil(remainingMs / 1000);
    return `${workflowProgress.label}: collecting, about ${remainingSeconds} s remaining.`;
  }

  function workflowStepProgress(key) {
    if (!workflowProgress || workflowProgress.key !== key) {
      return null;
    }

    return {
      label: workflowProgress.label,
      detail: `Collecting ${Math.round(workflowProgress.percent)}%`,
      percent: workflowProgress.percent
    };
  }

  function workflowStepTitle(step) {
    if (workflowStepEnabled(step.key)) {
      return workflowStepActionLabel(step.key);
    }

    if (step.key === "coordinates" && step.status === "pending") {
      return "Set coordinates on the tank controller, then the app will continue.";
    }

    return step.detail;
  }

  function workflowStepActionLabel(key) {
    const modeLabel = detectorModeLabel(readDetectorMode());
    return {
      connected: snapshot?.connected ? "Disconnect from the CCU" : "Connect to the CCU",
      coordinates: missingCoordinateMessage(snapshot?.coordinateState?.missing || []) || "Coordinates are already set",
      bias: biasWorkflowLabel(),
      background: `Take ${modeLabel} background with beam off`,
      normalization: `Take ${modeLabel} normalization with beam on`,
      centering: "Run a setup profile and offer to write the measured center to tank isocenter"
    }[key] || "";
  }

  function workflowStepEnabled(key) {
    const connected = !!snapshot?.connected;
    const busy = !!snapshot?.busy || isCommandPending();
    const missingCoordinates = snapshot?.coordinateState?.missing || [];
    if (busy) {
      return false;
    }

    switch (key) {
      case "connected":
        return connected || !connected;
      case "coordinates":
        return connected && missingCoordinates.length > 0;
      case "bias":
        return connected;
      case "background":
        return connected && biasWorkflowComplete;
      case "normalization":
        return connected && biasWorkflowComplete && hasCurrentModeBackground();
      case "centering":
        return connected && biasWorkflowComplete && !centeringReadinessIssue();
      default:
        return false;
    }
  }

  async function runWorkflowStep(key) {
    if (!workflowStepEnabled(key)) {
      return;
    }

    switch (key) {
      case "connected":
        await post(snapshot?.connected ? "/api/disconnect" : "/api/connect");
        setBiasWorkflowComplete(false);
        setCenteringWorkflowComplete(false);
        render();
        return;
      case "coordinates":
        pendingMessage = missingCoordinateMessage(snapshot?.coordinateState?.missing || [])
          || "Coordinates are already set.";
        render();
        return;
      case "bias":
        await runBiasWorkflow();
        return;
      case "background":
        await takeBackground();
        return;
      case "normalization":
        await takeNormalization();
        return;
      case "centering":
        await runCenteringScan();
        return;
      default:
        return;
    }
  }

  async function runBiasWorkflow() {
    const mode = elements.biasModeSelect?.value || "zero";
    if (mode === "zero") {
      setBiasWorkflowComplete(true);
      pendingMessage = "Info: Bias/HV set to 0 V locally for diode-safe operation. No HV command was sent to the CCU.";
      render();
      return;
    }

    if (mode === "off") {
      await post("/api/bias-off");
      setBiasWorkflowComplete(true);
      render();
      return;
    }

    if (!snapshot?.biasHvEnabled) {
      pendingMessage = "Error: Applying HV is disabled. Keep 0 V for diode work, or set TANK_ALLOW_BIAS_HV=1 only for an intentional biased chamber setup.";
      render();
      return;
    }

    const confirmed = window.confirm("Apply configured high voltage to the detector channels? Do not continue with diode detectors connected.");
    if (!confirmed) {
      pendingMessage = "Bias/HV apply cancelled.";
      render();
      return;
    }

    await post("/api/bias");
    setBiasWorkflowComplete(true);
    render();
  }

  function setBiasWorkflowComplete(value) {
    biasWorkflowComplete = value === true;
    window.localStorage.setItem("tank.biasWorkflowComplete", biasWorkflowComplete ? "1" : "0");
  }

  function setCenteringWorkflowComplete(value, text) {
    centeringWorkflowComplete = value === true;
    centeringStatusText = text || (centeringWorkflowComplete ? "Centered" : "Not centered");
    centeringStatusClass = centeringWorkflowComplete ? "badge good" : "badge neutral";
    renderCenteringStatus();
  }

  function setCenteringStatus(text, className, complete = centeringWorkflowComplete) {
    centeringWorkflowComplete = complete === true;
    centeringStatusText = text;
    centeringStatusClass = className;
    renderCenteringStatus();
  }

  function restoreCenteringFromSnapshot() {
    const centering = snapshot?.latestCentering;
    if (!centering) {
      return;
    }

    const adjusted = numberOrNull(centering.adjustedCenterMm);
    if (isFiniteNumber(adjusted) && elements.queueCenterXInput) {
      elements.queueCenterXInput.value = adjusted.toFixed(3);
      coordinateInputsTouched = true;
    }

    const correction = numberOrNull(centering.isoCorrectionMm);
    const detail = isFiniteNumber(correction) && isFiniteNumber(adjusted)
      ? `Centered ${formatTrimmed(correction, 2)} mm; X ${formatTrimmed(adjusted, 3)}`
      : "Centered";
    setCenteringStatus(detail, "badge good", true);
  }

  function renderCenteringStatus() {
    if (!elements.centeringStatus) {
      return;
    }

    elements.centeringStatus.textContent = centeringStatusText;
    elements.centeringStatus.className = centeringStatusClass;
  }

  function readStoredBiasWorkflowComplete() {
    return window.localStorage.getItem("tank.biasWorkflowComplete") === "1";
  }

  function biasWorkflowLabel() {
    const mode = elements.biasModeSelect?.value || "zero";
    if (mode === "zero") {
      return "0 V selected; no HV command will be sent";
    }

    if (mode === "off") {
      return "Send HV off command";
    }

    return snapshot?.biasHvEnabled
      ? "Apply configured HV"
      : "HV apply disabled for diode safety";
  }

  function renderTelemetry() {
    const latest = snapshot?.latestStatus;
    const callbackParts = [];
    if (snapshot?.callbackEndpoint) {
      callbackParts.push(`listen ${snapshot.callbackEndpoint}`);
    }
    if (snapshot?.callbackPeerEndpoint) {
      callbackParts.push(`peer ${snapshot.callbackPeerEndpoint}`);
    }
    if (snapshot?.callbackConnectionCount) {
      callbackParts.push(`${snapshot.callbackConnectionCount} callback connection${snapshot.callbackConnectionCount === 1 ? "" : "s"}`);
    }
    if (elements.callbackText) {
      elements.callbackText.textContent = latest
        ? `frame ${latest.frameCount}, opcode ${latest.opcode}, ${latest.frameLength} bytes, ${formatTime(latest.timestamp, true)}`
        : callbackParts.length
          ? `waiting; ${callbackParts.join(", ")}`
          : "idle";
    }
    const position = formatRelativePosition(latest);
    elements.positionText.textContent = position.text;
    elements.positionText.title = position.title;
    if (elements.temperatureText) {
      elements.temperatureText.textContent = isFiniteNumber(snapshot?.environment?.temperatureC)
        ? `${formatNumber(snapshot.environment.temperatureC, 1)} C`
        : "-";
    }
    if (elements.pressureText) {
      elements.pressureText.textContent = isFiniteNumber(snapshot?.environment?.pressureHpa)
        ? `${formatNumber(snapshot.environment.pressureHpa, 1)} hPa`
        : "-";
    }
    const fieldOnly = readDetectorMode() === "fieldOnly";
    if (elements.biasText) {
      elements.biasText.textContent = fieldOnly
        ? `Field ${formatVolts(snapshot?.highVoltage?.fieldVolts)}`
        : `Field ${formatVolts(snapshot?.highVoltage?.fieldVolts)}, Reference ${formatVolts(snapshot?.highVoltage?.referenceVolts)}`;
    }
    renderSignalHealth(latest);
    renderControllerErrors();
    if (elements.landmarkText) {
      elements.landmarkText.textContent = `Isocenter ${formatLandmark(snapshot?.isocenter)}; Surface ${formatLandmark(snapshot?.surface)}`;
    }
    if (elements.preparationText) {
      elements.preparationText.textContent = formatPreparation(snapshot?.preparation);
    }
    if (elements.doubleText) {
      elements.doubleText.textContent = latest?.doubleCandidates?.length ? latest.doubleCandidates.join(" | ") : "-";
    }
    if (elements.floatText) {
      elements.floatText.textContent = latest?.floatCandidates?.length ? latest.floatCandidates.join(" | ") : "-";
    }
    renderChamberMeters(latest);
  }

  function formatRelativePosition(latest) {
    if (!latest || !isFiniteNumber(latest.x) || !isFiniteNumber(latest.y) || !isFiniteNumber(latest.z)) {
      return { text: "-", title: "" };
    }

    const iso = snapshot?.isocenter;
    const absolute = `Absolute X ${formatTrimmed(latest.x, 3)} mm, Y ${formatTrimmed(latest.y, 3)} mm, Z ${formatTrimmed(latest.z, 3)} mm`;
    if (!iso || !isFiniteNumber(iso.x) || !isFiniteNumber(iso.y) || !isFiniteNumber(iso.z)) {
      return { text: "Waiting for isocenter", title: absolute };
    }

    return {
      text: `X ${formatSigned(latest.x - iso.x, 2)} mm, Y ${formatSigned(latest.y - iso.y, 2)} mm, Z ${formatSigned(latest.z - iso.z, 2)} mm relative to iso`,
      title: absolute
    };
  }

  function renderChamberMeters(latest) {
    if (!elements.fieldMeterFill || !elements.fieldMeterText) {
      return;
    }

    const fieldOnly = readDetectorMode() === "fieldOnly";
    const values = liveDisplayValues(latest);
    updateChamberMeter(elements.fieldMeterFill, elements.fieldMeterText, values.field, values.max, values.showNormalized);
    if (elements.referenceMeter) {
      elements.referenceMeter.hidden = fieldOnly;
    }
    if (!fieldOnly) {
      updateChamberMeter(elements.referenceMeterFill, elements.referenceMeterText, values.reference, values.referenceMax, values.showNormalized);
    }
  }

  function liveDisplayValues(latest) {
    const fieldOnly = readDetectorMode() === "fieldOnly";
    const background = snapshot?.latestBackground;
    const normalization = snapshot?.latestNormalization;
    const backgroundMatchesMode = hasCurrentModeBackground();
    const normalizationMatchesMode = hasCurrentModeNormalization();
    const fieldBackground = backgroundMatchesMode && isFiniteNumber(background?.fieldDetectorValue)
      ? background.fieldDetectorValue
      : null;
    const fieldNormalization = normalizationMatchesMode && isFiniteNumber(normalization?.fieldDetectorValue)
      ? normalization.fieldDetectorValue
      : null;
    const effectiveFieldBackground = isFiniteNumber(fieldBackground)
      ? fieldBackground
      : fieldOnly && isFiniteNumber(fieldNormalization)
        ? 0
        : null;
    const referenceBackground = backgroundMatchesMode && isFiniteNumber(background?.referenceDetectorValue)
      ? background.referenceDetectorValue
      : null;
    const referenceNormalization = normalizationMatchesMode && isFiniteNumber(normalization?.referenceDetectorValue)
      ? normalization.referenceDetectorValue
      : null;
    const showNormalized = isFiniteNumber(effectiveFieldBackground)
      && isFiniteNumber(fieldNormalization)
      && Math.abs(fieldNormalization - effectiveFieldBackground) > 0.000001;
    const samples = (snapshot?.samples || []).slice(-80);
    const fieldValues = samples
      .map((sample) => showNormalized
        ? normalizeLiveValue(sample.fieldX10e3Pa, effectiveFieldBackground, fieldNormalization)
        : sample.fieldX10e3Pa)
      .filter(isFiniteNumber);
    const referenceValues = samples
      .map((sample) => showNormalized && !fieldOnly
        ? normalizeLiveValue(sample.referenceX10e3Pa, referenceBackground, referenceNormalization)
        : sample.referenceX10e3Pa)
      .filter(isFiniteNumber);
    const field = showNormalized
      ? normalizeLiveValue(latest?.fieldCurrentX10e3Pa, effectiveFieldBackground, fieldNormalization)
      : latest?.fieldCurrentX10e3Pa;
    const reference = showNormalized && !fieldOnly
      ? normalizeLiveValue(latest?.referenceCurrentX10e3Pa, referenceBackground, referenceNormalization)
      : latest?.referenceCurrentX10e3Pa;

    return {
      field,
      reference,
      showNormalized,
      max: showNormalized ? 120 : Math.max(0.01, ...fieldValues.map((value) => Math.abs(value))) * 1.05,
      referenceMax: showNormalized ? 120 : Math.max(0.01, ...referenceValues.map((value) => Math.abs(value))) * 1.05
    };
  }

  function updateChamberMeter(fill, text, value, max, showNormalized) {
    if (!fill || !text) {
      return;
    }

    if (!isFiniteNumber(value) || !isFiniteNumber(max) || max <= 0) {
      fill.style.height = "0%";
      text.textContent = "-";
      return;
    }

    fill.style.height = `${clamp(Math.abs(value) / max * 100, 0, 100)}%`;
    text.textContent = showNormalized ? `${formatTrimmed(value, 1)}%` : formatChamber(value);
  }

  function renderSignalHealth(latest) {
    if (!elements.signalText) {
      return;
    }

    const samples = (snapshot?.samples || [])
      .filter((sample) => isFiniteNumber(sample.fieldX10e3Pa))
      .slice(-40);
    if (!latest || !samples.length) {
      elements.signalText.textContent = "No live chamber callback samples";
      elements.signalText.className = "signal-state warning";
      return;
    }

    const values = samples.map((sample) => sample.fieldX10e3Pa);
    const avg = values.reduce((sum, value) => sum + value, 0) / values.length;
    const min = Math.min(...values);
    const max = Math.max(...values);
    const span = max - min;
    const normalization = snapshot?.latestNormalization;
    const background = snapshot?.latestBackground;
    const normValue = hasCurrentModeNormalization() && isFiniteNumber(normalization?.fieldDetectorValue)
      ? normalization.fieldDetectorValue
      : null;
    const backgroundValue = hasCurrentModeBackground() && isFiniteNumber(background?.fieldDetectorValue)
      ? background.fieldDetectorValue
      : readDetectorMode() === "fieldOnly" && isFiniteNumber(normValue)
        ? 0
        : null;
    const corrected = isFiniteNumber(backgroundValue) ? avg - backgroundValue : avg;
    const normalizedPercent = isFiniteNumber(normValue) && isFiniteNumber(backgroundValue) && Math.abs(normValue - backgroundValue) > 0.000001
      ? (avg - backgroundValue) / (normValue - backgroundValue) * 100
      : isFiniteNumber(normValue) && Math.abs(normValue) > 0.000001
        ? avg / normValue * 100
        : null;

    let state = "usable";
    let message = `Live ${formatChamber(avg)} over ${values.length} samples`;
    const existingWarnings = latest.warnings || [];

    if (existingWarnings.length || Math.abs(avg) >= 100000 || Math.abs(max) >= 100000) {
      state = "danger";
      message = `Possible overcurrent/range limit. Live ${formatChamber(avg)}; reduce signal and recover electrometer state.`;
    } else if (span > Math.max(0.05, Math.abs(avg) * 0.5)) {
      state = "warning";
      message = `Unstable signal. Live ${formatChamber(avg)}, span ${formatChamber(span)}.`;
    } else if (isFiniteNumber(normalizedPercent) && Math.abs(normalizedPercent) < 1) {
      state = "warning";
      message = `Undercurrent/no meaningful live signal versus normalization. Live ${formatChamber(avg)}, normalized ${formatTrimmed(normalizedPercent, 3)}%.`;
    } else if (Math.abs(avg) < 0.001) {
      state = "warning";
      message = `Near-zero live current. Live ${formatChamber(avg)}; likely undercurrent/no beam.`;
    } else if (isFiniteNumber(backgroundValue) && Math.abs(corrected) < Math.max(0.000001, Math.abs(backgroundValue) * 0.05)) {
      state = "warning";
      message = `Near background. Live ${formatChamber(avg)}, background ${formatChamber(backgroundValue)}.`;
    } else if (isFiniteNumber(normalizedPercent)) {
      message = `Usable-looking live signal. Live ${formatChamber(avg)}, normalized ${formatTrimmed(normalizedPercent, 2)}%.`;
    }

    elements.signalText.textContent = message;
    elements.signalText.className = `signal-state ${state}`;
  }

  function renderControllerErrors() {
    if (!elements.controllerErrorText) {
      return;
    }

    const errors = snapshot?.controllerErrors || [];
    if (!errors.length) {
      elements.controllerErrorText.textContent = "No controller errors recorded";
      elements.controllerErrorText.className = "signal-state usable";
      return;
    }

    const latest = errors.slice(-4).map((error) =>
      `${formatTime(error.timestamp, true)} ${error.codeHex} ${error.message}`
    );
    elements.controllerErrorText.textContent = latest.join(" | ");
    elements.controllerErrorText.className = "signal-state danger";
  }

  async function loadOpabQueue() {
    setPending("Loading OPAB queue");
    try {
      const path = elements.opabPathInput.value.trim();
      if (!path) {
        openOpabFilePicker();
        return;
      }

      await loadOpabQueueFromPath(path);
    } catch (error) {
      openOpabFilePicker(`Could not load that OPAB path: ${error.message}`);
    }
  }

  async function loadOpabQueueFromPath(path) {
    opabQueue = await api(`/opab/queue?path=${encodeURIComponent(path)}`);
    elements.opabPathInput.value = opabQueue.sourceFile || path;
    await loadLatestQueueRunForOpab(opabQueue.sourceFile);
    pendingMessage = "";
    render();
  }

  function openOpabFilePicker(reason) {
    if (!elements.opabFileInput) {
      pendingMessage = reason ? `Error: ${reason}` : "Error: OPAB file picker is not available.";
      render();
      return;
    }

    pendingMessage = reason || "Choose an OPAB file.";
    render();
    elements.opabFileInput.value = "";
    elements.opabFileInput.click();
  }

  async function importSelectedOpabFile(file) {
    if (!file) {
      pendingMessage = "";
      render();
      return;
    }

    setPending("Importing OPAB file");
    try {
      opabQueue = await api("/opab/import", {
        method: "POST",
        body: JSON.stringify({
          fileName: file.name,
          contentBase64: await fileToBase64(file)
        })
      });
      elements.opabPathInput.value = opabQueue.sourceFile;
      await loadLatestQueueRunForOpab(opabQueue.sourceFile);
      pendingMessage = "";
      render();
    } catch (error) {
      pendingMessage = `Error: ${error.message}`;
      render();
    }
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.addEventListener("load", () => {
        const result = String(reader.result || "");
        resolve(result.includes(",") ? result.slice(result.indexOf(",") + 1) : result);
      });
      reader.addEventListener("error", () => reject(reader.error || new Error("Could not read selected file.")));
      reader.readAsDataURL(file);
    });
  }

  async function loadLatestQueueRunForOpab(sourceFile) {
    activeQueueRun = null;
    activeQueueRunId = null;
    selectedScanResult = null;
    selectedScanJobId = null;
    clearLiveScan();
    try {
      const runs = await api("/queue-runs");
      const match = (runs || []).find((run) => samePath(run.sourceOpabPath, sourceFile));
      if (match) {
        activeQueueRun = match;
        activeQueueRunId = ["completed", "aborted", "failed"].includes(match.status) ? null : match.queueRunId;
      }
    } catch {
      activeQueueRun = null;
      activeQueueRunId = null;
    }
  }

  function samePath(a, b) {
    return String(a || "").toLowerCase().replaceAll("/", "\\") === String(b || "").toLowerCase().replaceAll("/", "\\");
  }

  function renderQueue() {
    if (!opabQueue?.items?.length) {
      elements.queueState.textContent = activeQueueRun ? formatQueueRunStatus(activeQueueRun) : "No queue loaded";
      elements.queueState.className = "badge neutral";
      elements.queueItemsBody.innerHTML = '<tr><td colspan="10">Load an OPAB to inspect queue items.</td></tr>';
      renderQueueCheckAllState();
      return;
    }

    elements.queueState.textContent = activeQueueRun ? formatQueueRunStatus(activeQueueRun) : `${opabQueue.items.length} item(s) loaded`;
    elements.queueState.className = `badge ${activeQueueRun?.status === "running" ? "warning" : activeQueueRun?.status === "completed" ? "good" : "neutral"}`;
    const existingRows = new Map([...elements.queueItemsBody.querySelectorAll("tr[data-queue-index]")].map((row) => [Number(row.dataset.queueIndex), {
      selected: row.querySelector(".queue-select")?.checked,
      travel: row.querySelector(".queue-travel")?.value,
      speed: row.querySelector(".queue-speed")?.value,
      step: row.querySelector(".queue-step")?.value
    }]));
    const queueRunByIndex = new Map((activeQueueRun?.items || []).map((item) => [item.queueItemIndex, item]));
    elements.queueItemsBody.innerHTML = opabQueue.items.map((item) => {
      const runItem = queueRunByIndex.get(item.index);
      const previous = existingRows.get(item.index);
      const isSelected = previous?.selected ?? (runItem ? runItem.status !== "disabled" : item.index <= 2);
      const depth = item.scanType === "Beam"
        ? `${formatOptional(item.startPositionDepth)} -> ${formatOptional(item.endPositionDepth)}`
        : formatOptional(item.endPositionDepth);
      return `
        <tr data-queue-index="${item.index}" data-scan-job-id="${escapeHtml(runItem?.scanJobId || "")}" class="${runItem?.scanJobId ? "has-result-link" : ""}">
          <td><input class="queue-select" type="checkbox" ${isSelected ? "checked" : ""}></td>
          <td>${item.index}</td>
          <td>${escapeHtml(item.scanType || "-")}</td>
          <td>${escapeHtml(item.templateContinuous || "-")}</td>
          <td>${formatOptional(item.fieldSizeCrossline)} / ${formatOptional(item.fieldSizeInline)}</td>
          <td>${depth}</td>
          <td><input class="queue-travel" type="number" step="0.1" value="${escapeHtml(previous?.travel ?? formatInput(item.scanTravelDistance))}"></td>
          <td><input class="queue-speed" type="number" step="0.1" value="${escapeHtml(previous?.speed ?? formatInput(item.scanSpeedContinuous))}"></td>
          <td><input class="queue-step" type="number" step="0.1" value="${escapeHtml(previous?.step ?? formatInput(item.outputStepWidth))}"></td>
          <td>${escapeHtml(runItem ? `${runItem.status}${runItem.scanJobId ? ` (${runItem.scanJobId})` : ""}` : "not run")}</td>
        </tr>`;
    }).join("");
    renderQueueCheckAllState();
  }

  function renderQueueCheckAllState() {
    if (!elements.queueCheckAllInput) {
      return;
    }

    const busy = !!snapshot?.busy || isCommandPending();
    const rows = [...elements.queueItemsBody.querySelectorAll("tr[data-queue-index]")];
    const selectedCount = rows.filter((row) => row.querySelector(".queue-select")?.checked).length;
    elements.queueCheckAllInput.disabled = !rows.length || busy;
    elements.queueCheckAllInput.checked = rows.length > 0 && selectedCount === rows.length;
    elements.queueCheckAllInput.indeterminate = selectedCount > 0 && selectedCount < rows.length;
  }

  async function openQueueItemResult(row) {
    const jobId = row?.dataset?.scanJobId;
    if (!jobId) {
      return;
    }

    setPending(`Loading ${jobId}`);
    try {
      const status = await api(`/scan-jobs/${jobId}`);
      if (!isTerminalScanStatus(status.status)) {
        await loadLiveScan(jobId);
        pendingMessage = "";
        render();
        return;
      }

      selectedScanResult = await api(`/scan-jobs/${jobId}/result`);
      selectedScanJobId = jobId;
      if (liveScanJobId === jobId) {
        clearLiveScan(false);
      }
      pendingMessage = "";
      render();
    } catch (error) {
      pendingMessage = `Error: ${error.message}`;
      render();
    }
  }

  function formatQueueRunStatus(queueRun) {
    const active = queueRun.activeItemIndex ? `, item ${queueRun.activeItemIndex}` : "";
    return `${queueRun.status}${active}`;
  }

  async function runSelectedQueue() {
    if (!opabQueue?.items?.length) {
      return;
    }

    const selected = [];
    const overrides = [];
    elements.queueItemsBody.querySelectorAll("tr[data-queue-index]").forEach((row) => {
      const index = Number(row.dataset.queueIndex);
      const enabled = row.querySelector(".queue-select").checked;
      if (enabled) {
        selected.push(index);
      }
      overrides.push({
        queueItemIndex: index,
        enabled,
        scanTravelDistance: readNumber(row.querySelector(".queue-travel")),
        scanSpeedContinuous: readNumber(row.querySelector(".queue-speed")),
        outputStepWidth: readNumber(row.querySelector(".queue-step"))
      });
    });

    if (!selected.length) {
      pendingMessage = "Error: Select at least one queue item.";
      render();
      return;
    }

    setPending("Starting queue run");
    try {
      clearLiveScan();
      selectedScanResult = null;
      selectedScanJobId = null;
      activeQueueRun = await api("/queue-runs", {
        method: "POST",
        body: JSON.stringify({
          opabPath: elements.opabPathInput.value.trim(),
          queueItemIndexes: selected,
          overrides,
          releaseWhenComplete: false,
          connectIfNeeded: true,
          requiresOperatorConfirmation: false,
          allowTankMotion: true,
          dryRun: false,
          allowNoBeamCommissioning: elements.queueNoBeamInput?.checked ?? false,
          backgroundMode: "latest",
          normalizationMode: "latest",
          detectorMode: readDetectorMode(),
          coordinateBasis: {
            centerCrosslineMm: readNumber(elements.queueCenterXInput),
            centerInlineMm: readNumber(elements.queueCenterYInput),
            surfaceDepthMm: readNumber(elements.queueSurfaceZInput),
            depthTransform: "surfaceMinusDepth",
            crosslineScanInlineMm: readNumber(elements.queueCrosslineYInput)
          }
        })
      });
      activeQueueRunId = activeQueueRun.queueRunId;
      pendingMessage = "";
      render();
    } catch (error) {
      pendingMessage = `Error: ${error.message}`;
      render();
    }
  }

  async function pollQueueRun() {
    try {
      activeQueueRun = await api(`/queue-runs/${activeQueueRunId}`);
      if (["completed", "aborted", "failed"].includes(activeQueueRun.status)) {
        activeQueueRunId = null;
      }
    } catch {
      activeQueueRunId = null;
    }
  }

  async function pollLiveScan() {
    const activeJobId = currentActiveScanJobId();
    if (!activeJobId) {
      if (liveScanJobId && selectedScanJobId === liveScanJobId && liveScanResult && !isTerminalScanStatus(liveScanResult.status)) {
        await finalizeLiveScan(liveScanJobId);
      }
      return;
    }

    try {
      await loadLiveScan(activeJobId);
      if (isTerminalScanStatus(liveScanResult?.status)) {
        await finalizeLiveScan(activeJobId);
      }
    } catch (error) {
      if (liveScanJobId === activeJobId) {
        liveScanResult = liveScanResult || createEmptyLiveResult(activeJobId, "Live scan unavailable");
        liveScanResult.warnings = [...(liveScanResult.warnings || []), error.message];
      }
    }
  }

  function currentActiveScanJobId() {
    const activeItem = (activeQueueRun?.items || []).find((item) =>
      item.queueItemIndex === activeQueueRun?.activeItemIndex && item.scanJobId);
    if (activeItem?.scanJobId && !isTerminalScanStatus(activeItem.status)) {
      return activeItem.scanJobId;
    }

    const runningItem = (activeQueueRun?.items || []).find((item) =>
      item.scanJobId && !isTerminalScanStatus(item.status) && !["disabled", "notRun", "not run"].includes(item.status));
    return runningItem?.scanJobId || null;
  }

  async function loadLiveScan(jobId) {
    const live = await api(`/scan-jobs/${jobId}/live`);
    liveScanJobId = jobId;
    liveScanResult = liveSnapshotToResult(live);
    selectedScanJobId = jobId;
    selectedScanResult = liveScanResult;
  }

  async function finalizeLiveScan(jobId) {
    try {
      selectedScanResult = await api(`/scan-jobs/${jobId}/result`);
      selectedScanJobId = jobId;
      clearLiveScan(false);
    } catch {
      if (liveScanJobId === jobId && liveScanResult) {
        selectedScanResult = liveScanResult;
      }
    }
  }

  function clearLiveScan(clearSelection = true) {
    liveScanJobId = null;
    liveScanResult = null;
    if (clearSelection) {
      selectedScanResult = null;
      selectedScanJobId = null;
    }
  }

  function liveSnapshotToResult(live) {
    return {
      schemaVersion: "measurement-result-v1-live",
      jobId: live.jobId,
      clientJobId: live.clientJobId,
      status: live.status,
      phase: live.phase,
      dryRun: false,
      measurementDateTime: live.createdAt || live.updatedAt || new Date().toISOString(),
      sourceQueue: live.sourceQueue,
      machine: live.machine,
      measurementSystem: live.measurementSystem,
      scan: live.scan,
      background: live.background,
      normalization: live.normalization,
      points: live.points || [],
      summary: live.summary || {
        pointCount: live.pointsCollected || 0,
        plannedPointCount: live.estimatedTotalPoints || 0
      },
      integrity: null,
      artifacts: [],
      warnings: live.warnings || [],
      live: true,
      updatedAt: live.updatedAt,
      currentPositionMm: live.currentPositionMm,
      pointsCollected: live.pointsCollected || 0,
      estimatedTotalPoints: live.estimatedTotalPoints || 0
    };
  }

  function createEmptyLiveResult(jobId, message) {
    return {
      schemaVersion: "measurement-result-v1-live",
      jobId,
      status: "live",
      phase: "live",
      points: [],
      summary: { pointCount: 0, plannedPointCount: 0 },
      warnings: message ? [message] : [],
      artifacts: [],
      live: true
    };
  }

  function isTerminalScanStatus(status) {
    return ["completed", "aborted", "failed", "partial", "completedNoPoints"].includes(status || "");
  }

  async function abortQueueRun() {
    if (!activeQueueRunId) {
      return;
    }

    setPending("Aborting queue");
    try {
      activeQueueRun = await api(`/queue-runs/${activeQueueRunId}/abort`, { method: "POST" });
      activeQueueRunId = null;
      pendingMessage = "";
      render();
    } catch (error) {
      pendingMessage = `Error: ${error.message}`;
      render();
    }
  }

  function renderSelectedScanResult() {
    const canvas = elements.scanResultCanvas;
    if (!canvas) {
      return;
    }

    const result = selectedScanResult;
    if (!result) {
      elements.scanResultState.textContent = "Click a completed queue item";
      elements.scanResultState.className = "badge neutral";
      elements.scanResultSummary.textContent = "No scan selected.";
      elements.scanResultJsonLink.hidden = true;
      elements.scanResultAscLink.hidden = true;
      drawEmptyScanResult("No scan selected.");
      return;
    }

    const points = result.points || [];
    const scan = result.scan || {};
    const axis = resolveResultAxis(scan);
    const plottedPoints = selectPlannedScanSegment(points, scan);
    const isLive = result.live === true && !isTerminalScanStatus(result.status);
    const plannedPoints = result.summary?.plannedPointCount ?? result.estimatedTotalPoints ?? "-";
    elements.scanResultState.textContent = isLive
      ? `Live ${selectedScanJobId || result.jobId}: ${result.phase || result.status}`
      : `${selectedScanJobId || result.jobId}: ${result.status}`;
    elements.scanResultState.className = `badge ${result.status === "completed" ? "good" : isLive ? "warning" : "warning"}`;
    elements.scanResultSummary.textContent = [
      scan.scanType || "Scan",
      plottedPoints.length === points.length ? `${points.length} point(s)` : `${plottedPoints.length} plotted / ${points.length} point(s)`,
      `planned ${plannedPoints}`,
      `field ${formatOptional(scan.fieldSizeMm?.crossline)} / ${formatOptional(scan.fieldSizeMm?.inline)} mm`,
      `axis ${axis.label} relative to ${axis.reference}`
    ].join(" | ");
    const hasAsc = (result.artifacts || []).some((artifact) => artifact.type === "asc");
    const hasJson = (result.artifacts || []).some((artifact) => artifact.type === "measurement-json");
    elements.scanResultJsonLink.href = `/scan-jobs/${result.jobId}/artifacts/measurement-result.json`;
    elements.scanResultJsonLink.hidden = isLive || !hasJson;
    elements.scanResultAscLink.href = `/scan-jobs/${result.jobId}/artifacts/result.asc`;
    elements.scanResultAscLink.hidden = isLive || !hasAsc;
    drawScanResult(result, axis);
  }

  function resolveResultAxis(scan) {
    if ((scan.scanType || "").toLowerCase() === "depthdose") {
      return { key: "depth", label: "Depth", reference: "surface" };
    }

    const start = scan.startPositionMm || {};
    const end = scan.endPositionMm || {};
    const inlineTravel = Math.abs((end.inline ?? 0) - (start.inline ?? 0));
    const crosslineTravel = Math.abs((end.crossline ?? 0) - (start.crossline ?? 0));
    return inlineTravel > crosslineTravel
      ? { key: "inline", label: "Inline", reference: "iso" }
      : { key: "crossline", label: "Crossline", reference: "iso" };
  }

  function drawEmptyScanResult(message) {
    const canvas = elements.scanResultCanvas;
    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.round(rect.width * dpr));
    canvas.height = Math.max(1, Math.round(rect.height * dpr));
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, rect.width, rect.height);
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, rect.width, rect.height);
    ctx.fillStyle = "#5e6c72";
    ctx.font = "13px Segoe UI, Arial, sans-serif";
    ctx.fillText(message, 16, 28);
  }

  function drawScanResult(result, axis) {
    const canvas = elements.scanResultCanvas;
    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.round(rect.width * dpr));
    canvas.height = Math.max(1, Math.round(rect.height * dpr));
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, rect.width, rect.height);
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, rect.width, rect.height);

    const rawPoints = buildProfilePoints(result, axis);

    const plot = { left: 70, top: 26, right: rect.width - 28, bottom: rect.height - 44 };
    plot.width = Math.max(10, plot.right - plot.left);
    plot.height = Math.max(10, plot.bottom - plot.top);
    drawRect(ctx, plot, "#5e6c72");

    if (!rawPoints.length) {
      ctx.fillStyle = "#5e6c72";
      ctx.font = "13px Segoe UI, Arial, sans-serif";
      ctx.fillText("No plotted points available in this result.", plot.left + 12, plot.top + 24);
      return;
    }

    const minX = Math.min(...rawPoints.map((point) => point.x));
    const maxX = Math.max(...rawPoints.map((point) => point.x));
    let minY = Math.min(...rawPoints.map((point) => point.y));
    let maxY = Math.max(...rawPoints.map((point) => point.y));
    const yMargin = Math.max(1, Math.abs(maxY - minY) * 0.08);
    minY -= yMargin;
    maxY += yMargin;
    if (Math.abs(maxX - minX) < 0.001) {
      rawPoints.forEach((point, index) => {
        point.x = index;
      });
    }
    if (Math.abs(maxY - minY) < 0.001) {
      minY -= 1;
      maxY += 1;
    }

    const actualMinX = Math.min(...rawPoints.map((point) => point.x));
    const actualMaxX = Math.max(...rawPoints.map((point) => point.x));
    const xScale = (value) => plot.left + ((value - actualMinX) / Math.max(0.001, actualMaxX - actualMinX)) * plot.width;
    const yScale = (value) => plot.bottom - ((value - minY) / Math.max(0.001, maxY - minY)) * plot.height;

    ctx.strokeStyle = "#e4e9eb";
    ctx.lineWidth = 1;
    ctx.fillStyle = "#5e6c72";
    ctx.font = "12px Segoe UI, Arial, sans-serif";
    for (let tick = 0; tick <= 5; tick += 1) {
      const y = plot.top + (plot.height * tick) / 5;
      drawLine(ctx, plot.left, y, plot.right, y);
      const value = maxY - ((maxY - minY) * tick) / 5;
      drawRightText(ctx, formatTrimmed(value, 2), plot.left - 8, y + 4);
    }
    for (let tick = 0; tick <= 5; tick += 1) {
      const x = plot.left + (plot.width * tick) / 5;
      drawLine(ctx, x, plot.top, x, plot.bottom);
      const value = actualMinX + ((actualMaxX - actualMinX) * tick) / 5;
      const label = formatTrimmed(value, 1);
      ctx.fillText(label, x - ctx.measureText(label).width / 2, plot.bottom + 24);
    }

    drawRect(ctx, plot, "#5e6c72");
    drawSeries(ctx, rawPoints.map((point) => [xScale(point.x), yScale(point.y)]), "#087685");
    ctx.fillStyle = "#1f292d";
    ctx.font = "700 13px Segoe UI, Arial, sans-serif";
    ctx.fillText(`${axis.label} mm relative to ${axis.reference}`, plot.left, rect.height - 10);
    ctx.save();
    ctx.translate(16, plot.top + plot.height / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillText("Dose / ratio", 0, 0);
    ctx.restore();
  }

  function clinicalAxisValue(position, key) {
    if (!position) {
      return null;
    }

    const centerCrossline = readNumber(elements.queueCenterXInput) ?? 0;
    const centerInline = readNumber(elements.queueCenterYInput) ?? 0;
    const surfaceDepth = readNumber(elements.queueSurfaceZInput) ?? 0;
    if (key === "depth") {
      return surfaceDepth - position.depth;
    }

    if (key === "inline") {
      return position.inline - centerInline;
    }

    return position.crossline - centerCrossline;
  }

  function clinicalAxisValueForResult(result, position, key) {
    if (!position) {
      return null;
    }

    const basis = result?.sourceQueue?.coordinateBasis;
    if (!basis) {
      return clinicalAxisValue(position, key);
    }

    if (key === "depth") {
      const surfaceDepth = numberOrNull(basis.surfaceDepthMm);
      return isFiniteNumber(surfaceDepth) ? surfaceDepth - position.depth : clinicalAxisValue(position, key);
    }

    if (key === "inline") {
      const centerInline = numberOrNull(basis.centerInlineMm);
      return isFiniteNumber(centerInline) ? position.inline - centerInline : clinicalAxisValue(position, key);
    }

    const centerCrossline = numberOrNull(basis.centerCrosslineMm);
    return isFiniteNumber(centerCrossline) ? position.crossline - centerCrossline : clinicalAxisValue(position, key);
  }

  function buildProfilePoints(result, axis) {
    return selectPlannedScanSegment(result?.points || [], result?.scan || {})
      .map((point) => ({
        x: clinicalAxisValueForResult(result, point.positionMm, axis.key),
        y: profileDoseValue(point)
      }))
      .filter((point) => isFiniteNumber(point.x) && isFiniteNumber(point.y));
  }

  function profileDoseValue(point) {
    if (!point) {
      return null;
    }

    if (isFiniteNumber(point.ratioValue)) {
      return point.ratioValue;
    }

    if (isFiniteNumber(point.currentFieldValue)
      && isFiniteNumber(point.currentReferenceValue)
      && Math.abs(point.currentReferenceValue) > 0.000000001) {
      return point.currentFieldValue / point.currentReferenceValue * 100;
    }

    return point.scaledFieldValue ?? point.currentFieldValue;
  }

  function analyzeProfileCenter(result, axis) {
    if (!result) {
      return { valid: false, message: "No scan selected." };
    }

    if (!axis || axis.key === "depth") {
      return { valid: false, message: "Centering uses a crossline or inline profile, not a PDD." };
    }

    const profile = buildProfilePoints(result, axis)
      .sort((a, b) => a.x - b.x);
    if (profile.length < 8) {
      return { valid: false, message: "Not enough plotted profile points for 50% centering." };
    }

    const signalProfile = profile.map((point) => ({
      x: point.x,
      y: Math.abs(point.y)
    }));
    const values = signalProfile.map((point) => point.y).sort((a, b) => a - b);
    const low = percentile(values, 0.1);
    const high = percentile(values, 0.95);
    if (!isFiniteNumber(low) || !isFiniteNumber(high) || high <= low) {
      return { valid: false, message: "Profile signal is not separated enough for 50% centering." };
    }

    const threshold = low + (high - low) * 0.5;
    const bestRun = findBestBoundedThresholdRun(signalProfile, threshold, true)
      ?? findBestBoundedThresholdRun(signalProfile, threshold, false);
    if (!bestRun) {
      return { valid: false, message: "The profile did not capture both 50% beam edges. Run a wider centering profile." };
    }

    const leftEdge = interpolateThresholdX(signalProfile[bestRun.start - 1], signalProfile[bestRun.start], threshold);
    const rightEdge = interpolateThresholdX(signalProfile[bestRun.end], signalProfile[bestRun.end + 1], threshold);
    if (!isFiniteNumber(leftEdge) || !isFiniteNumber(rightEdge) || rightEdge <= leftEdge) {
      return { valid: false, message: "Could not interpolate both 50% beam edges." };
    }

    const centerOffset = (leftEdge + rightEdge) / 2;
    const currentCenter = resultCenterCoordinate(result, axis.key);
    return {
      valid: true,
      leftEdgeMm: leftEdge,
      rightEdgeMm: rightEdge,
      thresholdValue: threshold,
      centerOffsetMm: centerOffset,
      fieldWidthMm: rightEdge - leftEdge,
      currentCenterMm: currentCenter,
      adjustedCenterMm: currentCenter + centerOffset
    };
  }

  function findBestBoundedThresholdRun(profile, threshold, aboveThreshold) {
    let bestRun = null;
    let runStart = null;
    for (let index = 0; index < profile.length; index += 1) {
      const inRun = aboveThreshold ? profile[index].y >= threshold : profile[index].y < threshold;
      if (inRun && runStart === null) {
        runStart = index;
      }
      if ((!inRun || index === profile.length - 1) && runStart !== null) {
        const runEnd = inRun && index === profile.length - 1 ? index : index - 1;
        const isBounded = runStart > 0 && runEnd < profile.length - 1;
        if (isBounded && (!bestRun || runEnd - runStart > bestRun.end - bestRun.start)) {
          bestRun = { start: runStart, end: runEnd, aboveThreshold };
        }
        runStart = null;
      }
    }

    return bestRun;
  }

  function applySelectedScanCentering() {
    const result = selectedScanResult;
    const axis = resolveResultAxis(result?.scan || {});
    const centering = analyzeProfileCenter(result, axis);
    if (!centering.valid) {
      pendingMessage = `Error: ${centering.message}`;
      render();
      return;
    }

    const input = axis.key === "inline" ? elements.queueCenterYInput : elements.queueCenterXInput;
    input.value = centering.adjustedCenterMm.toFixed(3);
    pendingMessage = `Applied 50% center: ${axis.label} shifted ${formatTrimmed(centering.centerOffsetMm, 2)} mm; next ${axis.key === "inline" ? "Center Y" : "Center X"} ${formatTrimmed(centering.adjustedCenterMm, 3)} mm.`;
    render();
  }

  function resultCenterCoordinate(result, key) {
    const basis = result?.sourceQueue?.coordinateBasis;
    if (key === "inline") {
      return numberOrNull(basis?.centerInlineMm) ?? readNumber(elements.queueCenterYInput) ?? 0;
    }

    return numberOrNull(basis?.centerCrosslineMm) ?? readNumber(elements.queueCenterXInput) ?? 0;
  }

  function interpolateThresholdX(left, right, threshold) {
    if (!left || !right) {
      return null;
    }

    const delta = right.y - left.y;
    if (Math.abs(delta) < 0.000001) {
      return (left.x + right.x) / 2;
    }

    const fraction = (threshold - left.y) / delta;
    return left.x + clamp(fraction, 0, 1) * (right.x - left.x);
  }

  function percentile(sortedValues, fraction) {
    if (!sortedValues.length) {
      return null;
    }

    const index = clamp((sortedValues.length - 1) * fraction, 0, sortedValues.length - 1);
    const lower = Math.floor(index);
    const upper = Math.ceil(index);
    const weight = index - lower;
    return sortedValues[lower] * (1 - weight) + sortedValues[upper] * weight;
  }

  function numberOrNull(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number : null;
  }

  function selectPlannedScanSegment(points, scan) {
    const start = scan?.startPositionMm;
    const end = scan?.endPositionMm;
    if (!points.length || !start || !end) {
      return points;
    }

    const plannedDistance = positionDistance(start, end);
    if (plannedDistance <= 0.001) {
      return points;
    }

    const startTolerance = clamp(plannedDistance * 0.05, 5, 20);
    const corridorTolerance = clamp(plannedDistance * 0.03, 5, 12);
    const startIndex = points.findIndex((point) => positionDistance(point.positionMm, start) <= startTolerance);
    if (startIndex < 0) {
      return points;
    }

    const filtered = points
      .slice(startIndex)
      .filter((point) => isInsidePlannedScanCorridor(point.positionMm, start, end, plannedDistance, corridorTolerance));
    return filtered.length ? filtered : points;
  }

  function isInsidePlannedScanCorridor(position, start, end, plannedDistance, tolerance) {
    if (!position) {
      return false;
    }

    const vx = (end.inline ?? 0) - (start.inline ?? 0);
    const vy = (end.crossline ?? 0) - (start.crossline ?? 0);
    const vz = (end.depth ?? 0) - (start.depth ?? 0);
    const wx = (position.inline ?? 0) - (start.inline ?? 0);
    const wy = (position.crossline ?? 0) - (start.crossline ?? 0);
    const wz = (position.depth ?? 0) - (start.depth ?? 0);
    const projection = (wx * vx + wy * vy + wz * vz) / (plannedDistance * plannedDistance);
    if (projection < -0.05 || projection > 1.05) {
      return false;
    }

    const nearest = {
      inline: (start.inline ?? 0) + projection * vx,
      crossline: (start.crossline ?? 0) + projection * vy,
      depth: (start.depth ?? 0) + projection * vz
    };
    return positionDistance(position, nearest) <= tolerance;
  }

  function positionDistance(a, b) {
    if (!a || !b) {
      return Number.POSITIVE_INFINITY;
    }

    return Math.hypot(
      (a.inline ?? 0) - (b.inline ?? 0),
      (a.crossline ?? 0) - (b.crossline ?? 0),
      (a.depth ?? 0) - (b.depth ?? 0)
    );
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  function renderLog() {
    const lines = snapshot?.logs ? [...snapshot.logs] : [];
    if (pendingMessage?.startsWith("Error") || pendingMessage?.startsWith("API offline")) {
      lines.push(`[${new Date().toLocaleTimeString("en-US", { hour12: false })}] ${pendingMessage}`);
    }
    elements.logOutput.textContent = lines.join("\n");
    elements.logOutput.scrollTop = elements.logOutput.scrollHeight;
  }

  function updateSpigotDirection() {
    if (!elements.turnAngleSelect || !elements.spigotText) {
      return;
    }

    const angle = Number(elements.turnAngleSelect.value);
    const text = {
      0: "back-left corner, pointing left",
      90: "back-right corner, pointing back",
      180: "front-right corner, pointing right",
      270: "front-left corner, pointing forward"
    }[angle] || "unknown orientation";
    elements.spigotText.textContent = text;
  }

  function formatTime(value, includeMs) {
    if (!value) {
      return "-";
    }
    const date = new Date(value);
    const base = date.toLocaleTimeString("en-US", { hour12: false });
    return includeMs ? `${base}.${String(date.getMilliseconds()).padStart(3, "0")}` : base;
  }

  function isFiniteNumber(value) {
    return typeof value === "number" && Number.isFinite(value);
  }

  function formatNumber(value, places) {
    return isFiniteNumber(value) ? value.toFixed(places) : "-";
  }

  function formatTrimmed(value, places) {
    return isFiniteNumber(value) ? Number(value.toFixed(places)).toString() : "-";
  }

  function formatSigned(value, places) {
    if (!isFiniteNumber(value)) {
      return "-";
    }

    const trimmed = formatTrimmed(value, places);
    return value > 0 ? `+${trimmed}` : trimmed;
  }

  function formatVolts(value) {
    return isFiniteNumber(value) ? `${formatTrimmed(value, 3)} V` : "-";
  }

  function formatChamber(value) {
    return isFiniteNumber(value) ? `${formatTrimmed(value, 5)} x10^3 pA` : "-";
  }

  function formatLandmark(landmark) {
    if (!landmark) {
      return "-";
    }
    return `X ${formatTrimmed(landmark.x, 3)}, Y ${formatTrimmed(landmark.y, 3)}, Z ${formatTrimmed(landmark.z, 3)}`;
  }

  function formatPreparation(preparation) {
    if (!preparation) {
      return "Not prepared";
    }
    const ratio = isFiniteNumber(preparation.normalizationRatio)
      ? `, ratio ${formatTrimmed(preparation.normalizationRatio, 5)}`
      : "";
    const target = preparation.normalizationTargetMm
      ? `, target X ${formatTrimmed(preparation.normalizationTargetMm.crossline, 3)}, Y ${formatTrimmed(preparation.normalizationTargetMm.inline, 3)}, Z ${formatTrimmed(preparation.normalizationTargetMm.depth, 3)}`
      : "";
    return `${preparation.status}: ${preparation.message}${ratio}${target}`;
  }

  function formatOptional(value) {
    return isFiniteNumber(value) ? formatTrimmed(value, 2) : "-";
  }

  function formatInput(value) {
    return isFiniteNumber(value) ? Number(value.toFixed(3)).toString() : "";
  }

  function readNumber(input) {
    const value = Number(input.value);
    return Number.isFinite(value) ? value : null;
  }

  function readDetectorMode() {
    return elements.detectorModeSelect?.value === "fieldReference" ? "fieldReference" : "fieldOnly";
  }

  async function maybePromptForAcquisitionRecovery() {
    const cache = snapshot?.recentAcquisitionCache;
    if (!recoveryStartupCheckArmed || recoveryPromptOpen || isCommandPending()) {
      return;
    }

    if (!snapshot?.connected) {
      return;
    }

    if (!cache) {
      recoveryStartupCheckArmed = false;
      return;
    }

    const detectorMode = readDetectorMode();
    const promptKey = `${cache.cacheKey || ""}:${detectorMode}`;
    if (!promptKey || window.localStorage.getItem("tank.acquisitionRecoveryPrompted") === promptKey) {
      recoveryStartupCheckArmed = false;
      return;
    }

    const modeMatches = cache.detectorMode === detectorMode;
    const canRestoreAcquisition = modeMatches && (cache.hasUsableBackground || cache.hasUsableNormalization);
    const canRestoreCentering = cache.hasUsableCentering === true;
    if (!canRestoreAcquisition && !canRestoreCentering) {
      recoveryStartupCheckArmed = false;
      return;
    }

    recoveryPromptOpen = true;
    try {
      const lines = [
        "Recent setup values were found in the bridge temp recovery cache.",
        "",
        `Age: about ${formatTrimmed(numberOrNull(cache.ageMinutes) ?? 0, 0)} minute(s).`
      ];
      if (canRestoreAcquisition) {
        lines.push(`Detector mode: ${detectorModeLabel(detectorMode)}.`);
      }
      if (modeMatches && cache.hasUsableBackground) {
        lines.push(`Background: ${formatCacheTime(cache.backgroundCompletedAt)}.`);
      }
      if (modeMatches && cache.hasUsableNormalization) {
        lines.push(`Normalization: ${formatCacheTime(cache.normalizationCompletedAt)}.`);
      }
      if (canRestoreCentering) {
        const centeredX = numberOrNull(cache.centering?.adjustedCenterMm);
        lines.push(isFiniteNumber(centeredX)
          ? `Centering: ${formatCacheTime(cache.centeringCompletedAt)}, X ${formatTrimmed(centeredX, 3)} mm.`
          : `Centering: ${formatCacheTime(cache.centeringCompletedAt)}.`);
      }
      lines.push("", "Restore these values into the preparation tree?");

      window.localStorage.setItem("tank.acquisitionRecoveryPrompted", promptKey);
      recoveryStartupCheckArmed = false;
      if (!window.confirm(lines.join("\n"))) {
        return;
      }

      setPending("Restoring recent setup values");
      snapshot = await api("/api/acquisition-cache/restore", {
        method: "POST",
        body: JSON.stringify({ detectorMode })
      });
      pendingMessage = "";
      restoreCenteringFromSnapshot();
      render();
    } catch (error) {
      pendingMessage = `Error: ${error.message}`;
      render();
    } finally {
      recoveryPromptOpen = false;
    }
  }

  function formatCacheTime(value) {
    if (!value) {
      return "time unknown";
    }

    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "time unknown";
    }

    return date.toLocaleTimeString("en-US", { hour12: false });
  }

  function restoreUserPreferences() {
    const detectorMode = window.localStorage.getItem("tank.detectorMode");
    if (elements.detectorModeSelect && (detectorMode === "fieldOnly" || detectorMode === "fieldReference")) {
      elements.detectorModeSelect.value = detectorMode;
    }
  }

  function storeUserPreferences() {
    window.localStorage.setItem("tank.detectorMode", readDetectorMode());
  }

  function detectorModeLabel(mode) {
    return readiness.detectorModeLabel(mode);
  }

  function hasCurrentModeBackground() {
    return readiness.hasCurrentModeBackground(snapshot, readDetectorMode());
  }

  function hasCurrentModeNormalization() {
    return readiness.hasCurrentModeNormalization(snapshot, readDetectorMode());
  }

  function centeringReadinessIssue() {
    return readiness.centeringReadinessIssue(snapshot, readDetectorMode());
  }

  function readSensitivitySettings() {
    return {
      field: elements.fieldSensitivitySelect?.value || "normal",
      reference: elements.referenceSensitivitySelect?.value || "normal",
      detectorMode: readDetectorMode()
    };
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (char) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;"
    }[char]));
  }

  async function prepareToScan() {
    const missing = snapshot?.coordinateState?.missing || [];
    if (missing.length) {
      pendingMessage = `Error: ${missingCoordinateMessage(missing)} before preparing to scan.`;
      render();
      return false;
    }

    const detectorMode = readDetectorMode();
    setPending(`Prepare ${detectorModeLabel(detectorMode)}: background`);
    try {
      snapshot = await api("/api/prepare-to-scan/start", {
        method: "POST",
        body: JSON.stringify({ detectorMode })
      });
      pendingMessage = "";
      render();

      if (snapshot?.preparation?.status !== "backgroundMeasured") {
        return false;
      }

      const continueNormalization = window.confirm("Turn beam on for normalization, then press OK.");
      if (!continueNormalization) {
        pendingMessage = "Background complete. Turn beam on, then press Normalize when ready.";
        render();
        return false;
      }

      setPending(`Prepare ${detectorModeLabel(detectorMode)}: normalize`);
      snapshot = await api("/api/prepare-to-scan/normalize", {
        method: "POST",
        body: JSON.stringify({ detectorMode })
      });
      pendingMessage = "";
      render();
      return snapshot?.preparation?.status === "readyToScan";
    } catch (error) {
      pendingMessage = `Error: ${error.message}`;
      await pollState();
      return false;
    }
  }

  async function takeBackground() {
    if (!snapshot?.connected) {
      pendingMessage = "Error: Connect before taking background.";
      render();
      return;
    }

    if (!biasWorkflowComplete) {
      pendingMessage = "Error: Set Bias/HV first. For diode-safe work, choose 0 V; no HV command will be sent.";
      render();
      return;
    }

    const detectorMode = readDetectorMode();
    setCenteringWorkflowComplete(false);
    const ok = await postTimedAcquisition(
      "/api/background",
      "background",
      `Taking background for ${detectorModeLabel(detectorMode)}`,
      detectorMode
    );

    if (ok && !hasCurrentModeBackground()) {
      pendingMessage = `Error: Background finished, but ${detectorModeLabel(detectorMode)} readings were not valid. Check detector mode and chamber signal, then retry.`;
      render();
    }
  }

  async function takeNormalization() {
    if (!hasCurrentModeBackground()) {
      pendingMessage = `Error: Take ${detectorModeLabel(readDetectorMode())} background first with beam off.`;
      render();
      return;
    }

    const positioned = await ensureNormalizationPosition();
    if (!positioned) {
      return;
    }

    const detectorMode = readDetectorMode();
    setCenteringWorkflowComplete(false);
    const ok = await postTimedAcquisition(
      "/api/normalize",
      "normalization",
      `Taking normalization for ${detectorModeLabel(detectorMode)}`,
      detectorMode
    );

    if (ok && !hasCurrentModeNormalization()) {
      pendingMessage = `Error: Normalization finished, but ${detectorModeLabel(detectorMode)} was not valid. Confirm beam signal is above background, then retry.`;
      render();
    }
  }

  async function ensureNormalizationPosition() {
    const iso = snapshot?.isocenter;
    if (!iso || !isFiniteNumber(iso.x) || !isFiniteNumber(iso.y) || !isFiniteNumber(iso.z)) {
      pendingMessage = "Error: Set or refresh isocenter before normalization.";
      render();
      return false;
    }

    const target = normalizationTarget();
    const current = currentRelativePosition();
    if (current && isNearNormalizationTarget(current)) {
      return true;
    }

    const currentText = current
      ? `Current relative position: X ${formatSigned(current.x, 2)} mm, Y ${formatSigned(current.y, 2)} mm, Z ${formatSigned(current.z, 2)} mm.`
      : "Current live position is not available yet.";
    const go = window.confirm([
      "Move to the normalization point before taking normalization?",
      "",
      "Target: iso center with detector 1.5 cm deeper.",
      `Target relative position: X 0 mm, Y 0 mm, Z +${formatTrimmed(NORMALIZATION_DEPTH_FROM_ISO_MM, 1)} mm.`,
      currentText,
      "",
      "Press OK to move there, then take normalization."
    ].join("\n"));

    if (!go) {
      pendingMessage = "Normalization canceled before positioning move.";
      render();
      return false;
    }

    setPending(`Moving to normalization point X 0, Y 0, Z +${formatTrimmed(NORMALIZATION_DEPTH_FROM_ISO_MM, 1)} mm`);
    try {
      snapshot = await api("/api/move", {
        method: "POST",
        body: JSON.stringify({
          crossline: target.absoluteX,
          inline: target.absoluteY,
          depth: target.absoluteZ,
          speed: clamp(readTargetInputs().speed, 1, 100)
        })
      });
      await waitForNormalizationPosition();
      pendingMessage = "";
      render();
      return true;
    } catch (error) {
      pendingMessage = `Error: ${error.message}`;
      await pollState();
      return false;
    }
  }

  function normalizationTarget() {
    const iso = snapshot?.isocenter;
    return {
      relativeX: 0,
      relativeY: 0,
      relativeZ: NORMALIZATION_DEPTH_FROM_ISO_MM,
      absoluteX: iso.x,
      absoluteY: iso.y,
      absoluteZ: iso.z + NORMALIZATION_DEPTH_FROM_ISO_MM
    };
  }

  function currentRelativePosition() {
    const latest = snapshot?.latestStatus;
    const iso = snapshot?.isocenter;
    if (!latest || !iso
      || !isFiniteNumber(latest.x) || !isFiniteNumber(latest.y) || !isFiniteNumber(latest.z)
      || !isFiniteNumber(iso.x) || !isFiniteNumber(iso.y) || !isFiniteNumber(iso.z)) {
      return null;
    }

    return {
      x: latest.x - iso.x,
      y: latest.y - iso.y,
      z: latest.z - iso.z
    };
  }

  function isNearNormalizationTarget(position) {
    return Math.abs(position.x) <= NORMALIZATION_POSITION_TOLERANCE_MM
      && Math.abs(position.y) <= NORMALIZATION_POSITION_TOLERANCE_MM
      && Math.abs(position.z - NORMALIZATION_DEPTH_FROM_ISO_MM) <= NORMALIZATION_POSITION_TOLERANCE_MM;
  }

  async function waitForNormalizationPosition() {
    const started = Date.now();
    while (Date.now() - started < 45 * 1000) {
      await delay(1000);
      await pollState();
      const position = currentRelativePosition();
      if (position && isNearNormalizationTarget(position)) {
        return;
      }
    }

    throw new Error("Move to normalization point did not settle near X 0, Y 0, Z +15 mm within 45 seconds.");
  }

  async function postTimedAcquisition(path, progressKey, progressLabel, detectorMode) {
    const durationSec = DEFAULT_TIMED_ACQUISITION_SECONDS;
    startWorkflowProgress(progressKey, progressLabel, durationSec);
    pendingMessage = `${progressLabel}...`;
    renderControls();
    try {
      snapshot = await api(path, {
        method: "POST",
        body: JSON.stringify({ durationSec, detectorMode })
      });
      pendingMessage = "";
      clearWorkflowProgress();
      render();
      return true;
    } catch (error) {
      clearWorkflowProgress();
      pendingMessage = `Error: ${error.message}`;
      await pollState();
      return false;
    }
  }

  async function runCenteringScan() {
    if (!snapshot?.connected) {
      pendingMessage = "Error: Connect before running a centering scan.";
      render();
      return;
    }

    const readinessIssue = centeringReadinessIssue();
    if (readinessIssue) {
      const startPrepare = window.confirm(`${readinessIssue}\n\nStart Prepare to Scan now?`);
      if (!startPrepare) {
        pendingMessage = `Error: ${readinessIssue}`;
        render();
        return;
      }

      const prepared = await prepareToScan();
      if (!prepared) {
        pendingMessage = pendingMessage || "Error: Prepare to Scan did not complete; centering was not started.";
        render();
        return;
      }
    }

    const remainingIssue = centeringReadinessIssue();
    if (remainingIssue) {
      pendingMessage = `Error: ${remainingIssue}`;
      render();
      return;
    }

    const basis = readCenteringBasis();
    if (!basis) {
      pendingMessage = "Error: Centering requires known Center X, Center Y, and Surface Z.";
      render();
      return;
    }

    const width = clamp(readNumber(elements.centeringWidthInput) ?? 300, 20, 500);
    elements.centeringWidthInput.value = width.toFixed(0);
    const halfWidth = width / 2;
    const speed = clamp(readTargetInputs().speed, 1, 100);
    const outputStep = width >= 250 ? 1 : 0.5;
    const scanDepth = numberOrNull(snapshot?.preparation?.normalizationTargetMm?.depth)
      ?? basis.surfaceDepthMm + (numberOrNull(snapshot?.preparation?.normalizationDepthMm) ?? 15);
    const scanInline = readNumber(elements.queueCrosslineYInput) ?? basis.centerInlineMm;
    const clientJobId = `setup-centering-crossline-${compactTimestamp(new Date())}`;

    setCenteringStatus("Centering scan running", "badge warning", false);
    setPending("Running centering scan");

    try {
      const created = await api("/scan-jobs", {
        method: "POST",
        body: JSON.stringify({
          schemaVersion: "scan-job-request-v1",
          clientJobId,
          sourceQueue: {
            opabFileName: "setup-centering",
            opabFileSha256: null,
            queueItemIndex: null,
            queueItemId: "setup-centering-crossline",
            coordinateBasis: {
              centerCrosslineMm: basis.centerCrosslineMm,
              centerInlineMm: basis.centerInlineMm,
              surfaceDepthMm: basis.surfaceDepthMm,
              depthTransform: "surfaceMinusDepth"
            }
          },
          machine: {
            site: null,
            radiationDevice: "setup-centering"
          },
          measurementSystem: {
            measurementDevice: "IBA tank",
            controller: null,
            fieldDetector: "field",
            referenceDetector: readDetectorMode() === "fieldReference" ? "reference" : null,
            medium: "water",
            detectorOrientation: null
          },
          scan: {
            scanType: "Crossline",
            scanMode: "Continuous",
            radiationType: "setup",
            energy: null,
            fieldSizeMm: { inline: width, crossline: width },
            ssdMm: null,
            sadMm: null,
            gantryAngleDeg: null,
            collimatorAngleDeg: null,
            startPositionMm: {
              inline: scanInline,
              crossline: basis.centerCrosslineMm + halfWidth,
              depth: scanDepth,
              w: null
            },
            endPositionMm: {
              inline: scanInline,
              crossline: basis.centerCrosslineMm - halfWidth,
              depth: scanDepth,
              w: null
            },
            scanSpeedMmPerSec: speed,
            positioningSpeedMmPerSec: speed,
            outputStepWidthMm: outputStep,
            measurementTimeSec: null,
            ratioCalculationSkipThresholdPercent: 1
          },
          acquisition: {
            backgroundMode: "latest",
            normalizationMode: "latest",
            includeRawChannels: true,
            includeCorrectedChannels: true,
            includePositionW: false,
            detectorMode: readDetectorMode()
          },
          safety: {
            requiresOperatorConfirmation: false,
            allowTankMotion: true,
            allowBeamControl: false,
            dryRun: false,
            allowNoBeamCommissioning: elements.queueNoBeamInput?.checked ?? false
          },
          output: {
            resultFormat: "json",
            alsoWriteAsc: false
          }
        })
      });

      const staleAfterMs = Math.min(4 * 60 * 1000, Math.max(45 * 1000, (width / speed) * 2500));
      const status = await waitForScanJob(created.jobId, { staleAfterMs });
      if (!["completed", "partial"].includes(status.status)) {
        throw new Error(`Centering scan ended as ${status.status}.`);
      }

      const result = await api(`/scan-jobs/${created.jobId}/result`);
      const axis = resolveResultAxis(result.scan || {});
      const centering = analyzeProfileCenter(result, axis);
      if (!centering.valid) {
        throw new Error(centering.message);
      }

      elements.queueCenterXInput.value = centering.adjustedCenterMm.toFixed(3);
      coordinateInputsTouched = true;
      setCenteringStatus(`Centered ${formatTrimmed(centering.centerOffsetMm, 2)} mm; X ${formatTrimmed(centering.adjustedCenterMm, 3)}`, "badge good", true);
      const currentIso = snapshot?.isocenter;
      const previousIsocenter = copyLandmark(currentIso);
      const isoX = numberOrNull(currentIso?.x);
      const isoY = numberOrNull(currentIso?.y);
      const isoZ = numberOrNull(currentIso?.z);
      if (!isFiniteNumber(isoX) || !isFiniteNumber(isoY) || !isFiniteNumber(isoZ)) {
        throw new Error("Centering succeeded, but the tank isocenter Y/Z are not known. Set or refresh isocenter before writing the centered isocenter.");
      }

      render();
      const isoCorrection = centering.adjustedCenterMm - isoX;
      const correctionSign = isoCorrection >= 0 ? "+" : "";
      const writeTankIso = window.confirm([
        "Centering found the beam center.",
        "",
        `Current iso X: ${formatTrimmed(isoX, 3)} mm`,
        `New iso X: ${formatTrimmed(centering.adjustedCenterMm, 3)} mm`,
        `Isocenter correction: ${correctionSign}${formatTrimmed(isoCorrection, 3)} mm`,
        "",
        "Write this X position into the tank isocenter now? Current iso Y/Z will be preserved."
      ].join("\n"));
      if (writeTankIso) {
        setPending("Writing measured beam center to tank isocenter");
        snapshot = await api("/api/landmarks/isocenter", {
          method: "POST",
          body: JSON.stringify({
            crossline: centering.adjustedCenterMm,
            inline: isoY,
            depth: isoZ,
            w: 0
          })
        });
        setCenteringStatus(`Tank iso X set to ${formatTrimmed(centering.adjustedCenterMm, 3)}`, "badge good", true);
      }

      await recordCenteringRecovery(centering, status, isoX, previousIsocenter, writeTankIso);
      pendingMessage = "";
      await pollState();
    } catch (error) {
      setCenteringStatus("Centering failed", "badge warning", false);
      pendingMessage = `Error: ${error.message}`;
      await pollState();
    }
  }

  async function recordCenteringRecovery(centering, status, previousIsoX, previousIsocenter, wroteTankIso) {
    try {
      const adjustedCenterMm = numberOrNull(centering.adjustedCenterMm);
      const centerOffsetMm = numberOrNull(centering.centerOffsetMm);
      if (!isFiniteNumber(adjustedCenterMm) || !isFiniteNumber(centerOffsetMm)) {
        return;
      }

      const currentIsocenter = copyLandmark(snapshot?.isocenter) || previousIsocenter;
      snapshot = await api("/api/centering/result", {
        method: "POST",
        body: JSON.stringify({
          completedAt: new Date().toISOString(),
          detectorMode: readDetectorMode(),
          adjustedCenterMm,
          centerOffsetMm,
          previousIsocenter,
          currentIsocenter,
          previousIsoX: isFiniteNumber(previousIsoX) ? previousIsoX : null,
          newIsoX: wroteTankIso ? adjustedCenterMm : null,
          isoCorrectionMm: isFiniteNumber(previousIsoX) ? adjustedCenterMm - previousIsoX : null,
          wroteTankIso: wroteTankIso === true,
          pointsCollected: status?.pointsCollected ?? null,
          jobId: status?.jobId ?? null
        })
      });
      restoreCenteringFromSnapshot();
    } catch (error) {
      console.warn("Centering recovery cache was not saved:", error);
    }
  }

  function copyLandmark(landmark) {
    if (!landmark || !isFiniteNumber(landmark.x) || !isFiniteNumber(landmark.y) || !isFiniteNumber(landmark.z)) {
      return null;
    }

    return {
      x: landmark.x,
      y: landmark.y,
      z: landmark.z,
      w: landmark.w == null ? null : numberOrNull(landmark.w),
      timestamp: landmark.timestamp || new Date().toISOString()
    };
  }

  function readCenteringBasis() {
    const centerCrosslineMm = readNumber(elements.queueCenterXInput) ?? numberOrNull(snapshot?.isocenter?.x);
    const centerInlineMm = readNumber(elements.queueCenterYInput) ?? numberOrNull(snapshot?.isocenter?.y);
    const surfaceDepthMm = readNumber(elements.queueSurfaceZInput) ?? numberOrNull(snapshot?.surface?.z);
    if (!isFiniteNumber(centerCrosslineMm) || !isFiniteNumber(centerInlineMm) || !isFiniteNumber(surfaceDepthMm)) {
      return null;
    }

    return { centerCrosslineMm, centerInlineMm, surfaceDepthMm };
  }

  async function waitForScanJob(jobId, options = {}) {
    const started = Date.now();
    const staleAfterMs = options.staleAfterMs ?? 90 * 1000;
    let lastStatus = null;
    while (Date.now() - started < 6 * 60 * 1000) {
      await delay(1000);
      const status = await api(`/scan-jobs/${jobId}`);
      lastStatus = status;
      pendingMessage = `Centering scan: ${status.phase || status.status}, ${status.pointsCollected || 0} point(s)`;
      await pollState();
      if (["completed", "aborted", "failed", "partial", "completedNoPoints"].includes(status.status)) {
        return status;
      }

      if (shouldFinalizeStaleCenteringScan(status, Date.now() - started, staleAfterMs)) {
        pendingMessage = `Centering scan: collected ${status.pointsCollected} point(s), finalizing`;
        renderControls();
        return await api(`/scan-jobs/${jobId}/finalize-partial`, { method: "POST" });
      }
    }

    if (lastStatus?.pointsCollected > 0) {
      pendingMessage = `Centering scan: timed out after ${lastStatus.pointsCollected} point(s), finalizing`;
      renderControls();
      return await api(`/scan-jobs/${jobId}/finalize-partial`, { method: "POST" });
    }

    throw new Error("Centering scan timed out without collecting points.");
  }

  function shouldFinalizeStaleCenteringScan(status, elapsedMs, staleAfterMs) {
    if (status.status !== "measuring" || elapsedMs < staleAfterMs) {
      return false;
    }

    const collected = Number(status.pointsCollected);
    const estimated = Number(status.estimatedTotalPoints);
    if (!Number.isFinite(collected) || collected <= 0 || !Number.isFinite(estimated) || estimated <= 0) {
      return false;
    }

    return collected >= Math.max(estimated + 25, Math.ceil(estimated * 1.25));
  }

  function delay(milliseconds) {
    return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
  }

  function compactTimestamp(date) {
    const pad = (value, length = 2) => String(value).padStart(length, "0");
    return [
      date.getFullYear(),
      pad(date.getMonth() + 1),
      pad(date.getDate()),
      "_",
      pad(date.getHours()),
      pad(date.getMinutes()),
      pad(date.getSeconds())
    ].join("");
  }

  function drawGraph() {
    const canvas = elements.chamberCanvas;
    if (!canvas) {
      return;
    }

    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.round(rect.width * dpr));
    canvas.height = Math.max(1, Math.round(rect.height * dpr));

    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, rect.width, rect.height);
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, rect.width, rect.height);

    const plot = {
      left: 72,
      top: 30,
      right: rect.width - 72,
      bottom: rect.height - 40
    };
    plot.width = Math.max(10, plot.right - plot.left);
    plot.height = Math.max(10, plot.bottom - plot.top);

    ctx.font = "700 13px Segoe UI, Arial, sans-serif";
    ctx.fillStyle = "#1f292d";
    ctx.fillText("Live chamber readings", 10, 18);
    ctx.font = "12px Segoe UI, Arial, sans-serif";

    const fieldOnly = readDetectorMode() === "fieldOnly";
    const background = snapshot?.latestBackground;
    const normalization = snapshot?.latestNormalization;
    const backgroundMatchesMode = hasCurrentModeBackground();
    const normalizationMatchesMode = hasCurrentModeNormalization();
    const fieldBackground = backgroundMatchesMode && isFiniteNumber(background.fieldDetectorValue)
      ? background.fieldDetectorValue
      : null;
    const fieldNormalization = normalizationMatchesMode && isFiniteNumber(normalization.fieldDetectorValue)
      ? normalization.fieldDetectorValue
      : null;
    const effectiveFieldBackground = isFiniteNumber(fieldBackground)
      ? fieldBackground
      : fieldOnly && isFiniteNumber(fieldNormalization)
        ? 0
        : null;
    const referenceBackground = backgroundMatchesMode && isFiniteNumber(background.referenceDetectorValue)
      ? background.referenceDetectorValue
      : null;
    const referenceNormalization = normalizationMatchesMode && isFiniteNumber(normalization.referenceDetectorValue)
      ? normalization.referenceDetectorValue
      : null;
    const showNormalized = isFiniteNumber(effectiveFieldBackground)
      && isFiniteNumber(fieldNormalization)
      && Math.abs(fieldNormalization - effectiveFieldBackground) > 0.000001;
    const samples = (snapshot?.samples || [])
      .filter((sample) => isFiniteNumber(sample.fieldX10e3Pa) && (fieldOnly || isFiniteNumber(sample.referenceX10e3Pa)))
      .map((sample) => ({
        timestamp: new Date(sample.timestamp),
        rawField: sample.fieldX10e3Pa,
        rawReference: sample.referenceX10e3Pa,
        field: showNormalized
          ? normalizeLiveValue(sample.fieldX10e3Pa, effectiveFieldBackground, fieldNormalization)
          : sample.fieldX10e3Pa,
        reference: showNormalized && !fieldOnly
          ? normalizeLiveValue(sample.referenceX10e3Pa, referenceBackground, referenceNormalization)
          : sample.referenceX10e3Pa
      }));

    if (!samples.length) {
      drawRect(ctx, plot, "#5e6c72");
      ctx.fillText("Waiting for tank callback chamber readings...", plot.left + 12, plot.top + 22);
      return;
    }

    const latest = samples[samples.length - 1];
    const firstTime = latest.timestamp.getTime() - 120000;
    const fieldMax = Math.max(showNormalized ? 110 : 0.01, Math.max(...samples.map((sample) => sample.field)) * 1.08);
    let referenceMin = fieldOnly ? 0 : Math.min(...samples.map((sample) => sample.reference));
    let referenceMax = fieldOnly ? 1 : Math.max(...samples.map((sample) => sample.reference));
    const referenceMargin = Math.max(0.00005, Math.max(Math.abs(referenceMax - referenceMin) * 0.25, Math.abs(referenceMax) * 0.02));
    referenceMin -= referenceMargin;
    referenceMax += referenceMargin;
    if (referenceMax <= referenceMin) {
      referenceMax = referenceMin + 0.0001;
    }

    const xScale = (timestamp) => {
      const fraction = (timestamp.getTime() - firstTime) / 120000;
      return plot.left + Math.max(0, Math.min(1, fraction)) * plot.width;
    };
    const fieldY = (value) => plot.bottom - Math.max(0, Math.min(1, value / fieldMax)) * plot.height;
    const referenceY = (value) => {
      const fraction = (value - referenceMin) / (referenceMax - referenceMin);
      return plot.bottom - Math.max(0, Math.min(1, fraction)) * plot.height;
    };

    ctx.strokeStyle = "#e4e9eb";
    ctx.lineWidth = 1;
    for (let i = 0; i <= 6; i += 1) {
      const y = plot.top + (plot.height * i) / 6;
      drawLine(ctx, plot.left, y, plot.right, y);
      const fieldTick = (fieldMax * (6 - i)) / 6;
      ctx.fillStyle = "#bd3f38";
      drawRightText(ctx, showNormalized ? formatTrimmed(fieldTick, 1) : formatTrimmed(fieldTick, fieldTick >= 1 ? 2 : 3), 62, y + 4);
      if (!fieldOnly) {
        const refTick = referenceMin + ((referenceMax - referenceMin) * (6 - i)) / 6;
        ctx.fillStyle = "#286ea8";
        ctx.fillText(formatTrimmed(refTick, 5), plot.right + 8, y + 4);
      }
    }

    ctx.fillStyle = "#5e6c72";
    for (let i = 0; i <= 4; i += 1) {
      const x = plot.left + (plot.width * i) / 4;
      drawLine(ctx, x, plot.top, x, plot.bottom);
      const secondsAgo = (4 - i) * 30;
      const label = i === 4 ? "now" : `-${secondsAgo}s`;
      ctx.fillText(label, x - ctx.measureText(label).width / 2, plot.bottom + 22);
    }

    drawRect(ctx, plot, "#5e6c72");
    drawSeries(ctx, samples.map((sample) => [xScale(sample.timestamp), fieldY(sample.field)]), "#bd3f38");
    if (!fieldOnly) {
      drawSeries(ctx, samples.map((sample) => [xScale(sample.timestamp), referenceY(sample.reference)]), "#286ea8");
    }

    const spikeSamples = samples.filter((sample) => sample.rawField >= 0.1).length;
    ctx.fillStyle = "rgba(255, 255, 255, 0.94)";
    ctx.fillRect(plot.left + 8, plot.top + 8, Math.min(plot.width - 16, 620), 25);
    ctx.strokeStyle = "#bd3f38";
    ctx.lineWidth = 2.2;
    drawLine(ctx, plot.left + 16, plot.top + 21, plot.left + 48, plot.top + 21);
    ctx.fillStyle = "#1f292d";
    if (fieldOnly) {
      const fieldLabel = showNormalized
        ? `Field ${formatTrimmed(latest.field, 2)}% (${formatChamber(latest.rawField)})`
        : `Field ${formatChamber(latest.field)}`;
      ctx.fillText(`${fieldLabel}   Field samples >= 0.10: ${spikeSamples}`, plot.left + 62, plot.top + 23);
    } else {
      ctx.strokeStyle = "#286ea8";
      ctx.lineWidth = 2;
      drawLine(ctx, plot.left + 80, plot.top + 21, plot.left + 112, plot.top + 21);
      const fieldLabel = showNormalized
        ? `Field ${formatTrimmed(latest.field, 2)}% (${formatChamber(latest.rawField)})`
        : `Field ${formatChamber(latest.field)}`;
      const referenceLabel = showNormalized && isFiniteNumber(latest.reference)
        ? `Reference ${formatTrimmed(latest.reference, 2)}% (${formatChamber(latest.rawReference)})`
        : `Reference ${formatChamber(latest.reference)}`;
      ctx.fillText(`${fieldLabel}   ${referenceLabel}   Field samples >= 0.10: ${spikeSamples}`, plot.left + 122, plot.top + 23);
    }
  }

  function normalizeLiveValue(value, background, normalization) {
    if (!isFiniteNumber(value) || !isFiniteNumber(background) || !isFiniteNumber(normalization)) {
      return value;
    }

    const denominator = normalization - background;
    return Math.abs(denominator) > 0.000001
      ? (value - background) / denominator * 100
      : value;
  }

  function drawLine(ctx, x1, y1, x2, y2) {
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
  }

  function drawRect(ctx, plot, color) {
    ctx.strokeStyle = color;
    ctx.lineWidth = 1;
    ctx.strokeRect(plot.left, plot.top, plot.width, plot.height);
  }

  function drawRightText(ctx, text, right, y) {
    ctx.fillText(text, right - ctx.measureText(text).width, y);
  }

  function drawSeries(ctx, points, color) {
    if (!points.length) {
      return;
    }
    ctx.strokeStyle = color;
    ctx.lineWidth = 2;
    ctx.beginPath();
    points.forEach(([x, y], index) => {
      if (index === 0) {
        ctx.moveTo(x, y);
      } else {
        ctx.lineTo(x, y);
      }
    });
    ctx.stroke();

    ctx.fillStyle = color;
    const step = Math.max(1, Math.floor(points.length / 160));
    points.forEach(([x, y], index) => {
      if (index % step === 0) {
        ctx.beginPath();
        ctx.arc(x, y, 2.5, 0, Math.PI * 2);
        ctx.fill();
      }
    });
  }

  elements.connectButton.addEventListener("click", () => post("/api/connect"));
  elements.disconnectButton.addEventListener("click", () => post("/api/disconnect"));
  elements.sendMoveButton.addEventListener("click", () => post("/api/move", readTargetInputs()));
  elements.goIsocenterButton?.addEventListener("click", () => post("/api/move/isocenter", { speed: readTargetInputs().speed }));
  elements.goSurfaceButton.addEventListener("click", () => post("/api/move/surface", { speed: readTargetInputs().speed }));
  elements.biasButton.addEventListener("click", () => post("/api/bias"));
  elements.biasOffButton.addEventListener("click", () => post("/api/bias-off"));
  elements.applySensitivityButton.addEventListener("click", () => post("/api/electrometer/sensitivity", readSensitivitySettings()));
  elements.backgroundButton.addEventListener("click", takeBackground);
  elements.normalizeButton.addEventListener("click", takeNormalization);
  elements.prepareButton.addEventListener("click", prepareToScan);
  elements.runCenteringScanButton?.addEventListener("click", runCenteringScan);
  elements.workflowSteps?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-workflow-step]");
    if (!button || button.disabled) {
      return;
    }

    runWorkflowStep(button.dataset.workflowStep);
  });
  elements.detectorModeSelect.addEventListener("change", () => {
    storeUserPreferences();
    setCenteringWorkflowComplete(false);
    render();
  });
  elements.biasModeSelect?.addEventListener("change", () => {
    setBiasWorkflowComplete(false);
    setCenteringWorkflowComplete(false);
    if (snapshot?.connected && elements.biasModeSelect.value === "zero") {
      setBiasWorkflowComplete(true);
      pendingMessage = "Info: Bias/HV set to 0 V locally for diode-safe operation. No HV command was sent to the CCU.";
    }
    render();
  });
  elements.loadOpabButton.addEventListener("click", loadOpabQueue);
  elements.opabFileInput?.addEventListener("change", () => importSelectedOpabFile(elements.opabFileInput.files?.[0]));
  elements.queueCheckAllInput?.addEventListener("change", () => {
    const checked = elements.queueCheckAllInput.checked;
    elements.queueItemsBody.querySelectorAll(".queue-select").forEach((checkbox) => {
      checkbox.checked = checked;
    });
    renderQueueCheckAllState();
  });
  elements.runQueueButton.addEventListener("click", runSelectedQueue);
  elements.abortQueueButton.addEventListener("click", abortQueueRun);
  elements.queueItemsBody.addEventListener("change", (event) => {
    if (event.target.matches(".queue-select")) {
      renderQueueCheckAllState();
    }
  });
  elements.queueItemsBody.addEventListener("click", (event) => {
    if (event.target.closest("input")) {
      return;
    }

    openQueueItemResult(event.target.closest("tr[data-queue-index]"));
  });
  elements.turnAngleSelect?.addEventListener("change", updateSpigotDirection);
  elements.detectorModeSelect.addEventListener("change", render);
  elements.clearLogButton.addEventListener("click", () => {
    post("/api/logs/clear");
  });

  [elements.crosslineInput, elements.inlineInput, elements.depthInput, elements.speedInput].forEach((input) => {
    input.addEventListener("change", readTargetInputs);
  });

  [elements.queueCenterXInput, elements.queueCenterYInput, elements.queueSurfaceZInput].forEach((input) => {
    input.addEventListener("input", () => {
      coordinateInputsTouched = true;
      setCenteringWorkflowComplete(false);
    });
  });

  document.querySelectorAll("[data-preset]").forEach((button) => {
    button.addEventListener("click", () => {
      const preset = presets[button.dataset.preset];
      setTarget(preset.x, preset.y, preset.z);
    });
  });

  window.addEventListener("resize", () => {
    drawGraph();
    renderSelectedScanResult();
  });

  updateSpigotDirection();
  readTargetInputs();
  render();
  pollState();
  window.setInterval(pollState, 1000);
})();
