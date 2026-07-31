(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  }
  root.TankWorkflowReadiness = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  function isFiniteNumber(value) {
    return typeof value === "number" && Number.isFinite(value);
  }

  function normalizeDetectorMode(mode) {
    return mode === "fieldReference" ? "fieldReference" : "fieldOnly";
  }

  function detectorModeLabel(mode) {
    return normalizeDetectorMode(mode) === "fieldReference" ? "Field + reference" : "Field only";
  }

  function acquisitionMatchesMode(acquisition, mode) {
    return acquisition?.detectorMode === normalizeDetectorMode(mode);
  }

  function hasCurrentModeBackground(snapshot, mode) {
    const detectorMode = normalizeDetectorMode(mode);
    const background = snapshot?.latestBackground;
    if (background?.enabled !== true
      || !acquisitionMatchesMode(background, detectorMode)
      || !isFiniteNumber(background.fieldDetectorValue)) {
      return false;
    }

    return detectorMode === "fieldOnly" || isFiniteNumber(background.referenceDetectorValue);
  }

  function hasCurrentModeNormalization(snapshot, mode) {
    const detectorMode = normalizeDetectorMode(mode);
    const normalization = snapshot?.latestNormalization;
    const background = snapshot?.latestBackground;
    const normalizationUsable = normalization?.enabled === true
      || normalization?.status === "staleCoordinatesChanged";
    if (!normalizationUsable
      || (normalization.status !== "valid" && normalization.status !== "staleCoordinatesChanged")
      || !acquisitionMatchesMode(normalization, detectorMode)
      || !isFiniteNumber(normalization.fieldDetectorValue)
      || !normalizationIsCurrentForBackground(normalization, background)) {
      return false;
    }

    return detectorMode === "fieldOnly" || isFiniteNumber(normalization.referenceDetectorValue);
  }

  function normalizationIsCurrentForBackground(normalization, background) {
    if (background?.enabled !== true || !background.completedAt || !normalization?.completedAt) {
      return true;
    }

    return new Date(normalization.completedAt).getTime() >= new Date(background.completedAt).getTime();
  }

  function missingCoordinateMessage(missing) {
    if (!missing?.length) {
      return "";
    }

    if (missing.includes("iso") && missing.includes("surface")) {
      return "Set iso and surface. Surface can be optional on an SSD setup, but this prepare workflow uses it for depth.";
    }

    if (missing.includes("surface")) {
      return "Surface is not set. That can be fine on an SSD setup, but this prepare workflow uses surface for depth.";
    }

    return "Set iso";
  }

  function centeringReadinessIssue(snapshot, mode) {
    const detectorMode = normalizeDetectorMode(mode);
    if (!snapshot?.connected) {
      return "Connect before running a centering scan.";
    }

    const missing = snapshot?.coordinateState?.missing || [];
    if (missing.length) {
      return missingCoordinateMessage(missing);
    }

    if (!hasCurrentModeBackground(snapshot, detectorMode)) {
      return `${detectorModeLabel(detectorMode)} background is not ready. Take background with beam off.`;
    }

    if (!hasCurrentModeNormalization(snapshot, detectorMode)) {
      return `${detectorModeLabel(detectorMode)} normalization is not ready. Turn beam on and normalize.`;
    }

    if (snapshot?.preparation?.status !== "readyToScan") {
      return "Prepare to scan is not ready. Repeat Prepare to Scan if coordinates or background changed.";
    }

    return "";
  }

  function workflowHint(snapshot, mode, options = {}) {
    const detectorMode = normalizeDetectorMode(mode);
    const centeringReady = options.centeringReady === true;
    if (!snapshot?.connected) {
      return {
        className: "workflow-hint",
        text: "Connect to the tank, choose detector mode, then prepare to scan."
      };
    }

    const missing = snapshot?.coordinateState?.missing || [];
    if (missing.length) {
      return {
        className: "workflow-hint warning",
        text: missingCoordinateMessage(missing)
      };
    }

    const modeLabel = detectorModeLabel(detectorMode);
    if (!hasCurrentModeBackground(snapshot, detectorMode)) {
      return {
        className: "workflow-hint warning",
        text: `${modeLabel}: beam off, press Prepare to Scan or Background.`
      };
    }

    if (!hasCurrentModeNormalization(snapshot, detectorMode)) {
      return {
        className: "workflow-hint warning",
        text: `${modeLabel}: background is ready. Turn beam on, then press Normalize or continue Prepare to Scan.`
      };
    }

    if (snapshot?.preparation?.status === "readyToScan") {
      if (!centeringReady) {
        return {
          className: "workflow-hint warning",
          text: `${modeLabel}: ready for centering. Run the centering scan before treating setup as centered.`
        };
      }

      return {
        className: "workflow-hint good",
        text: `${modeLabel}: centered and ready to scan.`
      };
    }

    return {
      className: "workflow-hint",
      text: `${modeLabel}: measurements are present; press Prepare to Scan if the setup changed.`
    };
  }

  function workflowSteps(snapshot, mode, options = {}) {
    const detectorMode = normalizeDetectorMode(mode);
    const missing = snapshot?.coordinateState?.missing || [];
    const biasReady = options.biasReady === true || snapshot?.biasReady === true;
    const centeringReady = options.centeringReady === true;
    const backgroundReady = hasCurrentModeBackground(snapshot, detectorMode);
    const normalizationReady = hasCurrentModeNormalization(snapshot, detectorMode);
    const centeringIssue = centeringReadinessIssue(snapshot, detectorMode);

    return [
      {
        key: "connected",
        label: snapshot?.connected ? "Disconnect" : "Connect",
        status: snapshot?.connected ? "good" : "pending",
        detail: snapshot?.connected ? "Connected" : "Not connected"
      },
      {
        key: "coordinates",
        label: "Coordinates",
        status: snapshot?.connected && missing.length === 0 ? "good" : "pending",
        detail: missing.length === 0 ? "Set" : missingCoordinateMessage(missing)
      },
      {
        key: "bias",
        label: "Bias",
        status: snapshot?.connected && biasReady ? "good" : "pending",
        detail: snapshot?.connected ? biasReady ? "0 V protected" : "Set bias" : "After connect"
      },
      {
        key: "background",
        label: "Background",
        status: backgroundReady ? "good" : "pending",
        detail: backgroundReady ? "Ready" : `${detectorModeLabel(detectorMode)} needed`
      },
      {
        key: "normalization",
        label: "Normalize",
        status: normalizationReady ? "good" : "pending",
        detail: normalizationReady ? "Ready" : backgroundReady ? "Beam on" : "After background"
      },
      {
        key: "centering",
        label: "Centering",
        status: centeringReady ? "good" : "pending",
        detail: centeringReady ? "Centered" : !centeringIssue ? "Run scan" : "Blocked"
      }
    ];
  }

  return {
    centeringReadinessIssue,
    detectorModeLabel,
    hasCurrentModeBackground,
    hasCurrentModeNormalization,
    missingCoordinateMessage,
    normalizeDetectorMode,
    normalizationIsCurrentForBackground,
    workflowHint,
    workflowSteps
  };
});
