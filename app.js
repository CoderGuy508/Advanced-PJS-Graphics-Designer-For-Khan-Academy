const STORAGE_KEY = "pjs_graphics_designer_state_v2";
const AUTOSAVE_DELAY_MS = 900;
const FILL_TOLERANCE = 24;
const MIN_CANVAS_SIZE = 100;
const MAX_CANVAS_SIZE = 2000;
const TAU = Math.PI * 2;
const SELECT_HANDLE_SIZE = 12;
const ROTATE_HANDLE_SIZE = 14;
const EDGE_HANDLE_SIZE = 10;
const MAX_HISTORY_STEPS = 120;
const COPY_PASTE_OFFSET = 18;

const drawCanvas = document.getElementById("drawCanvas");
const previewCanvas = document.getElementById("previewCanvas");
const canvasStage = document.getElementById("canvasStage");
const canvasInfo = document.getElementById("canvasInfo");
const ctx = drawCanvas.getContext("2d", { willReadFrequently: true });
const previewCtx = previewCanvas.getContext("2d");

const toolButtonsRoot = document.getElementById("toolButtons");
const colorPicker = document.getElementById("colorPicker");
const sizeSlider = document.getElementById("sizeSlider");
const strokeColorPicker = document.getElementById("strokeColorPicker");
const strokeWeightSlider = document.getElementById("strokeWeightSlider");
const arcStartSlider = document.getElementById("arcStartSlider");
const arcStopSlider = document.getElementById("arcStopSlider");
const opacitySlider = document.getElementById("opacitySlider");
const shapeFillToggle = document.getElementById("shapeFillToggle");
const gradientToggle = document.getElementById("gradientToggle");
const gradientColorPicker = document.getElementById("gradientColorPicker");
const gradientDirectionSelect = document.getElementById("gradientDirectionSelect");
const sizeValue = document.getElementById("sizeValue");
const strokeWeightValue = document.getElementById("strokeWeightValue");
const arcStartValue = document.getElementById("arcStartValue");
const arcStopValue = document.getElementById("arcStopValue");
const opacityValue = document.getElementById("opacityValue");
const saveStatus = document.getElementById("saveStatus");

const canvasWidthInput = document.getElementById("canvasWidthInput");
const canvasHeightInput = document.getElementById("canvasHeightInput");
const resizeCanvasBtn = document.getElementById("resizeCanvasBtn");

const selectionInfo = document.getElementById("selectionInfo");
const pathInfo = document.getElementById("pathInfo");
const applyStyleBtn = document.getElementById("applyStyleBtn");
const deleteSelectedBtn = document.getElementById("deleteSelectedBtn");
const finishPathBtn = document.getElementById("finishPathBtn");
const cancelPathBtn = document.getElementById("cancelPathBtn");
const bringForwardBtn = document.getElementById("bringForwardBtn");
const sendBackwardBtn = document.getElementById("sendBackwardBtn");
const bringToFrontBtn = document.getElementById("bringToFrontBtn");
const sendToBackBtn = document.getElementById("sendToBackBtn");

const undoBtn = document.getElementById("undoBtn");
const redoBtn = document.getElementById("redoBtn");
const clearBtn = document.getElementById("clearBtn");
const saveBtn = document.getElementById("saveBtn");
const restoreBtn = document.getElementById("restoreBtn");
const downloadPngBtn = document.getElementById("downloadPngBtn");

const generatePjsBtn = document.getElementById("generatePjsBtn");
const copyPjsBtn = document.getElementById("copyPjsBtn");
const downloadPjsBtn = document.getElementById("downloadPjsBtn");
const imageModeToggle = document.getElementById("imageModeToggle");
const pjsOutput = document.getElementById("pjsOutput");

const state = {
  tool: "select",
  color: "#00b7ff",
  size: 8,
  strokeColor: "#0d2238",
  strokeWeight: 2,
  arcStartDeg: 0,
  arcStopDeg: 180,
  opacity: 1,
  fillShapes: false,
  gradientEnabled: false,
  gradientColor: "#ff8a3d",
  gradientDirection: "left-right",
  exportImageMode: false,
  canvasWidth: 400,
  canvasHeight: 400,
  actions: [],
  undoStack: [],
  redoStack: [],
  drawing: false,
  currentAction: null,
  startPoint: null,
  activePointerId: null,
  autosaveTimer: null,
  selectedActionId: null,
  selectedActionIds: [],
  dragMode: null,
  dragStart: null,
  dragOriginalAction: null,
  dragOriginalSelection: null,
  dragOriginalBounds: null,
  dragDidChange: false,
  pathDraft: null,
  hoverPoint: null,
  dragPointIndex: -1,
  dragEdgeInfo: null,
  dragRotateCenter: null,
  dragRotateStartAngle: 0,
  dragMarqueeCurrent: null,
  clipboardAction: null
};

let actionCounter = 1;

init();

function init() {
  setCanvasSize(state.canvasWidth, state.canvasHeight);
  resetCanvas();
  bindUI();

  const restored = restoreFromStorage(false);
  if (!restored) {
    redrawAll();
    updateSaveStatus("Autosave ready");
  }

  syncControlsFromState();
  highlightActiveTool();
  updateUndoRedoButtons();
  updateSelectionInfo();
  updatePathInfo();
  refreshGradientUi();
  updateCursor();
  renderOverlay();
  updatePjsOutput();
}

function bindUI() {
  toolButtonsRoot.addEventListener("click", (event) => {
    const button = event.target.closest("[data-tool]");
    if (!button) {
      return;
    }
    setTool(button.dataset.tool);
  });

  colorPicker.addEventListener("input", () => {
    state.color = colorPicker.value;
    queueAutosave();
  });

  sizeSlider.addEventListener("input", () => {
    state.size = Number(sizeSlider.value);
    sizeValue.textContent = String(state.size);
    queueAutosave();
  });

  strokeColorPicker.addEventListener("input", () => {
    state.strokeColor = strokeColorPicker.value;
    queueAutosave();
  });

  strokeWeightSlider.addEventListener("input", () => {
    state.strokeWeight = Number(strokeWeightSlider.value);
    strokeWeightValue.textContent = String(state.strokeWeight);
    queueAutosave();
  });

  arcStartSlider.addEventListener("input", () => {
    state.arcStartDeg = clamp(Math.round(Number(arcStartSlider.value) || 0), 0, 360);
    arcStartValue.textContent = `${state.arcStartDeg}°`;
    queueAutosave();
  });

  arcStopSlider.addEventListener("input", () => {
    state.arcStopDeg = clamp(Math.round(Number(arcStopSlider.value) || 180), 0, 360);
    arcStopValue.textContent = `${state.arcStopDeg}°`;
    queueAutosave();
  });

  opacitySlider.addEventListener("input", () => {
    state.opacity = Number(opacitySlider.value) / 100;
    opacityValue.textContent = `${Math.round(state.opacity * 100)}%`;
    queueAutosave();
  });

  shapeFillToggle.addEventListener("change", () => {
    state.fillShapes = shapeFillToggle.checked;
    queueAutosave();
  });

  gradientToggle.addEventListener("change", () => {
    if (gradientToggle.disabled) {
      gradientToggle.checked = false;
      return;
    }
    state.gradientEnabled = gradientToggle.checked;
    refreshGradientUi();
    queueAutosave();
  });

  gradientColorPicker.addEventListener("input", () => {
    state.gradientColor = gradientColorPicker.value;
    queueAutosave();
  });

  gradientDirectionSelect.addEventListener("change", () => {
    state.gradientDirection = normalizeGradientDirection(gradientDirectionSelect.value);
    queueAutosave();
  });

  imageModeToggle.addEventListener("change", () => {
    state.exportImageMode = Boolean(imageModeToggle.checked);
    updatePjsOutput();
    queueAutosave();
  });

  resizeCanvasBtn.addEventListener("click", onResizeCanvas);

  drawCanvas.addEventListener("pointerdown", onPointerDown);
  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
  window.addEventListener("pointercancel", onPointerUp);

  applyStyleBtn.addEventListener("click", applyStyleToSelected);
  deleteSelectedBtn.addEventListener("click", deleteSelectedAction);
  finishPathBtn.addEventListener("click", finalizePathDraft);
  cancelPathBtn.addEventListener("click", cancelPathDraft);
  bringForwardBtn.addEventListener("click", bringSelectionForwardOne);
  sendBackwardBtn.addEventListener("click", sendSelectionBackwardOne);
  bringToFrontBtn.addEventListener("click", bringSelectionToFront);
  sendToBackBtn.addEventListener("click", sendSelectionToBack);

  undoBtn.addEventListener("click", undo);
  redoBtn.addEventListener("click", redo);

  clearBtn.addEventListener("click", () => {
    if (state.actions.length === 0) {
      return;
    }
    if (!window.confirm("Clear the entire canvas and history?")) {
      return;
    }
    pushUndoSnapshot();
    state.redoStack = [];
    state.actions = [];
    state.selectedActionId = null;
    state.selectedActionIds = [];
    state.pathDraft = null;
    state.hoverPoint = null;
    refreshAfterActionChange("Canvas cleared");
  });

  saveBtn.addEventListener("click", () => {
    saveToStorage(true);
  });

  restoreBtn.addEventListener("click", () => {
    restoreFromStorage(true);
  });

  downloadPngBtn.addEventListener("click", downloadPng);

  generatePjsBtn.addEventListener("click", updatePjsOutput);
  copyPjsBtn.addEventListener("click", copyPjsOutput);
  downloadPjsBtn.addEventListener("click", downloadPjs);

  window.addEventListener("keydown", onKeyDown);
  window.addEventListener("beforeunload", () => saveToStorage(false));
}

function onKeyDown(event) {
  const targetTag = event.target && event.target.tagName ? event.target.tagName.toLowerCase() : "";
  const typingInInput = targetTag === "input" || targetTag === "textarea";

  if (!typingInInput && (event.key === "Delete" || event.key === "Backspace")) {
    if (state.selectedActionIds.length > 0) {
      event.preventDefault();
      deleteSelectedAction();
    }
    return;
  }

  if (!typingInInput && isPathTool(state.tool)) {
    if (event.key === "Enter") {
      event.preventDefault();
      finalizePathDraft();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      cancelPathDraft();
      return;
    }
  }

  const isModifier = event.metaKey || event.ctrlKey;
  if (!isModifier) {
    return;
  }

  const key = event.key.toLowerCase();

  if (key === "z" && event.shiftKey) {
    event.preventDefault();
    redo();
    return;
  }

  if (key === "z") {
    event.preventDefault();
    undo();
    return;
  }

  if (key === "y") {
    event.preventDefault();
    redo();
    return;
  }

  if (!typingInInput && key === "c") {
    event.preventDefault();
    copySelectedAction();
    return;
  }

  if (!typingInInput && key === "v") {
    event.preventDefault();
    pasteCopiedAction();
    return;
  }

  if (key === "s") {
    event.preventDefault();
    saveToStorage(true);
  }
}

function onResizeCanvas() {
  const requestedWidth = Number(canvasWidthInput.value);
  const requestedHeight = Number(canvasHeightInput.value);

  const width = clamp(Math.round(requestedWidth || state.canvasWidth), MIN_CANVAS_SIZE, MAX_CANVAS_SIZE);
  const height = clamp(Math.round(requestedHeight || state.canvasHeight), MIN_CANVAS_SIZE, MAX_CANVAS_SIZE);

  if (width === state.canvasWidth && height === state.canvasHeight) {
    updateSaveStatus("Canvas size unchanged");
    return;
  }

  setCanvasSize(width, height);
  redrawAll();
  sanitizeSelection();
  renderOverlay();
  updatePjsOutput();
  queueAutosave(true);
  updateSaveStatus(`Canvas resized to ${width}x${height}`);
}

function setTool(tool) {
  if (state.tool !== tool && state.pathDraft) {
    const keepDraft = isPathTool(tool) && state.pathDraft.tool === tool;
    if (!keepDraft) {
      state.pathDraft = null;
      state.hoverPoint = null;
    }
  }
  state.tool = tool;
  if (tool === "vertex" || tool === "curve") {
    state.gradientEnabled = false;
  }
  highlightActiveTool();
  updateCursor();
  updatePathInfo();
  refreshGradientUi();
  renderOverlay();
  queueAutosave();
}

function highlightActiveTool() {
  const buttons = toolButtonsRoot.querySelectorAll(".tool-btn");
  buttons.forEach((button) => {
    button.classList.toggle("active", button.dataset.tool === state.tool);
  });
}

function syncControlsFromState() {
  colorPicker.value = state.color;
  sizeSlider.value = String(state.size);
  strokeColorPicker.value = state.strokeColor;
  strokeWeightSlider.value = String(state.strokeWeight);
  arcStartSlider.value = String(state.arcStartDeg);
  arcStopSlider.value = String(state.arcStopDeg);
  opacitySlider.value = String(Math.round(state.opacity * 100));
  shapeFillToggle.checked = state.fillShapes;
  gradientToggle.checked = state.gradientEnabled;
  gradientColorPicker.value = state.gradientColor;
  gradientDirectionSelect.value = normalizeGradientDirection(state.gradientDirection);
  imageModeToggle.checked = state.exportImageMode;
  sizeValue.textContent = String(state.size);
  strokeWeightValue.textContent = String(state.strokeWeight);
  arcStartValue.textContent = `${Math.round(state.arcStartDeg)}°`;
  arcStopValue.textContent = `${Math.round(state.arcStopDeg)}°`;
  opacityValue.textContent = `${Math.round(state.opacity * 100)}%`;
  canvasWidthInput.value = String(state.canvasWidth);
  canvasHeightInput.value = String(state.canvasHeight);
  refreshGradientUi();
}

function loadStyleFromAction(action) {
  if (!action) {
    return;
  }

  if (typeof action.fillColor === "string") {
    state.color = action.fillColor;
  } else if (typeof action.color === "string") {
    state.color = action.color;
  }

  if (typeof action.strokeColor === "string") {
    state.strokeColor = action.strokeColor;
  }

  if (typeof action.size === "number") {
    state.size = clamp(action.size, 1, 60);
  }
  if (typeof action.strokeWeight === "number") {
    state.strokeWeight = clamp(action.strokeWeight, 1, 40);
  }
  if (action.type === "arc") {
    state.arcStartDeg = clamp(Math.round(((action.start || 0) * 180) / Math.PI), 0, 360);
    state.arcStopDeg = clamp(Math.round(((action.stop || Math.PI) * 180) / Math.PI), 0, 360);
  }
  if (typeof action.opacity === "number") {
    state.opacity = clamp(action.opacity, 0.05, 1);
  }
  if ((action.type === "rect" || action.type === "ellipse" || action.type === "arc" || action.type === "polygonShape" || action.type === "vertexPath") && typeof action.filled === "boolean") {
    state.fillShapes = action.filled;
    state.gradientEnabled = (action.type === "vertexPath" || action.type === "arc") ? false : Boolean(action.gradient);
    if (action.gradient && typeof action.gradient.endColor === "string") {
      state.gradientColor = action.gradient.endColor;
    }
    if (action.gradient && typeof action.gradient.direction === "string") {
      state.gradientDirection = normalizeGradientDirection(action.gradient.direction);
    }
  }

  syncControlsFromState();
}

function canUseGradientForCurrentTarget() {
  if (state.tool === "vertex" || state.tool === "curve" || state.tool === "arc") {
    return false;
  }
  if (state.tool === "select") {
    const selectedActions = getSelectedActions();
    if (selectedActions.some((selected) => selected.type === "vertexPath" || selected.type === "arc")) {
      return false;
    }
  }
  return true;
}

function refreshGradientUi() {
  const allowed = canUseGradientForCurrentTarget();
  if (!allowed) {
    state.gradientEnabled = false;
  }
  gradientToggle.disabled = !allowed;
  gradientToggle.checked = allowed ? state.gradientEnabled : false;
  gradientColorPicker.disabled = !allowed || !state.gradientEnabled;
  gradientDirectionSelect.disabled = !allowed || !state.gradientEnabled;
}

function updateCursor() {
  if (state.tool === "select") {
    drawCanvas.style.cursor = state.selectedActionIds.length > 0 ? "grab" : "default";
    return;
  }

  if (state.tool === "eraser") {
    drawCanvas.style.cursor = "cell";
    return;
  }

  const crosshairTools = new Set(["line", "rect", "ellipse", "arc", "triangle", "quad", "bezier", "vertex", "curve", "fill"]);
  drawCanvas.style.cursor = crosshairTools.has(state.tool) ? "crosshair" : "default";
}

function onPointerDown(event) {
  if (state.activePointerId !== null) {
    return;
  }

  event.preventDefault();
  state.activePointerId = event.pointerId;
  drawCanvas.setPointerCapture(event.pointerId);

  const point = getCanvasPoint(event);

  if (isPathTool(state.tool) && state.pathDraft) {
    state.hoverPoint = point;
    renderOverlay();
  }

  if (state.tool === "select") {
    startSelectInteraction(point);
    return;
  }

  if (state.tool === "fill") {
    const fillAction = {
      id: createActionId(),
      type: "fill",
      x: Math.round(point.x),
      y: Math.round(point.y),
      color: state.color,
      opacity: state.opacity,
      tolerance: FILL_TOLERANCE
    };

    pushUndoSnapshot();
    const changed = drawAction(ctx, fillAction);
    if (changed) {
      state.redoStack = [];
      state.actions.push(fillAction);
      state.selectedActionId = null;
      state.selectedActionIds = [];
      refreshAfterActionChange("Fill applied");
    } else {
      state.undoStack.pop();
      updateSaveStatus("Fill region unchanged");
    }

    releasePointer(event.pointerId);
    return;
  }

  if (isPathTool(state.tool)) {
    handlePathPoint(point);
    releasePointer(event.pointerId);
    return;
  }

  if (isDragShapeTool(state.tool)) {
    state.drawing = true;
    state.startPoint = point;
    state.currentAction = buildShapeAction(state.tool, point, point);
    renderOverlay();
    return;
  }

  state.drawing = true;
  state.startPoint = point;

  state.currentAction = {
    id: createActionId(),
    type: "stroke",
    brush: state.tool,
    color: state.color,
    size: state.size,
    opacity: state.opacity,
    points: [point],
    dots: []
  };

  if (state.tool === "spray") {
    sprayAt(point, state.currentAction);
  } else {
    drawStrokeDot(ctx, state.currentAction, point);
  }
}

function onPointerMove(event) {
  if (isPathTool(state.tool) && state.pathDraft) {
    state.hoverPoint = getCanvasPoint(event);
    renderOverlay();
  }

  if (state.activePointerId !== event.pointerId) {
    return;
  }

  const point = getCanvasPoint(event);

  if (state.tool === "select") {
    if (!state.drawing || !state.dragMode) {
      return;
    }
    if (state.dragMode !== "marquee" && state.selectedActionIds.length === 0) {
      return;
    }
    event.preventDefault();
    updateSelectInteraction(point);
    return;
  }

  if (!state.drawing || !state.currentAction) {
    return;
  }

  event.preventDefault();

  if (isDragShapeTool(state.tool)) {
    state.currentAction = buildShapeAction(state.tool, state.startPoint, point);
    renderOverlay();
    return;
  }

  const action = state.currentAction;
  if (action.brush === "spray") {
    const lastPoint = action.points[action.points.length - 1];
    if (distance(lastPoint, point) < 1.4) {
      return;
    }
    action.points.push(point);
    sprayAt(point, action);
    return;
  }

  const previousPoint = action.points[action.points.length - 1];
  if (distance(previousPoint, point) < 0.4) {
    return;
  }

  action.points.push(point);
  drawStrokeSegment(ctx, action, previousPoint, point);
}

function onPointerUp(event) {
  if (state.activePointerId !== event.pointerId) {
    return;
  }

  if (state.tool === "select") {
    if (state.dragMode === "marquee") {
      event.preventDefault();
      finalizeMarqueeSelection(getCanvasPoint(event));
      endSelectInteraction();
      renderOverlay();
      releasePointer(event.pointerId);
      return;
    }
    if (state.dragDidChange) {
      refreshAfterActionChange("Object updated");
    } else {
      renderOverlay();
    }
    endSelectInteraction();
    releasePointer(event.pointerId);
    return;
  }

  if (!state.drawing || !state.currentAction) {
    releasePointer(event.pointerId);
    return;
  }

  event.preventDefault();

  if (isDragShapeTool(state.tool)) {
    const finalPoint = getCanvasPoint(event);
    const finalAction = buildShapeAction(state.tool, state.startPoint, finalPoint);

    if (state.tool === "line" || state.tool === "bezier" || isActionLargeEnough(finalAction)) {
      commitNewAction(finalAction, "Shape added");
    } else {
      updateSaveStatus("Shape too small to add");
      renderOverlay();
    }
  } else {
    const finalAction = state.currentAction;
    if (finalAction.brush === "spray") {
      if (finalAction.dots.length > 0) {
        commitNewAction(finalAction, "Stroke added");
      }
    } else if (finalAction.points.length > 0) {
      commitNewAction(finalAction, "Stroke added");
    }
  }

  state.drawing = false;
  state.currentAction = null;
  state.startPoint = null;
  releasePointer(event.pointerId);
}

function releasePointer(pointerId) {
  if (drawCanvas.hasPointerCapture(pointerId)) {
    drawCanvas.releasePointerCapture(pointerId);
  }

  state.activePointerId = null;
}

function isDragShapeTool(tool) {
  return tool === "line" || tool === "rect" || tool === "ellipse" || tool === "arc" || tool === "bezier";
}

function isPathTool(tool) {
  return tool === "vertex" || tool === "curve" || tool === "triangle" || tool === "quad";
}

function getPathToolConfig(tool) {
  if (tool === "vertex") {
    return { type: "vertexPath", curved: false, minPoints: 2, canCloseOnFirst: true };
  }
  if (tool === "curve") {
    return { type: "vertexPath", curved: true, minPoints: 2, canCloseOnFirst: true };
  }
  if (tool === "triangle") {
    return { type: "polygonShape", shape: "triangle", minPoints: 3, autoCompletePoints: 3, closed: true };
  }
  if (tool === "quad") {
    return { type: "polygonShape", shape: "quad", minPoints: 4, autoCompletePoints: 4, closed: true };
  }
  return null;
}

function isActionLargeEnough(action) {
  if (typeof action.w === "number" && typeof action.h === "number") {
    return Math.abs(action.w * action.h) > 0.5;
  }
  if (action.type === "line") {
    return distance({ x: action.x1, y: action.y1 }, { x: action.x2, y: action.y2 }) > 0.5;
  }
  return true;
}

function handlePathPoint(point) {
  const config = getPathToolConfig(state.tool);
  if (!config) {
    return;
  }

  if (!state.pathDraft || state.pathDraft.tool !== state.tool) {
    state.pathDraft = { tool: state.tool, curved: Boolean(config.curved), points: [] };
  }

  const points = state.pathDraft.points;
  if (config.canCloseOnFirst && points.length >= 3 && distance(points[0], point) <= 8) {
    finalizePathDraft(true);
    return;
  }

  points.push(point);
  if (config.autoCompletePoints && points.length >= config.autoCompletePoints) {
    finalizePathDraft(Boolean(config.closed));
    return;
  }
  updatePathInfo();
  renderOverlay();
}

function finalizePathDraft(closeShape) {
  if (!state.pathDraft || !Array.isArray(state.pathDraft.points)) {
    updatePathInfo();
    return;
  }

  const config = getPathToolConfig(state.pathDraft.tool);
  if (!config || state.pathDraft.points.length < config.minPoints) {
    updatePathInfo();
    return;
  }

  let action;
  if (config.type === "polygonShape") {
    action = {
      id: createActionId(),
      type: "polygonShape",
      shape: config.shape,
      points: clonePointList(state.pathDraft.points),
      fillColor: state.color,
      strokeColor: state.strokeColor,
      strokeWeight: state.strokeWeight,
      arcStartDeg: state.arcStartDeg,
      arcStopDeg: state.arcStopDeg,
      opacity: state.opacity,
      filled: state.fillShapes,
      gradient: state.gradientEnabled
        ? { endColor: state.gradientColor, direction: state.gradientDirection }
        : null
    };
  } else {
    action = {
      id: createActionId(),
      type: "vertexPath",
      curved: state.pathDraft.curved,
      closed: Boolean(closeShape),
      points: clonePointList(state.pathDraft.points),
      fillColor: state.color,
      strokeColor: state.strokeColor,
      strokeWeight: state.strokeWeight,
      opacity: state.opacity,
      filled: state.fillShapes,
      gradient: null
    };
  }

  state.pathDraft = null;
  state.hoverPoint = null;
  commitNewAction(action, "Path added");
}

function cancelPathDraft() {
  if (!state.pathDraft) {
    updatePathInfo();
    return;
  }
  state.pathDraft = null;
  state.hoverPoint = null;
  updatePathInfo();
  renderOverlay();
}

function startSelectInteraction(point) {
  const selectedActions = getSelectedActions();
  const selected = getSelectedAction();
  const hasMultiSelection = selectedActions.length > 1;

  if (hasMultiSelection) {
    const groupBounds = getSelectedBounds();
    if (groupBounds && isPointOnRotateHandle(point, groupBounds)) {
      beginSelectionDrag("rotate", point, null, groupBounds);
      state.dragRotateCenter = {
        x: (groupBounds.xMin + groupBounds.xMax) / 2,
        y: (groupBounds.yMin + groupBounds.yMax) / 2
      };
      state.dragRotateStartAngle = angleBetween(state.dragRotateCenter, point);
      return;
    }
    if (groupBounds && isPointOnResizeHandle(point, groupBounds)) {
      beginSelectionDrag("resize", point, null, groupBounds);
      return;
    }
    if (groupBounds && isPointInsideBounds(point, groupBounds)) {
      beginSelectionDrag("move", point, null, groupBounds);
      return;
    }
  }

  if (selected && !hasMultiSelection) {
    const selectedBounds = getActionBounds(selected);
    const editablePoints = getEditablePoints(selected);
    if (editablePoints.length > 0) {
      const pointHit = findControlPointHit(point, editablePoints);
      if (pointHit >= 0) {
        beginSelectionDrag("point", point, selected, selectedBounds);
        state.dragPointIndex = pointHit;
        return;
      }
    }
    const edgeHandles = getEditableEdgeHandles(selected);
    if (edgeHandles.length > 0) {
      const edgeHit = findEdgeHandleHit(point, edgeHandles);
      if (edgeHit >= 0) {
        beginSelectionDrag("edge", point, selected, selectedBounds);
        state.dragEdgeInfo = edgeHandles[edgeHit];
        return;
      }
    }
    if (selectedBounds && isPointOnRotateHandle(point, selectedBounds) && isActionRotatable(selected)) {
      beginSelectionDrag("rotate", point, selected, selectedBounds);
      state.dragRotateCenter = {
        x: (selectedBounds.xMin + selectedBounds.xMax) / 2,
        y: (selectedBounds.yMin + selectedBounds.yMax) / 2
      };
      state.dragRotateStartAngle = angleBetween(state.dragRotateCenter, point);
      return;
    }
    if (selectedBounds && isPointOnResizeHandle(point, selectedBounds)) {
      beginSelectionDrag("resize", point, selected, selectedBounds);
      return;
    }
  }

  const hitAction = findTopActionAtPoint(point);
  if (!hitAction) {
    state.drawing = false;
    state.dragMode = null;
    state.dragStart = null;
    state.dragOriginalAction = null;
    state.dragOriginalBounds = null;
    state.dragDidChange = false;
    state.selectedActionId = null;
    state.selectedActionIds = [];
    state.dragMarqueeCurrent = point;
    beginSelectionDrag("marquee", point, null, null);
    updateSelectionInfo();
    refreshGradientUi();
    updateCursor();
    renderOverlay();
    return;
  }

  if (state.selectedActionIds.includes(hitAction.id) && state.selectedActionIds.length > 1) {
    beginSelectionDrag("move", point, null, getSelectedBounds());
  } else {
    state.selectedActionId = hitAction.id;
    state.selectedActionIds = [hitAction.id];
    loadStyleFromAction(hitAction);
    beginSelectionDrag("move", point, hitAction, getActionBounds(hitAction));
  }
  updateSelectionInfo();
  refreshGradientUi();
  renderOverlay();
}

function beginSelectionDrag(mode, point, action, bounds) {
  state.drawing = true;
  state.dragMode = mode;
  state.dragStart = point;
  state.dragOriginalAction = action ? cloneAction(action) : null;
  state.dragOriginalSelection = null;
  if ((mode === "move" || mode === "resize" || mode === "rotate") && state.selectedActionIds.length > 1) {
    state.dragOriginalSelection = state.selectedActionIds
      .map((id) => {
        const index = findActionIndexById(id);
        if (index < 0) {
          return null;
        }
        return { id, action: cloneAction(state.actions[index]) };
      })
      .filter(Boolean);
  }
  state.dragOriginalBounds = bounds ? { ...bounds } : (action ? getActionBounds(action) : null);
  state.dragDidChange = false;
  state.dragPointIndex = -1;
  state.dragEdgeInfo = null;
  state.dragRotateCenter = null;
  state.dragRotateStartAngle = 0;
  state.dragMarqueeCurrent = mode === "marquee" ? point : null;
  drawCanvas.style.cursor = "grabbing";
}

function updateSelectInteraction(point) {
  if (state.dragMode === "marquee") {
    state.dragMarqueeCurrent = point;
    renderOverlay();
    return;
  }

  const actionIndex = findActionIndexById(state.selectedActionId);
  if (actionIndex < 0) {
    endSelectInteraction();
    return;
  }

  if (state.dragMode === "move") {
    const dx = point.x - state.dragStart.x;
    const dy = point.y - state.dragStart.y;
    if (Math.abs(dx) < 0.05 && Math.abs(dy) < 0.05) {
      return;
    }
    beginDragMutationIfNeeded();
    if (Array.isArray(state.dragOriginalSelection) && state.dragOriginalSelection.length > 0) {
      for (const item of state.dragOriginalSelection) {
        const index = findActionIndexById(item.id);
        if (index >= 0) {
          state.actions[index] = moveAction(item.action, dx, dy);
        }
      }
    } else {
      state.actions[actionIndex] = moveAction(state.dragOriginalAction, dx, dy);
    }
  } else if (state.dragMode === "point") {
    if (state.dragPointIndex < 0) {
      return;
    }
    beginDragMutationIfNeeded();
    state.actions[actionIndex] = moveActionControlPoint(state.dragOriginalAction, state.dragPointIndex, point);
  } else if (state.dragMode === "edge") {
    if (!state.dragEdgeInfo) {
      return;
    }
    const dx = point.x - state.dragStart.x;
    const dy = point.y - state.dragStart.y;
    if (Math.abs(dx) < 0.05 && Math.abs(dy) < 0.05) {
      return;
    }
    beginDragMutationIfNeeded();
    state.actions[actionIndex] = moveActionEdge(state.dragOriginalAction, state.dragEdgeInfo, dx, dy);
  } else if (state.dragMode === "rotate") {
    if (!state.dragRotateCenter) {
      return;
    }
    const currentAngle = angleBetween(state.dragRotateCenter, point);
    const delta = currentAngle - state.dragRotateStartAngle;
    if (Math.abs(delta) < 0.001) {
      return;
    }
    beginDragMutationIfNeeded();
    if (Array.isArray(state.dragOriginalSelection) && state.dragOriginalSelection.length > 0) {
      for (const item of state.dragOriginalSelection) {
        const index = findActionIndexById(item.id);
        if (index >= 0) {
          state.actions[index] = rotateAction(item.action, state.dragRotateCenter, delta);
        }
      }
    } else {
      state.actions[actionIndex] = rotateAction(state.dragOriginalAction, state.dragRotateCenter, delta);
    }
  } else if (state.dragMode === "resize") {
    const bounds = state.dragOriginalBounds;
    if (!bounds) {
      return;
    }

    const anchorX = bounds.xMin;
    const anchorY = bounds.yMin;
    const startWidth = Math.max(2, bounds.xMax - bounds.xMin);
    const startHeight = Math.max(2, bounds.yMax - bounds.yMin);

    const targetWidth = Math.max(2, point.x - anchorX);
    const targetHeight = Math.max(2, point.y - anchorY);

    const scaleX = targetWidth / startWidth;
    const scaleY = targetHeight / startHeight;

    if (Math.abs(scaleX - 1) < 0.001 && Math.abs(scaleY - 1) < 0.001) {
      return;
    }

    beginDragMutationIfNeeded();
    if (Array.isArray(state.dragOriginalSelection) && state.dragOriginalSelection.length > 0) {
      for (const item of state.dragOriginalSelection) {
        const index = findActionIndexById(item.id);
        if (index >= 0) {
          state.actions[index] = scaleActionFrom(item.action, anchorX, anchorY, scaleX, scaleY);
        }
      }
    } else {
      state.actions[actionIndex] = scaleActionFrom(state.dragOriginalAction, anchorX, anchorY, scaleX, scaleY);
    }
  }

  redrawAll();
  renderOverlay();
  updateSelectionInfo();
}

function beginDragMutationIfNeeded() {
  if (state.dragDidChange) {
    return;
  }
  pushUndoSnapshot();
  state.redoStack = [];
  state.dragDidChange = true;
  updateUndoRedoButtons();
}

function endSelectInteraction() {
  state.drawing = false;
  state.currentAction = null;
  state.startPoint = null;
  state.dragMode = null;
  state.dragStart = null;
  state.dragOriginalAction = null;
  state.dragOriginalSelection = null;
  state.dragOriginalBounds = null;
  state.dragDidChange = false;
  state.dragPointIndex = -1;
  state.dragEdgeInfo = null;
  state.dragRotateCenter = null;
  state.dragRotateStartAngle = 0;
  state.dragMarqueeCurrent = null;
  updateCursor();
}

function finalizeMarqueeSelection(point) {
  const start = state.dragStart;
  if (!start) {
    return;
  }
  const xMin = Math.min(start.x, point.x);
  const yMin = Math.min(start.y, point.y);
  const xMax = Math.max(start.x, point.x);
  const yMax = Math.max(start.y, point.y);
  const width = xMax - xMin;
  const height = yMax - yMin;

  if (width < 3 && height < 3) {
    state.selectedActionId = null;
    state.selectedActionIds = [];
    updateSelectionInfo();
    refreshGradientUi();
    renderOverlay();
    return;
  }

  const selected = [];
  for (let i = state.actions.length - 1; i >= 0; i -= 1) {
    const action = state.actions[i];
    if (action.type === "fill") {
      continue;
    }
    const bounds = getActionBounds(action);
    if (!bounds) {
      continue;
    }
    const overlaps = bounds.xMax >= xMin && bounds.xMin <= xMax && bounds.yMax >= yMin && bounds.yMin <= yMax;
    if (overlaps) {
      selected.push(action);
    }
  }

  state.selectedActionIds = selected.map((action) => action.id);
  state.selectedActionId = state.selectedActionIds[0] || null;
  if (state.selectedActionId) {
    loadStyleFromAction(selected[0]);
  }
  updateSelectionInfo();
  refreshGradientUi();
  updateCursor();
  renderOverlay();
}

function applyStyleToSelected() {
  const selectedIds = state.selectedActionIds.slice();
  if (selectedIds.length === 0) {
    updateSaveStatus("No object selected");
    return;
  }

  pushUndoSnapshot();
  state.redoStack = [];

  for (const selectedId of selectedIds) {
    const selectedIndex = findActionIndexById(selectedId);
    if (selectedIndex < 0) {
      continue;
    }

    const action = cloneAction(state.actions[selectedIndex]);
    if (action.type === "fill") {
      continue;
    }

    action.fillColor = state.color;
    action.strokeColor = state.strokeColor;
    action.opacity = state.opacity;

    if (action.type === "stroke") {
      action.size = state.size;
      if (action.brush !== "eraser") {
        action.color = state.color;
      }
    }

    if (
      action.type === "line" ||
      action.type === "rect" ||
      action.type === "ellipse" ||
      action.type === "arc" ||
      action.type === "polygonShape" ||
      action.type === "bezier" ||
      action.type === "vertexPath"
    ) {
      action.strokeWeight = state.strokeWeight;
    }

    if (action.type === "rect" || action.type === "ellipse" || action.type === "arc" || action.type === "polygonShape" || action.type === "vertexPath") {
      action.filled = state.fillShapes;
      action.gradient = (action.type === "vertexPath" || action.type === "arc")
        ? null
        : (state.gradientEnabled ? { endColor: state.gradientColor, direction: state.gradientDirection } : null);
    }

    if (action.type === "arc") {
      action.start = (clamp(state.arcStartDeg, 0, 360) * Math.PI) / 180;
      action.stop = (clamp(state.arcStopDeg, 0, 360) * Math.PI) / 180;
      action.filled = state.fillShapes;
      action.gradient = null;
    }

    state.actions[selectedIndex] = action;
  }
  refreshAfterActionChange("Style applied");
}

function deleteSelectedAction() {
  const selectedIds = new Set(state.selectedActionIds);
  if (selectedIds.size === 0) {
    updateSaveStatus("No object selected");
    return;
  }

  pushUndoSnapshot();
  state.redoStack = [];
  state.actions = state.actions.filter((action) => !selectedIds.has(action.id));
  state.selectedActionId = null;
  state.selectedActionIds = [];
  refreshAfterActionChange("Object deleted");
}

function bringSelectionForwardOne() {
  reorderSelectionByOne(true);
}

function sendSelectionBackwardOne() {
  reorderSelectionByOne(false);
}

function reorderSelectionByOne(forward) {
  const selectedSet = new Set(state.selectedActionIds);
  if (selectedSet.size === 0) {
    updateSaveStatus("No object selected");
    return;
  }

  const next = state.actions.slice();
  let changed = false;

  if (forward) {
    for (let i = next.length - 2; i >= 0; i -= 1) {
      const currSelected = selectedSet.has(next[i].id);
      const nextSelected = selectedSet.has(next[i + 1].id);
      if (currSelected && !nextSelected) {
        const tmp = next[i];
        next[i] = next[i + 1];
        next[i + 1] = tmp;
        changed = true;
      }
    }
  } else {
    for (let i = 1; i < next.length; i += 1) {
      const currSelected = selectedSet.has(next[i].id);
      const prevSelected = selectedSet.has(next[i - 1].id);
      if (currSelected && !prevSelected) {
        const tmp = next[i];
        next[i] = next[i - 1];
        next[i - 1] = tmp;
        changed = true;
      }
    }
  }

  if (!changed) {
    updateSaveStatus(forward ? "Already at front layer edge" : "Already at back layer edge");
    return;
  }

  pushUndoSnapshot();
  state.redoStack = [];
  state.actions = next;
  refreshAfterActionChange(forward ? "Moved forward one layer" : "Moved back one layer");
}

function bringSelectionToFront() {
  reorderSelectionAbsolute(true);
}

function sendSelectionToBack() {
  reorderSelectionAbsolute(false);
}

function reorderSelectionAbsolute(toFront) {
  const selectedSet = new Set(state.selectedActionIds);
  if (selectedSet.size === 0) {
    updateSaveStatus("No object selected");
    return;
  }

  const selected = state.actions.filter((action) => selectedSet.has(action.id));
  const unselected = state.actions.filter((action) => !selectedSet.has(action.id));
  const next = toFront ? [...unselected, ...selected] : [...selected, ...unselected];

  let changed = false;
  for (let i = 0; i < next.length; i += 1) {
    if (next[i] !== state.actions[i]) {
      changed = true;
      break;
    }
  }
  if (!changed) {
    updateSaveStatus(toFront ? "Already at front" : "Already at back");
    return;
  }

  pushUndoSnapshot();
  state.redoStack = [];
  state.actions = next;
  refreshAfterActionChange(toFront ? "Moved to front" : "Moved to back");
}

function copySelectedAction() {
  const selected = getSelectedAction();
  if (!selected) {
    updateSaveStatus("No object selected");
    return;
  }
  state.clipboardAction = cloneAction(selected);
  delete state.clipboardAction.id;
  updateSaveStatus("Object copied");
}

function pasteCopiedAction() {
  if (!state.clipboardAction) {
    updateSaveStatus("Clipboard empty");
    return;
  }

  const pasted = moveAction(cloneAction(state.clipboardAction), COPY_PASTE_OFFSET, COPY_PASTE_OFFSET);
  pasted.id = createActionId();

  pushUndoSnapshot();
  state.redoStack = [];
  state.actions.push(pasted);
  state.selectedActionId = pasted.id;
  state.selectedActionIds = [pasted.id];

  state.clipboardAction = cloneAction(pasted);
  delete state.clipboardAction.id;

  refreshAfterActionChange("Object pasted");
}

function commitNewAction(action, statusMessage) {
  const actionToAdd = cloneAction(action);
  if (!actionToAdd.id) {
    actionToAdd.id = createActionId();
  }

  pushUndoSnapshot();
  state.redoStack = [];
  state.actions.push(actionToAdd);
  state.selectedActionId = actionToAdd.id;
  state.selectedActionIds = [actionToAdd.id];
  refreshAfterActionChange(statusMessage);
}

function undo() {
  if (state.undoStack.length === 0) {
    return;
  }

  state.redoStack.push(cloneActions(state.actions));
  state.actions = state.undoStack.pop();
  ensureActionIds(state.actions);
  sanitizeSelection();
  refreshAfterActionChange("Undo applied");
}

function redo() {
  if (state.redoStack.length === 0) {
    return;
  }

  state.undoStack.push(cloneActions(state.actions));
  state.actions = state.redoStack.pop();
  ensureActionIds(state.actions);
  sanitizeSelection();
  refreshAfterActionChange("Redo applied");
}

function pushUndoSnapshot() {
  state.undoStack.push(cloneActions(state.actions));
  if (state.undoStack.length > MAX_HISTORY_STEPS) {
    state.undoStack.shift();
  }
}

function updateUndoRedoButtons() {
  undoBtn.disabled = state.undoStack.length === 0;
  redoBtn.disabled = state.redoStack.length === 0;
}

function refreshAfterActionChange(message, saveImmediately = true) {
  sanitizeSelection();
  redrawAll();
  renderOverlay();
  updateUndoRedoButtons();
  updateSelectionInfo();
  refreshGradientUi();
  updatePathInfo();
  updateCursor();
  updatePjsOutput();

  if (saveImmediately) {
    queueAutosave(true);
  }

  if (message) {
    updateSaveStatus(message);
  }
}

function sanitizeSelection() {
  if (!state.selectedActionId && state.selectedActionIds.length === 0) {
    return;
  }
  const existingIds = new Set(state.actions.map((action) => action.id));
  state.selectedActionIds = state.selectedActionIds.filter((id) => existingIds.has(id));
  if (state.selectedActionId && !existingIds.has(state.selectedActionId)) {
    state.selectedActionId = null;
  }
  if (!state.selectedActionId && state.selectedActionIds.length > 0) {
    state.selectedActionId = state.selectedActionIds[0];
  }
  if (state.selectedActionId && !state.selectedActionIds.includes(state.selectedActionId)) {
    state.selectedActionIds.unshift(state.selectedActionId);
  }
}

function getSelectedAction() {
  const index = findActionIndexById(state.selectedActionId);
  return index >= 0 ? state.actions[index] : null;
}

function getSelectedActions() {
  if (!Array.isArray(state.selectedActionIds) || state.selectedActionIds.length === 0) {
    return [];
  }
  const selected = [];
  for (const id of state.selectedActionIds) {
    const index = findActionIndexById(id);
    if (index >= 0) {
      selected.push(state.actions[index]);
    }
  }
  return selected;
}

function getSelectedBounds() {
  const selected = getSelectedActions();
  if (selected.length === 0) {
    return null;
  }
  let xMin = Infinity;
  let yMin = Infinity;
  let xMax = -Infinity;
  let yMax = -Infinity;
  for (const action of selected) {
    const bounds = getActionBounds(action);
    if (!bounds) {
      continue;
    }
    xMin = Math.min(xMin, bounds.xMin);
    yMin = Math.min(yMin, bounds.yMin);
    xMax = Math.max(xMax, bounds.xMax);
    yMax = Math.max(yMax, bounds.yMax);
  }
  if (!Number.isFinite(xMin) || !Number.isFinite(yMin) || !Number.isFinite(xMax) || !Number.isFinite(yMax)) {
    return null;
  }
  return { xMin, yMin, xMax, yMax };
}

function isPointInsideBounds(point, bounds) {
  return Boolean(bounds) &&
    point.x >= bounds.xMin &&
    point.x <= bounds.xMax &&
    point.y >= bounds.yMin &&
    point.y <= bounds.yMax;
}

function findActionIndexById(actionId) {
  if (!actionId) {
    return -1;
  }
  return state.actions.findIndex((action) => action.id === actionId);
}

function updateSelectionInfo() {
  const selectedActions = getSelectedActions();
  const action = getSelectedAction();
  const hasSelection = selectedActions.length > 0;

  applyStyleBtn.disabled = !hasSelection;
  deleteSelectedBtn.disabled = !hasSelection;
  bringForwardBtn.disabled = !hasSelection;
  sendBackwardBtn.disabled = !hasSelection;
  bringToFrontBtn.disabled = !hasSelection;
  sendToBackBtn.disabled = !hasSelection;

  if (!hasSelection) {
    selectionInfo.textContent = "No object selected.";
    return;
  }

  if (selectedActions.length > 1) {
    selectionInfo.textContent = `${selectedActions.length} objects selected`;
    return;
  }

  const bounds = getActionBounds(action);
  const width = bounds ? Math.round(bounds.xMax - bounds.xMin) : 0;
  const height = bounds ? Math.round(bounds.yMax - bounds.yMin) : 0;

  selectionInfo.textContent = `${actionLabel(action)} selected (${width}x${height})`;
}

function updatePathInfo() {
  const activePathTool = isPathTool(state.tool);
  const config = getPathToolConfig(state.tool);
  const draftPoints = state.pathDraft && Array.isArray(state.pathDraft.points) ? state.pathDraft.points.length : 0;

  finishPathBtn.disabled = !(activePathTool && config && draftPoints >= config.minPoints);
  cancelPathBtn.disabled = !(activePathTool && draftPoints > 0);

  if (!activePathTool) {
    pathInfo.textContent = "Path tool idle.";
    return;
  }

  if (!state.pathDraft || draftPoints === 0) {
    if (state.tool === "triangle") {
      pathInfo.textContent = "Triangle: click 3 points.";
    } else if (state.tool === "quad") {
      pathInfo.textContent = "Quad: click 4 points.";
    } else {
      pathInfo.textContent = "Click canvas to add points.";
    }
    return;
  }

  if (config && config.autoCompletePoints) {
    const left = Math.max(0, config.autoCompletePoints - draftPoints);
    pathInfo.textContent = left > 0
      ? `${draftPoints} point(s). ${left} more to finish.`
      : `${draftPoints} point(s).`;
    return;
  }

  pathInfo.textContent = `${draftPoints} point(s). Press Enter to finish, Esc to cancel.`;
}

function actionLabel(action) {
  if (action.type === "stroke") {
    return action.brush === "eraser" ? "Eraser stroke" : `${capitalize(action.brush)} stroke`;
  }
  if (action.type === "line") {
    return "Line";
  }
  if (action.type === "rect") {
    return action.filled ? "Filled rectangle" : "Rectangle";
  }
  if (action.type === "ellipse") {
    return action.filled ? "Filled ellipse" : "Ellipse";
  }
  if (action.type === "arc") {
    return action.filled ? "Filled arc" : "Arc";
  }
  if (action.type === "polygonShape") {
    return action.shape === "triangle" ? "Triangle" : "Quad";
  }
  if (action.type === "bezier") {
    return "Bezier curve";
  }
  if (action.type === "vertexPath") {
    return action.curved ? "Curve path" : "Vertex path";
  }
  if (action.type === "fill") {
    return "Bucket fill";
  }
  return "Object";
}

function renderOverlay() {
  clearPreview();

  const shapeTool = isDragShapeTool(state.tool);
  if (state.drawing && state.currentAction && shapeTool) {
    drawAction(previewCtx, state.currentAction, { preview: true });
  }

  if (isPathTool(state.tool) && state.pathDraft && state.pathDraft.points.length > 0) {
    const config = getPathToolConfig(state.tool);
    const draftPoints = state.hoverPoint
      ? [...state.pathDraft.points, state.hoverPoint]
      : [...state.pathDraft.points];
    if (config && config.type === "polygonShape") {
      const draftAction = {
        type: "polygonShape",
        shape: config.shape,
        points: draftPoints,
        strokeColor: state.strokeColor,
        strokeWeight: state.strokeWeight,
        opacity: state.opacity,
        filled: false
      };
      if (draftPoints.length >= 3) {
        drawPolygonShapeAction(previewCtx, draftAction, true);
      } else {
        drawVertexPathAction(previewCtx, { type: "vertexPath", curved: false, closed: false, points: draftPoints, strokeColor: state.strokeColor, strokeWeight: state.strokeWeight, opacity: state.opacity, filled: false }, true);
      }
    } else {
      const draftAction = {
        type: "vertexPath",
        curved: state.tool === "curve",
        closed: false,
        points: draftPoints,
        strokeColor: state.strokeColor,
        strokeWeight: state.strokeWeight,
        opacity: state.opacity,
        filled: false
      };
      drawVertexPathAction(previewCtx, draftAction, true);
    }
  }

  if (state.tool === "select") {
    if (state.drawing && state.dragMode === "marquee" && state.dragStart && state.dragMarqueeCurrent) {
      const x = Math.min(state.dragStart.x, state.dragMarqueeCurrent.x);
      const y = Math.min(state.dragStart.y, state.dragMarqueeCurrent.y);
      const w = Math.abs(state.dragMarqueeCurrent.x - state.dragStart.x);
      const h = Math.abs(state.dragMarqueeCurrent.y - state.dragStart.y);
      previewCtx.save();
      previewCtx.setLineDash([6, 4]);
      previewCtx.lineWidth = 1.4;
      previewCtx.strokeStyle = "rgba(0, 167, 225, 0.95)";
      previewCtx.fillStyle = "rgba(0, 167, 225, 0.12)";
      previewCtx.fillRect(x, y, w, h);
      previewCtx.strokeRect(x, y, w, h);
      previewCtx.restore();
    }
    const selectedActions = getSelectedActions();
    if (selectedActions.length > 1) {
      const groupBounds = getSelectedBounds();
      if (groupBounds) {
        drawSelectionBoundsOverlay(groupBounds, true);
      }
    } else {
      for (const action of selectedActions) {
        drawSelectionOverlay(action, action.id === state.selectedActionId);
      }
    }
  }
}

function drawSelectionBoundsOverlay(bounds, showHandles) {
  if (!bounds) {
    return;
  }
  const width = bounds.xMax - bounds.xMin;
  const height = bounds.yMax - bounds.yMin;
  const handle = getResizeHandleRect(bounds);

  previewCtx.save();
  previewCtx.setLineDash([8, 5]);
  previewCtx.lineWidth = 1.6;
  previewCtx.strokeStyle = "rgba(0, 167, 225, 0.98)";
  previewCtx.strokeRect(bounds.xMin, bounds.yMin, width, height);
  previewCtx.setLineDash([]);

  if (showHandles) {
    previewCtx.fillStyle = "rgba(255, 138, 61, 0.95)";
    previewCtx.fillRect(handle.x, handle.y, handle.size, handle.size);
    previewCtx.strokeStyle = "#ffffff";
    previewCtx.lineWidth = 1;
    previewCtx.strokeRect(handle.x, handle.y, handle.size, handle.size);

    const rotate = getRotateHandle(bounds);
    previewCtx.fillStyle = "rgba(0, 167, 225, 0.95)";
    previewCtx.beginPath();
    previewCtx.arc(rotate.x, rotate.y, rotate.size / 2, 0, TAU);
    previewCtx.fill();
    previewCtx.strokeStyle = "#ffffff";
    previewCtx.stroke();
  }

  previewCtx.restore();
}

function drawSelectionOverlay(action, isPrimary = true) {
  const bounds = getActionBounds(action);
  if (!bounds) {
    return;
  }

  const width = bounds.xMax - bounds.xMin;
  const height = bounds.yMax - bounds.yMin;
  const handle = getResizeHandleRect(bounds);

  previewCtx.save();
  previewCtx.setLineDash([8, 5]);
  previewCtx.lineWidth = 1.5;
  previewCtx.strokeStyle = isPrimary ? "rgba(0, 167, 225, 0.95)" : "rgba(0, 167, 225, 0.7)";
  previewCtx.strokeRect(bounds.xMin, bounds.yMin, width, height);
  if (!isPrimary) {
    previewCtx.restore();
    return;
  }
  previewCtx.setLineDash([]);
  previewCtx.fillStyle = "rgba(255, 138, 61, 0.95)";
  previewCtx.fillRect(handle.x, handle.y, handle.size, handle.size);
  previewCtx.strokeStyle = "#ffffff";
  previewCtx.lineWidth = 1;
  previewCtx.strokeRect(handle.x, handle.y, handle.size, handle.size);

  if (isActionRotatable(action)) {
    const rotate = getRotateHandle(bounds);
    previewCtx.fillStyle = "rgba(0, 167, 225, 0.95)";
    previewCtx.beginPath();
    previewCtx.arc(rotate.x, rotate.y, rotate.size / 2, 0, TAU);
    previewCtx.fill();
    previewCtx.strokeStyle = "#ffffff";
    previewCtx.stroke();
  }

  const edgeHandles = getEditableEdgeHandles(action);
  if (edgeHandles.length > 0) {
    previewCtx.fillStyle = "rgba(255, 138, 61, 0.95)";
    previewCtx.strokeStyle = "#ffffff";
    for (const edge of edgeHandles) {
      previewCtx.fillRect(edge.x - EDGE_HANDLE_SIZE / 2, edge.y - EDGE_HANDLE_SIZE / 2, EDGE_HANDLE_SIZE, EDGE_HANDLE_SIZE);
      previewCtx.strokeRect(edge.x - EDGE_HANDLE_SIZE / 2, edge.y - EDGE_HANDLE_SIZE / 2, EDGE_HANDLE_SIZE, EDGE_HANDLE_SIZE);
    }
  }

  const points = getEditablePoints(action);
  if (points.length > 0) {
    previewCtx.fillStyle = "rgba(10, 38, 66, 0.92)";
    previewCtx.strokeStyle = "#ffffff";
    for (const point of points) {
      previewCtx.beginPath();
      previewCtx.arc(point.x, point.y, 4.2, 0, TAU);
      previewCtx.fill();
      previewCtx.stroke();
    }
  }
  previewCtx.restore();
}

function getResizeHandleRect(bounds) {
  return {
    x: bounds.xMax - SELECT_HANDLE_SIZE / 2,
    y: bounds.yMax - SELECT_HANDLE_SIZE / 2,
    size: SELECT_HANDLE_SIZE
  };
}

function getRotateHandle(bounds) {
  return {
    x: bounds.xMax + 16,
    y: bounds.yMin - 10,
    size: ROTATE_HANDLE_SIZE
  };
}

function isPointOnResizeHandle(point, bounds) {
  const handle = getResizeHandleRect(bounds);
  return point.x >= handle.x &&
    point.x <= handle.x + handle.size &&
    point.y >= handle.y &&
    point.y <= handle.y + handle.size;
}

function isPointOnRotateHandle(point, bounds) {
  const rotate = getRotateHandle(bounds);
  return distance(point, { x: rotate.x, y: rotate.y }) <= rotate.size / 2 + 2;
}

function getEditableEdgeHandles(action) {
  if (!action) {
    return [];
  }

  if (action.type === "line") {
    return [{
      x: (action.x1 + action.x2) / 2,
      y: (action.y1 + action.y2) / 2,
      segment: { type: "line" }
    }];
  }

  if (action.type === "polygonShape" || action.type === "vertexPath") {
    const points = clonePointList(action.points);
    if (points.length < 2) {
      return [];
    }

    const closed = action.type === "polygonShape" || Boolean(action.closed);
    const edges = [];
    const end = closed ? points.length : points.length - 1;

    for (let i = 0; i < end; i += 1) {
      const a = i;
      const b = (i + 1) % points.length;
      edges.push({
        x: (points[a].x + points[b].x) / 2,
        y: (points[a].y + points[b].y) / 2,
        segment: { type: "points", a, b }
      });
    }

    return edges;
  }

  return [];
}

function findEdgeHandleHit(point, edges) {
  for (let i = 0; i < edges.length; i += 1) {
    if (distance(point, edges[i]) <= EDGE_HANDLE_SIZE * 0.9) {
      return i;
    }
  }
  return -1;
}

function findTopActionAtPoint(point) {
  for (let i = state.actions.length - 1; i >= 0; i -= 1) {
    const action = state.actions[i];
    if (action.type === "fill") {
      continue;
    }
    if (hitTestAction(action, point)) {
      return action;
    }
  }
  return null;
}

function hitTestAction(action, point) {
  switch (action.type) {
    case "stroke":
      return hitTestStroke(action, point);
    case "line":
      return distanceToSegment(point.x, point.y, action.x1, action.y1, action.x2, action.y2) <= Math.max(5, getShapeStrokeWeight(action) * 0.75);
    case "rect":
      return hitTestRect(action, point);
    case "ellipse":
      return hitTestEllipse(action, point);
    case "arc":
      return hitTestArc(action, point);
    case "polygonShape":
      return hitTestPolygon(action, point);
    case "bezier":
      return hitTestBezier(action, point);
    case "vertexPath":
      return hitTestVertexPath(action, point);
    default:
      return false;
  }
}

function hitTestStroke(action, point) {
  if (action.brush === "spray") {
    const dots = Array.isArray(action.dots) ? action.dots : [];
    const radius = Math.max(4, action.size * 0.6);
    for (let i = dots.length - 1; i >= 0; i -= 1) {
      if (distance(dots[i], point) <= radius) {
        return true;
      }
    }
    return false;
  }

  const points = Array.isArray(action.points) ? action.points : [];
  if (points.length === 0) {
    return false;
  }

  const tolerance = Math.max(5, action.size * 0.75);

  if (points.length === 1) {
    return distance(points[0], point) <= tolerance;
  }

  for (let i = 1; i < points.length; i += 1) {
    const from = points[i - 1];
    const to = points[i];
    if (distanceToSegment(point.x, point.y, from.x, from.y, to.x, to.y) <= tolerance) {
      return true;
    }
  }

  return false;
}

function hitTestRect(action, point) {
  const center = { x: action.x + action.w / 2, y: action.y + action.h / 2 };
  const local = toLocalRotatedPoint(point, center, getShapeRotation(action));
  const left = action.x;
  const top = action.y;
  const right = action.x + action.w;
  const bottom = action.y + action.h;
  const tolerance = Math.max(4, getShapeStrokeWeight(action) * 0.6);

  if (action.filled) {
    return local.x >= left && local.x <= right && local.y >= top && local.y <= bottom;
  }

  const onVertical = local.y >= top - tolerance && local.y <= bottom + tolerance &&
    (Math.abs(local.x - left) <= tolerance || Math.abs(local.x - right) <= tolerance);
  const onHorizontal = local.x >= left - tolerance && local.x <= right + tolerance &&
    (Math.abs(local.y - top) <= tolerance || Math.abs(local.y - bottom) <= tolerance);

  return onVertical || onHorizontal;
}

function hitTestEllipse(action, point) {
  const center = { x: action.x + action.w / 2, y: action.y + action.h / 2 };
  const local = toLocalRotatedPoint(point, center, getShapeRotation(action));
  const rx = Math.max(1, action.w / 2);
  const ry = Math.max(1, action.h / 2);
  const cx = action.x + rx;
  const cy = action.y + ry;

  const nx = (local.x - cx) / rx;
  const ny = (local.y - cy) / ry;
  const value = nx * nx + ny * ny;

  if (action.filled) {
    return value <= 1;
  }

  const edgeTolerance = Math.max(0.08, (getShapeStrokeWeight(action) / Math.max(action.w, action.h)) * 2.2);
  return value >= 1 - edgeTolerance && value <= 1 + edgeTolerance;
}

function hitTestArc(action, point) {
  const center = getArcCenter(action);
  const local = toLocalRotatedPoint(point, center, getShapeRotation(action));
  const cx = center.x;
  const cy = center.y;
  const rx = Math.max(1, Math.abs(action.w) / 2);
  const ry = Math.max(1, Math.abs(action.h) / 2);
  const start = typeof action.start === "number" ? action.start : 0;
  const stop = typeof action.stop === "number" ? action.stop : Math.PI;

  if (action.filled) {
    const nx = (local.x - cx) / rx;
    const ny = (local.y - cy) / ry;
    if (nx * nx + ny * ny > 1) {
      return false;
    }
    const theta = Math.atan2(ny, nx);
    return isAngleWithinSweep(theta, start, stop);
  }

  const tol = Math.max(5, getShapeStrokeWeight(action) * 0.9);
  let prev = sampleArcPoint(action, 0);
  for (let i = 1; i <= 48; i += 1) {
    const p = sampleArcPoint(action, i / 48);
    if (distanceToSegment(local.x, local.y, prev.x, prev.y, p.x, p.y) <= tol) {
      return true;
    }
    prev = p;
  }
  // Small assist only for genuinely tiny arcs.
  if (rx <= 14 && ry <= 14) {
    return distance(local, { x: cx, y: cy }) <= 3.5;
  }
  return false;
}

function isAngleWithinSweep(theta, start, stop) {
  const normalizedStart = normalizeAngle(start);
  let normalizedStop = normalizeAngle(stop);
  while (normalizedStop < normalizedStart) {
    normalizedStop += TAU;
  }
  const sweep = normalizedStop - normalizedStart;
  if (sweep >= TAU - 0.0001) {
    return true;
  }
  let normalizedTheta = normalizeAngle(theta);
  while (normalizedTheta < normalizedStart) {
    normalizedTheta += TAU;
  }
  return normalizedTheta <= normalizedStop;
}

function hitTestPolygon(action, point) {
  const points = getPolygonShapePoints(action);
  return hitTestPointInPolyOrEdge(points, point, action.filled, getShapeStrokeWeight(action));
}

function hitTestVertexPath(action, point) {
  const points = Array.isArray(action.points) ? action.points : [];
  if (points.length < 2) {
    return false;
  }
  if (action.closed && action.filled && isPointInPolygon(points, point)) {
    return true;
  }
  return hitTestPolyline(points, point, getShapeStrokeWeight(action), action.closed);
}

function hitTestBezier(action, point) {
  const tol = Math.max(5, getShapeStrokeWeight(action) * 0.8);
  let prev = { x: action.x1, y: action.y1 };
  for (let i = 1; i <= 32; i += 1) {
    const t = i / 32;
    const p = sampleBezier(action, t);
    if (distanceToSegment(point.x, point.y, prev.x, prev.y, p.x, p.y) <= tol) {
      return true;
    }
    prev = p;
  }
  return false;
}

function getActionBounds(action) {
  if (action.type === "line") {
    const margin = Math.max(2, getShapeStrokeWeight(action) * 0.6);
    return {
      xMin: Math.min(action.x1, action.x2) - margin,
      yMin: Math.min(action.y1, action.y2) - margin,
      xMax: Math.max(action.x1, action.x2) + margin,
      yMax: Math.max(action.y1, action.y2) + margin
    };
  }

  if (action.type === "rect" || action.type === "ellipse" || action.type === "arc") {
    const margin = Math.max(2, getShapeStrokeWeight(action) * 0.6);
    const rotation = getShapeRotation(action);
    if (rotation) {
      const cx = action.x + action.w / 2;
      const cy = action.y + action.h / 2;
      const corners = [
        rotatePoint({ x: action.x, y: action.y }, { x: cx, y: cy }, rotation),
        rotatePoint({ x: action.x + action.w, y: action.y }, { x: cx, y: cy }, rotation),
        rotatePoint({ x: action.x + action.w, y: action.y + action.h }, { x: cx, y: cy }, rotation),
        rotatePoint({ x: action.x, y: action.y + action.h }, { x: cx, y: cy }, rotation)
      ];
      return getBoundsFromPoints(corners, margin);
    }
    return {
      xMin: action.x - margin,
      yMin: action.y - margin,
      xMax: action.x + action.w + margin,
      yMax: action.y + action.h + margin
    };
  }

  if (action.type === "stroke") {
    const list = action.brush === "spray" ? action.dots : action.points;
    if (!Array.isArray(list) || list.length === 0) {
      return null;
    }

    let xMin = list[0].x;
    let yMin = list[0].y;
    let xMax = list[0].x;
    let yMax = list[0].y;

    for (let i = 1; i < list.length; i += 1) {
      const point = list[i];
      xMin = Math.min(xMin, point.x);
      yMin = Math.min(yMin, point.y);
      xMax = Math.max(xMax, point.x);
      yMax = Math.max(yMax, point.y);
    }

    const margin = Math.max(2, action.size * 0.9);
    return {
      xMin: xMin - margin,
      yMin: yMin - margin,
      xMax: xMax + margin,
      yMax: yMax + margin
    };
  }

  if (action.type === "polygonShape") {
    const pts = getPolygonShapePoints(action);
    return getBoundsFromPoints(pts, Math.max(2, getShapeStrokeWeight(action) * 0.6));
  }

  if (action.type === "vertexPath") {
    return getBoundsFromPoints(action.points, Math.max(2, getShapeStrokeWeight(action) * 0.6));
  }

  if (action.type === "bezier") {
    const samples = [];
    for (let i = 0; i <= 40; i += 1) {
      samples.push(sampleBezier(action, i / 40));
    }
    return getBoundsFromPoints(samples, Math.max(2, getShapeStrokeWeight(action) * 0.6));
  }

  return null;
}

function getEditablePoints(action) {
  if (!action) {
    return [];
  }
  if (action.type === "line") {
    return [{ x: action.x1, y: action.y1 }, { x: action.x2, y: action.y2 }];
  }
  if (action.type === "bezier") {
    return [
      { x: action.x1, y: action.y1 },
      { x: action.cx1, y: action.cy1 },
      { x: action.cx2, y: action.cy2 },
      { x: action.x2, y: action.y2 }
    ];
  }
  if (action.type === "arc") {
    const start = typeof action.start === "number" ? action.start : 0;
    const stop = typeof action.stop === "number" ? action.stop : Math.PI;
    return [
      arcPointAtAngle(action, start, true),
      arcPointAtAngle(action, stop, true)
    ];
  }

  if (action.type === "polygonShape" || action.type === "vertexPath") {
    return clonePointList(action.points);
  }
  return [];
}

function findControlPointHit(point, controls) {
  for (let i = 0; i < controls.length; i += 1) {
    if (distance(point, controls[i]) <= 7) {
      return i;
    }
  }
  return -1;
}

function moveActionControlPoint(action, index, point) {
  const updated = cloneAction(action);
  if (updated.type === "line") {
    if (index === 0) {
      updated.x1 = point.x;
      updated.y1 = point.y;
    } else if (index === 1) {
      updated.x2 = point.x;
      updated.y2 = point.y;
    }
    return updated;
  }
  if (updated.type === "bezier") {
    if (index === 0) { updated.x1 = point.x; updated.y1 = point.y; }
    if (index === 1) { updated.cx1 = point.x; updated.cy1 = point.y; }
    if (index === 2) { updated.cx2 = point.x; updated.cy2 = point.y; }
    if (index === 3) { updated.x2 = point.x; updated.y2 = point.y; }
    return updated;
  }
  if ((updated.type === "polygonShape" || updated.type === "vertexPath") && Array.isArray(updated.points) && updated.points[index]) {
    updated.points[index] = { x: point.x, y: point.y };
    return updated;
  }
  if (updated.type === "arc" && (index === 0 || index === 1)) {
    const center = getArcCenter(updated);
    const local = toLocalRotatedPoint(point, center, getShapeRotation(updated));
    const rx = Math.max(1, Math.abs(updated.w) / 2);
    const ry = Math.max(1, Math.abs(updated.h) / 2);
    const dx = local.x - center.x;
    const dy = local.y - center.y;
    const ang = Math.atan2(dy / ry, dx / rx);
    if (index === 0) {
      updated.start = normalizeAngle(ang);
    } else {
      updated.stop = normalizeAngle(ang);
    }
    return updated;
  }
  return updated;
}

function moveActionEdge(action, edgeInfo, dx, dy) {
  const updated = cloneAction(action);
  if (!edgeInfo || !edgeInfo.segment) {
    return updated;
  }

  if (edgeInfo.segment.type === "line" && updated.type === "line") {
    updated.x1 += dx;
    updated.y1 += dy;
    updated.x2 += dx;
    updated.y2 += dy;
    return updated;
  }

  if (edgeInfo.segment.type === "points" && Array.isArray(updated.points)) {
    const a = edgeInfo.segment.a;
    const b = edgeInfo.segment.b;
    if (updated.points[a]) {
      updated.points[a] = { x: updated.points[a].x + dx, y: updated.points[a].y + dy };
    }
    if (updated.points[b]) {
      updated.points[b] = { x: updated.points[b].x + dx, y: updated.points[b].y + dy };
    }
  }

  return updated;
}

function isActionRotatable(action) {
  return action && action.type !== "fill";
}

function rotateAction(action, center, angleDelta) {
  const updated = cloneAction(action);
  if (!isActionRotatable(updated)) {
    return updated;
  }

  if (updated.type === "line") {
    const p1 = rotatePoint({ x: updated.x1, y: updated.y1 }, center, angleDelta);
    const p2 = rotatePoint({ x: updated.x2, y: updated.y2 }, center, angleDelta);
    updated.x1 = p1.x; updated.y1 = p1.y;
    updated.x2 = p2.x; updated.y2 = p2.y;
    return updated;
  }

  if (updated.type === "bezier") {
    const p1 = rotatePoint({ x: updated.x1, y: updated.y1 }, center, angleDelta);
    const c1 = rotatePoint({ x: updated.cx1, y: updated.cy1 }, center, angleDelta);
    const c2 = rotatePoint({ x: updated.cx2, y: updated.cy2 }, center, angleDelta);
    const p2 = rotatePoint({ x: updated.x2, y: updated.y2 }, center, angleDelta);
    updated.x1 = p1.x; updated.y1 = p1.y;
    updated.cx1 = c1.x; updated.cy1 = c1.y;
    updated.cx2 = c2.x; updated.cy2 = c2.y;
    updated.x2 = p2.x; updated.y2 = p2.y;
    return updated;
  }

  if (updated.type === "rect" || updated.type === "ellipse" || updated.type === "arc") {
    const shapeCenter = {
      x: updated.x + updated.w / 2,
      y: updated.y + updated.h / 2
    };
    const rotatedCenter = rotatePoint(shapeCenter, center, angleDelta);
    updated.x = rotatedCenter.x - updated.w / 2;
    updated.y = rotatedCenter.y - updated.h / 2;
    updated.rotation = normalizeAngle((updated.rotation || 0) + angleDelta);
    return updated;
  }

  if (updated.type === "stroke") {
    if (Array.isArray(updated.points)) {
      updated.points = updated.points.map((point) => rotatePoint(point, center, angleDelta));
    }
    if (Array.isArray(updated.dots)) {
      updated.dots = updated.dots.map((dot) => rotatePoint(dot, center, angleDelta));
    }
    return updated;
  }

  if (Array.isArray(updated.points)) {
    updated.points = updated.points.map((point) => rotatePoint(point, center, angleDelta));
  }
  return updated;
}

function moveAction(action, dx, dy) {
  const updated = cloneAction(action);

  if (updated.type === "line") {
    updated.x1 += dx;
    updated.y1 += dy;
    updated.x2 += dx;
    updated.y2 += dy;
    return updated;
  }

  if (updated.type === "rect" || updated.type === "ellipse" || updated.type === "arc") {
    updated.x += dx;
    updated.y += dy;
    return updated;
  }

  if (updated.type === "polygonShape") {
    updated.points = clonePointList(updated.points).map((p) => ({ x: p.x + dx, y: p.y + dy }));
    return updated;
  }

  if (updated.type === "bezier") {
    updated.x1 += dx; updated.y1 += dy;
    updated.cx1 += dx; updated.cy1 += dy;
    updated.cx2 += dx; updated.cy2 += dy;
    updated.x2 += dx; updated.y2 += dy;
    return updated;
  }

  if (updated.type === "vertexPath") {
    updated.points = clonePointList(updated.points).map((p) => ({ x: p.x + dx, y: p.y + dy }));
    return updated;
  }

  if (updated.type === "stroke") {
    if (Array.isArray(updated.points)) {
      updated.points = updated.points.map((point) => ({ x: point.x + dx, y: point.y + dy }));
    }
    if (Array.isArray(updated.dots)) {
      updated.dots = updated.dots.map((dot) => ({ x: dot.x + dx, y: dot.y + dy }));
    }
  }

  return updated;
}

function scaleActionFrom(action, originX, originY, scaleX, scaleY) {
  const updated = cloneAction(action);

  if (updated.type === "line") {
    updated.x1 = originX + (updated.x1 - originX) * scaleX;
    updated.y1 = originY + (updated.y1 - originY) * scaleY;
    updated.x2 = originX + (updated.x2 - originX) * scaleX;
    updated.y2 = originY + (updated.y2 - originY) * scaleY;
    updated.strokeWeight = Math.max(1, getShapeStrokeWeight(updated) * ((scaleX + scaleY) / 2));
    return updated;
  }

  if (updated.type === "rect" || updated.type === "ellipse" || updated.type === "arc") {
    updated.x = originX + (updated.x - originX) * scaleX;
    updated.y = originY + (updated.y - originY) * scaleY;
    updated.w = Math.max(2, updated.w * scaleX);
    updated.h = Math.max(2, updated.h * scaleY);
    updated.strokeWeight = Math.max(1, getShapeStrokeWeight(updated) * ((scaleX + scaleY) / 2));
    return updated;
  }

  if (updated.type === "polygonShape") {
    updated.points = clonePointList(updated.points).map((p) => ({
      x: originX + (p.x - originX) * scaleX,
      y: originY + (p.y - originY) * scaleY
    }));
    updated.strokeWeight = Math.max(1, getShapeStrokeWeight(updated) * ((scaleX + scaleY) / 2));
    return updated;
  }

  if (updated.type === "bezier") {
    updated.x1 = originX + (updated.x1 - originX) * scaleX;
    updated.y1 = originY + (updated.y1 - originY) * scaleY;
    updated.cx1 = originX + (updated.cx1 - originX) * scaleX;
    updated.cy1 = originY + (updated.cy1 - originY) * scaleY;
    updated.cx2 = originX + (updated.cx2 - originX) * scaleX;
    updated.cy2 = originY + (updated.cy2 - originY) * scaleY;
    updated.x2 = originX + (updated.x2 - originX) * scaleX;
    updated.y2 = originY + (updated.y2 - originY) * scaleY;
    updated.strokeWeight = Math.max(1, getShapeStrokeWeight(updated) * ((scaleX + scaleY) / 2));
    return updated;
  }

  if (updated.type === "vertexPath") {
    updated.points = clonePointList(updated.points).map((p) => ({
      x: originX + (p.x - originX) * scaleX,
      y: originY + (p.y - originY) * scaleY
    }));
    updated.strokeWeight = Math.max(1, getShapeStrokeWeight(updated) * ((scaleX + scaleY) / 2));
    return updated;
  }

  if (updated.type === "stroke") {
    if (Array.isArray(updated.points)) {
      updated.points = updated.points.map((point) => ({
        x: originX + (point.x - originX) * scaleX,
        y: originY + (point.y - originY) * scaleY
      }));
    }

    if (Array.isArray(updated.dots)) {
      updated.dots = updated.dots.map((dot) => ({
        x: originX + (dot.x - originX) * scaleX,
        y: originY + (dot.y - originY) * scaleY
      }));
    }

    updated.size = Math.max(1, updated.size * ((scaleX + scaleY) / 2));
  }

  return updated;
}

function buildShapeAction(tool, start, end) {
  if (tool === "line") {
    return {
      id: createActionId(),
      type: "line",
      x1: start.x,
      y1: start.y,
      x2: end.x,
      y2: end.y,
      strokeColor: state.strokeColor,
      strokeWeight: state.strokeWeight,
      opacity: state.opacity
    };
  }

  if (tool === "bezier") {
    const dx = end.x - start.x;
    const dy = end.y - start.y;
    const lift = Math.max(20, distance(start, end) * 0.32);
    const perpX = dy === 0 ? 0 : -dy / Math.max(1, Math.abs(dy));
    const perpY = dx === 0 ? -1 : dx / Math.max(1, Math.abs(dx));

    return {
      id: createActionId(),
      type: "bezier",
      x1: start.x,
      y1: start.y,
      x2: end.x,
      y2: end.y,
      cx1: start.x + dx * 0.28 + perpX * lift,
      cy1: start.y + dy * 0.28 + perpY * lift,
      cx2: start.x + dx * 0.72 + perpX * lift,
      cy2: start.y + dy * 0.72 + perpY * lift,
      strokeColor: state.strokeColor,
      strokeWeight: state.strokeWeight,
      opacity: state.opacity
    };
  }

  if (tool === "arc") {
    return {
      id: createActionId(),
      type: "arc",
      x: Math.min(start.x, end.x),
      y: Math.min(start.y, end.y),
      w: Math.abs(end.x - start.x),
      h: Math.abs(end.y - start.y),
      start: (clamp(state.arcStartDeg, 0, 360) * Math.PI) / 180,
      stop: (clamp(state.arcStopDeg, 0, 360) * Math.PI) / 180,
      fillColor: state.color,
      strokeColor: state.strokeColor,
      strokeWeight: state.strokeWeight,
      opacity: state.opacity,
      filled: state.fillShapes,
      gradient: null
    };
  }

  const shape = {
    id: createActionId(),
    type: tool === "triangle" || tool === "quad" ? "polygonShape" : tool,
    x: Math.min(start.x, end.x),
    y: Math.min(start.y, end.y),
    w: Math.abs(end.x - start.x),
    h: Math.abs(end.y - start.y),
    fillColor: state.color,
    strokeColor: state.strokeColor,
    strokeWeight: state.strokeWeight,
    opacity: state.opacity,
    filled: state.fillShapes,
    gradient: state.gradientEnabled
      ? { endColor: state.gradientColor, direction: state.gradientDirection }
      : null
  };

  if (tool === "triangle" || tool === "quad") {
    shape.shape = tool;
    shape.points = buildPolygonPointsFromBounds(tool, start, end);
  }

  return shape;
}

function redrawAll() {
  resetCanvas();
  for (const action of state.actions) {
    drawAction(ctx, action);
  }
}

function resetCanvas() {
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalCompositeOperation = "source-over";
  ctx.globalAlpha = 1;
  ctx.shadowBlur = 0;
  ctx.clearRect(0, 0, drawCanvas.width, drawCanvas.height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, drawCanvas.width, drawCanvas.height);
  ctx.restore();
}

function drawAction(targetCtx, action, options = {}) {
  switch (action.type) {
    case "stroke":
      drawStrokeAction(targetCtx, action);
      return true;
    case "line":
      drawLineAction(targetCtx, action, options.preview);
      return true;
    case "rect":
      drawRectAction(targetCtx, action, options.preview);
      return true;
    case "ellipse":
      drawEllipseAction(targetCtx, action, options.preview);
      return true;
    case "arc":
      drawArcAction(targetCtx, action, options.preview);
      return true;
    case "polygonShape":
      drawPolygonShapeAction(targetCtx, action, options.preview);
      return true;
    case "bezier":
      drawBezierAction(targetCtx, action, options.preview);
      return true;
    case "vertexPath":
      drawVertexPathAction(targetCtx, action, options.preview);
      return true;
    case "fill":
      return floodFill(targetCtx, action.x, action.y, action.color, action.opacity, action.tolerance);
    default:
      return false;
  }
}

function drawStrokeAction(targetCtx, action) {
  if (action.brush === "spray") {
    const dots = Array.isArray(action.dots) ? action.dots : [];
    for (const dot of dots) {
      drawSprayDot(targetCtx, action, dot);
    }
    return;
  }

  const points = Array.isArray(action.points) ? action.points : [];
  if (points.length === 0) {
    return;
  }

  targetCtx.save();
  applyStrokeStyle(targetCtx, action);

  if (points.length === 1) {
    const single = points[0];
    targetCtx.beginPath();
    targetCtx.arc(single.x, single.y, Math.max(1, targetCtx.lineWidth * 0.48), 0, TAU);
    targetCtx.fillStyle = targetCtx.strokeStyle;
    targetCtx.fill();
    targetCtx.restore();
    return;
  }

  targetCtx.beginPath();
  targetCtx.moveTo(points[0].x, points[0].y);
  for (let i = 1; i < points.length; i += 1) {
    targetCtx.lineTo(points[i].x, points[i].y);
  }
  targetCtx.stroke();
  targetCtx.restore();
}

function drawStrokeSegment(targetCtx, action, from, to) {
  targetCtx.save();
  applyStrokeStyle(targetCtx, action);
  targetCtx.beginPath();
  targetCtx.moveTo(from.x, from.y);
  targetCtx.lineTo(to.x, to.y);
  targetCtx.stroke();
  targetCtx.restore();
}

function drawStrokeDot(targetCtx, action, point) {
  targetCtx.save();
  applyStrokeStyle(targetCtx, action);
  targetCtx.beginPath();
  targetCtx.arc(point.x, point.y, Math.max(1, targetCtx.lineWidth * 0.48), 0, TAU);
  targetCtx.fillStyle = targetCtx.strokeStyle;
  targetCtx.fill();
  targetCtx.restore();
}

function applyStrokeStyle(targetCtx, action) {
  const isEraser = action.brush === "eraser";
  const baseColor = isEraser ? "#ffffff" : action.color;

  targetCtx.globalCompositeOperation = "source-over";
  targetCtx.lineCap = "round";
  targetCtx.lineJoin = "round";
  targetCtx.strokeStyle = rgbaString(baseColor, action.opacity);
  targetCtx.lineWidth = action.brush === "pencil" ? Math.max(1, action.size * 0.55) : action.size;

  if (action.brush === "glow") {
    targetCtx.shadowColor = rgbaString(action.color, clamp(action.opacity + 0.05, 0, 1));
    targetCtx.shadowBlur = action.size * 1.8;
  } else {
    targetCtx.shadowBlur = 0;
  }
}

function drawLineAction(targetCtx, action, preview = false) {
  targetCtx.save();
  shapeStrokeStyle(targetCtx, action, preview);
  targetCtx.beginPath();
  targetCtx.moveTo(action.x1, action.y1);
  targetCtx.lineTo(action.x2, action.y2);
  targetCtx.stroke();
  targetCtx.restore();
}

function drawRectAction(targetCtx, action, preview = false) {
  targetCtx.save();
  shapeStrokeStyle(targetCtx, action, preview);
  const rotation = getShapeRotation(action);
  const cx = action.x + action.w / 2;
  const cy = action.y + action.h / 2;
  if (rotation) {
    targetCtx.translate(cx, cy);
    targetCtx.rotate(rotation);
    targetCtx.translate(-cx, -cy);
  }

  if (action.filled) {
    targetCtx.fillStyle = getShapeFillStyle(targetCtx, action, preview);
    targetCtx.fillRect(action.x, action.y, action.w, action.h);
  }

  targetCtx.strokeRect(action.x, action.y, action.w, action.h);
  targetCtx.restore();
}

function drawEllipseAction(targetCtx, action, preview = false) {
  targetCtx.save();
  shapeStrokeStyle(targetCtx, action, preview);
  const rotation = getShapeRotation(action);
  const cx = action.x + action.w / 2;
  const cy = action.y + action.h / 2;
  if (rotation) {
    targetCtx.translate(cx, cy);
    targetCtx.rotate(rotation);
    targetCtx.translate(-cx, -cy);
  }
  targetCtx.beginPath();
  targetCtx.ellipse(
    action.x + action.w / 2,
    action.y + action.h / 2,
    Math.max(1, Math.abs(action.w) / 2),
    Math.max(1, Math.abs(action.h) / 2),
    0,
    0,
    TAU
  );

  if (action.filled) {
    targetCtx.fillStyle = getShapeFillStyle(targetCtx, action, preview);
    targetCtx.fill();
  }

  targetCtx.stroke();
  targetCtx.restore();
}

function sampleArcPoint(action, t) {
  const start = typeof action.start === "number" ? action.start : 0;
  const stop = typeof action.stop === "number" ? action.stop : Math.PI;
  let stopAdjusted = stop;
  while (stopAdjusted < start) {
    stopAdjusted += TAU;
  }
  const theta = start + (stopAdjusted - start) * t;
  return arcPointAtAngle(action, theta, false);
}

function drawArcAction(targetCtx, action, preview = false) {
  targetCtx.save();
  shapeStrokeStyle(targetCtx, action, preview);
  targetCtx.lineCap = "butt";
  const center = getArcCenter(action);
  const rotation = getShapeRotation(action);
  if (rotation) {
    targetCtx.translate(center.x, center.y);
    targetCtx.rotate(rotation);
    targetCtx.translate(-center.x, -center.y);
  }
  const cx = action.x + action.w / 2;
  const cy = action.y + action.h / 2;
  const rx = Math.max(1, Math.abs(action.w) / 2);
  const ry = Math.max(1, Math.abs(action.h) / 2);
  const start = typeof action.start === "number" ? action.start : 0;
  const stop = typeof action.stop === "number" ? action.stop : Math.PI;

  targetCtx.beginPath();
  if (action.filled) {
    // Match Khan/Processing arc fill behavior (PIE-style sector).
    targetCtx.moveTo(cx, cy);
  }
  targetCtx.ellipse(cx, cy, rx, ry, 0, start, stop);
  if (action.filled) {
    targetCtx.closePath();
    targetCtx.fillStyle = getShapeFillStyle(targetCtx, action, preview);
    targetCtx.fill();
  }
  targetCtx.beginPath();
  targetCtx.ellipse(cx, cy, rx, ry, 0, start, stop);
  targetCtx.stroke();
  targetCtx.restore();
}

function drawPolygonShapeAction(targetCtx, action, preview = false) {
  const pts = getPolygonShapePoints(action);
  drawClosedPolygon(targetCtx, pts, action, preview);
}

function drawBezierAction(targetCtx, action, preview = false) {
  targetCtx.save();
  shapeStrokeStyle(targetCtx, action, preview);
  targetCtx.beginPath();
  targetCtx.moveTo(action.x1, action.y1);
  targetCtx.bezierCurveTo(action.cx1, action.cy1, action.cx2, action.cy2, action.x2, action.y2);
  targetCtx.stroke();
  targetCtx.restore();
}

function drawVertexPathAction(targetCtx, action, preview = false) {
  const points = Array.isArray(action.points) ? action.points : [];
  if (points.length < 2) {
    return;
  }

  targetCtx.save();
  shapeStrokeStyle(targetCtx, action, preview);

  if (action.closed) {
    targetCtx.beginPath();
    drawPathGeometry(targetCtx, points, action.curved, true);
    if (action.filled) {
      targetCtx.fillStyle = getShapeFillStyle(targetCtx, action, preview);
      targetCtx.fill();
    }
    targetCtx.stroke();
    targetCtx.restore();
    return;
  }

  targetCtx.beginPath();
  drawPathGeometry(targetCtx, points, action.curved, false);
  targetCtx.stroke();
  targetCtx.restore();
}

function drawClosedPolygon(targetCtx, points, action, preview) {
  if (!Array.isArray(points) || points.length < 3) {
    return;
  }
  targetCtx.save();
  shapeStrokeStyle(targetCtx, action, preview);
  targetCtx.beginPath();
  targetCtx.moveTo(points[0].x, points[0].y);
  for (let i = 1; i < points.length; i += 1) {
    targetCtx.lineTo(points[i].x, points[i].y);
  }
  targetCtx.closePath();
  if (action.filled) {
    targetCtx.fillStyle = getShapeFillStyle(targetCtx, action, preview);
    targetCtx.fill();
  }
  targetCtx.stroke();
  targetCtx.restore();
}

function drawPathGeometry(targetCtx, points, curved, closed) {
  if (!curved) {
    targetCtx.moveTo(points[0].x, points[0].y);
    for (let i = 1; i < points.length; i += 1) {
      targetCtx.lineTo(points[i].x, points[i].y);
    }
    if (closed) {
      targetCtx.closePath();
    }
    return;
  }

  if (points.length < 3) {
    targetCtx.moveTo(points[0].x, points[0].y);
    targetCtx.lineTo(points[1].x, points[1].y);
    return;
  }

  // PJS-like curveVertex behavior (Catmull-Rom spline, default tension).
  const local = closed
    ? [points[points.length - 1], ...points, points[0], points[1]]
    : [points[0], ...points, points[points.length - 1]];

  let started = false;
  for (let i = 0; i <= local.length - 4; i += 1) {
    const p0 = local[i];
    const p1 = local[i + 1];
    const p2 = local[i + 2];
    const p3 = local[i + 3];

    const c1x = p1.x + (p2.x - p0.x) / 6;
    const c1y = p1.y + (p2.y - p0.y) / 6;
    const c2x = p2.x - (p3.x - p1.x) / 6;
    const c2y = p2.y - (p3.y - p1.y) / 6;

    if (!started) {
      targetCtx.moveTo(p1.x, p1.y);
      started = true;
    }
    targetCtx.bezierCurveTo(c1x, c1y, c2x, c2y, p2.x, p2.y);
  }

  if (closed) {
    targetCtx.closePath();
  }
}

function getPolygonShapePoints(action) {
  if (Array.isArray(action.points) && action.points.length >= 3) {
    return clonePointList(action.points);
  }

  const x = action.x;
  const y = action.y;
  const w = action.w;
  const h = action.h;

  if (action.shape === "triangle") {
    return [
      { x: x + w / 2, y },
      { x, y: y + h },
      { x: x + w, y: y + h }
    ];
  }

  return [
    { x: x + w * 0.18, y: y + h * 0.16 },
    { x: x + w * 0.82, y: y + h * 0.12 },
    { x: x + w, y: y + h * 0.86 },
    { x, y: y + h * 0.9 }
  ];
}

function buildPolygonPointsFromBounds(shape, start, end) {
  const x = Math.min(start.x, end.x);
  const y = Math.min(start.y, end.y);
  const w = Math.abs(end.x - start.x);
  const h = Math.abs(end.y - start.y);

  if (shape === "triangle") {
    return [
      { x: x + w / 2, y },
      { x, y: y + h },
      { x: x + w, y: y + h }
    ];
  }

  return [
    { x: x + w * 0.18, y: y + h * 0.16 },
    { x: x + w * 0.82, y: y + h * 0.12 },
    { x: x + w, y: y + h * 0.86 },
    { x, y: y + h * 0.9 }
  ];
}

function shapeStrokeStyle(targetCtx, action, preview) {
  targetCtx.globalCompositeOperation = "source-over";
  const strokeColor = action.strokeColor || action.color || "#0d2238";
  targetCtx.strokeStyle = rgbaString(strokeColor, action.opacity);
  targetCtx.lineWidth = getShapeStrokeWeight(action);
  targetCtx.lineJoin = "round";
  targetCtx.lineCap = "round";
  if (preview) {
    targetCtx.setLineDash([8, 5]);
  } else {
    targetCtx.setLineDash([]);
  }
}

function getShapeStrokeWeight(action) {
  if (typeof action.strokeWeight === "number") {
    return action.strokeWeight;
  }
  if (typeof action.size === "number") {
    return action.size;
  }
  return 2;
}

function getShapeFillStyle(targetCtx, action, preview) {
  const baseColor = action.fillColor || action.color || "#00b7ff";
  const alpha = action.opacity * (preview ? 0.4 : 1);

  if (!action.gradient || !action.gradient.endColor) {
    return rgbaString(baseColor, alpha);
  }

  const direction = normalizeGradientDirection(action.gradient.direction);
  const bounds = getActionFillBounds(action);
  const x1 = bounds ? bounds.xMin : (action.x || 0);
  const y1 = bounds ? bounds.yMin : (action.y || 0);
  const width = bounds ? (bounds.xMax - bounds.xMin) : (action.w || 1);
  const height = bounds ? (bounds.yMax - bounds.yMin) : (action.h || 1);
  const cx = x1 + width / 2;
  const cy = y1 + height / 2;
  let gradient;

  if (direction === "center-out" || direction === "center-in") {
    const radius = Math.max(1, Math.max(width, height) * 0.5);
    gradient = targetCtx.createRadialGradient(cx, cy, 0, cx, cy, radius);
  } else if (direction === "right-left") {
    gradient = targetCtx.createLinearGradient(x1 + width, y1, x1, y1);
  } else if (direction === "top-bottom") {
    gradient = targetCtx.createLinearGradient(x1, y1, x1, y1 + height);
  } else if (direction === "bottom-top") {
    gradient = targetCtx.createLinearGradient(x1, y1 + height, x1, y1);
  } else {
    gradient = targetCtx.createLinearGradient(x1, y1, x1 + width, y1);
  }

  if (direction === "center-in") {
    gradient.addColorStop(0, rgbaString(action.gradient.endColor, alpha));
    gradient.addColorStop(1, rgbaString(baseColor, alpha));
  } else {
    gradient.addColorStop(0, rgbaString(baseColor, alpha));
    gradient.addColorStop(1, rgbaString(action.gradient.endColor, alpha));
  }
  return gradient;
}

function getActionFillBounds(action) {
  if (action.type === "rect" || action.type === "ellipse") {
    return { xMin: action.x, yMin: action.y, xMax: action.x + action.w, yMax: action.y + action.h };
  }
  if (action.type === "polygonShape") {
    return getBoundsFromPoints(getPolygonShapePoints(action), 0);
  }
  if (action.type === "vertexPath") {
    return getBoundsFromPoints(action.points, 0);
  }
  return getActionBounds(action);
}

function sprayAt(point, action) {
  const radius = Math.max(5, action.size * 2.1);
  const density = Math.max(10, Math.round(action.size * 2.5));

  for (let i = 0; i < density; i += 1) {
    const angle = Math.random() * TAU;
    const distanceFromCenter = Math.sqrt(Math.random()) * radius;
    const x = point.x + Math.cos(angle) * distanceFromCenter;
    const y = point.y + Math.sin(angle) * distanceFromCenter;

    if (x < 0 || x >= drawCanvas.width || y < 0 || y >= drawCanvas.height) {
      continue;
    }

    const dot = { x, y };
    action.dots.push(dot);
    drawSprayDot(ctx, action, dot);
  }
}

function drawSprayDot(targetCtx, action, dot) {
  const isEraser = action.brush === "eraser";
  const baseColor = isEraser ? "#ffffff" : action.color;
  const alpha = action.opacity * 0.67;
  const size = Math.max(1, action.size * 0.25);

  targetCtx.save();
  targetCtx.globalCompositeOperation = "source-over";
  targetCtx.fillStyle = rgbaString(baseColor, alpha);
  targetCtx.beginPath();
  targetCtx.arc(dot.x, dot.y, size, 0, TAU);
  targetCtx.fill();
  targetCtx.restore();
}

function floodFill(targetCtx, x, y, hexColor, opacity, tolerance) {
  const width = drawCanvas.width;
  const height = drawCanvas.height;
  const startX = clamp(Math.floor(x), 0, width - 1);
  const startY = clamp(Math.floor(y), 0, height - 1);

  const imageData = targetCtx.getImageData(0, 0, width, height);
  const data = imageData.data;

  const startIndex = (startY * width + startX) * 4;
  const targetR = data[startIndex];
  const targetG = data[startIndex + 1];
  const targetB = data[startIndex + 2];
  const targetA = data[startIndex + 3];

  const [fillR, fillG, fillB] = hexToRgb(hexColor);
  const fillA = Math.round(clamp(opacity, 0, 1) * 255);
  const tol = Math.max(0, Number(tolerance) || 0);
  const colorTolSq = tol * tol * 3;
  const alphaTol = Math.max(10, tol * 1.35);

  const isMatch = (offset) => {
    const dr = data[offset] - targetR;
    const dg = data[offset + 1] - targetG;
    const db = data[offset + 2] - targetB;
    const da = Math.abs(data[offset + 3] - targetA);
    return (dr * dr + dg * dg + db * db) <= colorTolSq && da <= alphaTol;
  };

  const seedColorDiff = (
    (targetR - fillR) * (targetR - fillR) +
    (targetG - fillG) * (targetG - fillG) +
    (targetB - fillB) * (targetB - fillB)
  );
  if (seedColorDiff <= 3 && Math.abs(targetA - fillA) <= 2) {
    return false;
  }

  const visited = new Uint8Array(width * height);
  const stack = [startX, startY];
  let changed = false;

  while (stack.length > 0) {
    const py = stack.pop();
    let px = stack.pop();
    let xLeft = px;

    while (xLeft >= 0) {
      const leftIdx = py * width + xLeft;
      if (visited[leftIdx]) {
        break;
      }
      const leftOff = leftIdx * 4;
      if (!isMatch(leftOff)) {
        break;
      }
      xLeft -= 1;
    }
    xLeft += 1;

    let spanUp = false;
    let spanDown = false;

    for (let xScan = xLeft; xScan < width; xScan += 1) {
      const idx = py * width + xScan;
      if (visited[idx]) {
        break;
      }
      const off = idx * 4;
      if (!isMatch(off)) {
        break;
      }

      visited[idx] = 1;
      data[off] = fillR;
      data[off + 1] = fillG;
      data[off + 2] = fillB;
      data[off + 3] = fillA;
      changed = true;

      if (py > 0) {
        const upIdx = idx - width;
        const upOff = upIdx * 4;
        const upMatch = !visited[upIdx] && isMatch(upOff);
        if (upMatch && !spanUp) {
          stack.push(xScan, py - 1);
          spanUp = true;
        } else if (!upMatch) {
          spanUp = false;
        }
      }

      if (py < height - 1) {
        const downIdx = idx + width;
        const downOff = downIdx * 4;
        const downMatch = !visited[downIdx] && isMatch(downOff);
        if (downMatch && !spanDown) {
          stack.push(xScan, py + 1);
          spanDown = true;
        } else if (!downMatch) {
          spanDown = false;
        }
      }
    }
  }

  if (!changed) {
    return false;
  }

  targetCtx.putImageData(imageData, 0, 0);
  return true;
}

function clearPreview() {
  previewCtx.clearRect(0, 0, previewCanvas.width, previewCanvas.height);
}

function getCanvasPoint(event) {
  const rect = drawCanvas.getBoundingClientRect();
  const scaleX = drawCanvas.width / rect.width;
  const scaleY = drawCanvas.height / rect.height;
  return {
    x: clamp((event.clientX - rect.left) * scaleX, 0, drawCanvas.width),
    y: clamp((event.clientY - rect.top) * scaleY, 0, drawCanvas.height)
  };
}

function setCanvasSize(width, height) {
  state.canvasWidth = clamp(Math.round(width), MIN_CANVAS_SIZE, MAX_CANVAS_SIZE);
  state.canvasHeight = clamp(Math.round(height), MIN_CANVAS_SIZE, MAX_CANVAS_SIZE);

  drawCanvas.width = state.canvasWidth;
  drawCanvas.height = state.canvasHeight;
  previewCanvas.width = state.canvasWidth;
  previewCanvas.height = state.canvasHeight;

  canvasStage.style.setProperty("--canvas-aspect", `${state.canvasWidth} / ${state.canvasHeight}`);
  canvasWidthInput.value = String(state.canvasWidth);
  canvasHeightInput.value = String(state.canvasHeight);
  updateCanvasInfo();
}

function updateCanvasInfo() {
  canvasInfo.textContent = `Canvas: ${state.canvasWidth}x${state.canvasHeight}`;
}

function queueAutosave(includeStatus = false) {
  if (state.autosaveTimer) {
    clearTimeout(state.autosaveTimer);
  }

  state.autosaveTimer = setTimeout(() => {
    saveToStorage(includeStatus);
  }, AUTOSAVE_DELAY_MS);
}

function saveToStorage(showStatus) {
  if (state.autosaveTimer) {
    clearTimeout(state.autosaveTimer);
    state.autosaveTimer = null;
  }

  const payload = {
    version: 2,
    savedAt: new Date().toISOString(),
    settings: {
      tool: state.tool,
      color: state.color,
      size: state.size,
      strokeColor: state.strokeColor,
      strokeWeight: state.strokeWeight,
      opacity: state.opacity,
      fillShapes: state.fillShapes,
      gradientEnabled: state.gradientEnabled,
      gradientColor: state.gradientColor,
      gradientDirection: state.gradientDirection,
      exportImageMode: state.exportImageMode,
      canvasWidth: state.canvasWidth,
      canvasHeight: state.canvasHeight
    },
    actions: state.actions
  };

  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    if (showStatus) {
      updateSaveStatus(`Saved ${formatTime(payload.savedAt)}`);
    } else {
      updateSaveStatus(`Autosaved ${formatTime(payload.savedAt)}`);
    }
  } catch (error) {
    updateSaveStatus("Save failed: localStorage full", true);
  }
}

function restoreFromStorage(showStatus) {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    if (showStatus) {
      updateSaveStatus("No saved project found");
    }
    return false;
  }

  try {
    const parsed = JSON.parse(raw);
    const settings = parsed.settings && typeof parsed.settings === "object" ? parsed.settings : {};

    if (typeof settings.canvasWidth === "number" && typeof settings.canvasHeight === "number") {
      setCanvasSize(settings.canvasWidth, settings.canvasHeight);
    }

    if (typeof settings.tool === "string") {
      state.tool = settings.tool;
    }
    if (typeof settings.color === "string") {
      state.color = settings.color;
    }
    if (typeof settings.size === "number") {
      state.size = clamp(settings.size, 1, 60);
    }
    if (typeof settings.strokeColor === "string") {
      state.strokeColor = settings.strokeColor;
    }
    if (typeof settings.strokeWeight === "number") {
      state.strokeWeight = clamp(settings.strokeWeight, 1, 40);
    }
    if (typeof settings.arcStartDeg === "number") {
      state.arcStartDeg = clamp(Math.round(settings.arcStartDeg), 0, 360);
    }
    if (typeof settings.arcStopDeg === "number") {
      state.arcStopDeg = clamp(Math.round(settings.arcStopDeg), 0, 360);
    }
    if (typeof settings.opacity === "number") {
      state.opacity = clamp(settings.opacity, 0.05, 1);
    }
    if (typeof settings.fillShapes === "boolean") {
      state.fillShapes = settings.fillShapes;
    }
    if (typeof settings.gradientEnabled === "boolean") {
      state.gradientEnabled = settings.gradientEnabled;
    }
    if (typeof settings.gradientColor === "string") {
      state.gradientColor = settings.gradientColor;
    }
    state.gradientDirection = normalizeGradientDirection(settings.gradientDirection);
    if (typeof settings.exportImageMode === "boolean") {
      state.exportImageMode = settings.exportImageMode;
    }

    state.actions = Array.isArray(parsed.actions) ? parsed.actions : [];
    ensureActionIds(state.actions);
    state.undoStack = [];
    state.redoStack = [];
    state.selectedActionId = null;
    state.selectedActionIds = [];

    redrawAll();
    renderOverlay();
    syncControlsFromState();
    highlightActiveTool();
    updateUndoRedoButtons();
    updateSelectionInfo();
    refreshGradientUi();
    updatePathInfo();
    updateCursor();
    updatePjsOutput();

    if (showStatus) {
      updateSaveStatus(parsed.savedAt ? `Restored ${formatTime(parsed.savedAt)}` : "Project restored");
    }

    return true;
  } catch (error) {
    updateSaveStatus("Saved project is corrupted", true);
    return false;
  }
}

function updateSaveStatus(message, isError = false) {
  saveStatus.textContent = message;
  saveStatus.style.color = isError ? "#b12020" : "#2f506d";
}

function updatePjsOutput() {
  pjsOutput.value = generatePjsCode();
}

function generatePjsCode() {
  if (state.exportImageMode) {
    return generatePjsImageCode();
  }

  const encoded = [];
  let skippedFillCount = 0;

  for (const action of state.actions) {
    if (action.type === "fill") {
      skippedFillCount += 1;
      continue;
    }

    const encodedAction = encodeAction(action);
    if (encodedAction) {
      encoded.push(encodedAction);
    }
  }

  const lines = [];
  lines.push("// Generated by Graphics Designer");
  lines.push(`var ART_DATA=${JSON.stringify(encoded)};`);
  lines.push("");
  lines.push("var glowLine=function(x1,y1,x2,y2,r,g,b,a,w){");
  lines.push("    noFill();w=max(0.1,w);");
  lines.push("    for(var i=4;i>=1;i--){");
  lines.push("        stroke(r,g,b,constrain(a/(i*1.8),0,255));strokeWeight(w+i*1.9);line(x1,y1,x2,y2);");
  lines.push("    }");
  lines.push("    stroke(r,g,b,constrain(a,0,255));strokeWeight(w);line(x1,y1,x2,y2);");
  lines.push("};");
  lines.push("var glowDot=function(x,y,r,g,b,a,w){");
  lines.push("    noStroke();");
  lines.push("    for(var i=4;i>=1;i--){fill(r,g,b,constrain(a/(i*1.8),0,255));ellipse(x,y,w+i*2.2,w+i*2.2);}");
  lines.push("    fill(r,g,b,constrain(a,0,255));ellipse(x,y,max(1,w),max(1,w));");
  lines.push("};");
  lines.push("var gradientTriangleHQ=function(x1,y1,x2,y2,x3,y3,c1,c2,mode){");
  lines.push("    mode=(mode||\"center-out\").toLowerCase();");
  lines.push("    var xMin=max(0,floor(min(x1,min(x2,x3)))),xMax=min(width-1,ceil(max(x1,max(x2,x3))));");
  lines.push("    var yMin=max(0,floor(min(y1,min(y2,y3)))),yMax=min(height-1,ceil(max(y1,max(y2,y3))));");
  lines.push("    var den=((y2-y3)*(x1-x3)+(x3-x2)*(y1-y3));");
  lines.push("    if(abs(den)<0.000001){return;}");
  lines.push("    loadPixels();");
  lines.push("    for(var y=yMin;y<=yMax;y++){");
  lines.push("        for(var x=xMin;x<=xMax;x++){");
  lines.push("            var w1=((y2-y3)*(x-x3)+(x3-x2)*(y-y3))/den;");
  lines.push("            var w2=((y3-y1)*(x-x3)+(x1-x3)*(y-y3))/den;");
  lines.push("            var w3=1-w1-w2;");
  lines.push("            if(w1>=0&&w2>=0&&w3>=0){");
  lines.push("                var t=0;");
  lines.push("                if(mode===\"left-right\"||mode===\"horizontal\"){t=(x-xMin)/max(1,(xMax-xMin));}");
  lines.push("                else if(mode===\"right-left\"){t=1-(x-xMin)/max(1,(xMax-xMin));}");
  lines.push("                else if(mode===\"top-bottom\"||mode===\"vertical\"||mode===\"up-down\"){t=(y-yMin)/max(1,(yMax-yMin));}");
  lines.push("                else if(mode===\"bottom-top\"||mode===\"down-up\"){t=1-(y-yMin)/max(1,(yMax-yMin));}");
  lines.push("                else{var minW=min(w1,min(w2,w3));var centerOut=constrain(1-3*minW,0,1);t=(mode===\"center-in\")?(1-centerOut):centerOut;}");
  lines.push("                set(x,y,lerpColor(c1,c2,constrain(t,0,1)));");
  lines.push("            }");
  lines.push("        }");
  lines.push("    }");
  lines.push("    updatePixels();");
  lines.push("};");
  lines.push("var _inTri=function(px,py,ax,ay,bx,by,cx,cy){");
  lines.push("    var den=((by-cy)*(ax-cx)+(cx-bx)*(ay-cy));");
  lines.push("    if(abs(den)<0.000001){return false;}");
  lines.push("    var w1=((by-cy)*(px-cx)+(cx-bx)*(py-cy))/den;");
  lines.push("    var w2=((cy-ay)*(px-cx)+(ax-cx)*(py-cy))/den;");
  lines.push("    var w3=1-w1-w2;");
  lines.push("    return (w1>=0&&w2>=0&&w3>=0);");
  lines.push("};");
  lines.push("var _inPoly4=function(px,py,x1,y1,x2,y2,x3,y3,x4,y4){");
  lines.push("    var inside=false;");
  lines.push("    var xs=[x1,x2,x3,x4],ys=[y1,y2,y3,y4];");
  lines.push("    for(var i=0,j=3;i<4;j=i++){");
  lines.push("        var yi=ys[i],yj=ys[j],xi=xs[i],xj=xs[j];");
  lines.push("        var cross=((yi>py)!==(yj>py))&&(px<((xj-xi)*(py-yi))/((yj-yi)+0.000001)+xi);");
  lines.push("        if(cross){inside=!inside;}");
  lines.push("    }");
  lines.push("    return inside;");
  lines.push("};");
  lines.push("var modeName=function(m){");
  lines.push("    if(m===1){return \"right-left\";}");
  lines.push("    if(m===2){return \"top-bottom\";}");
  lines.push("    if(m===3){return \"bottom-top\";}");
  lines.push("    if(m===4){return \"center-out\";}");
  lines.push("    if(m===5){return \"center-in\";}");
  lines.push("    return \"left-right\";");
  lines.push("};");
  lines.push("var gradientQuadHQ=function(x1,y1,x2,y2,x3,y3,x4,y4,c1,c2,mode){");
  lines.push("    mode=(mode||\"center-out\").toLowerCase();");
  lines.push("    var xMin=max(0,floor(min(min(x1,x2),min(x3,x4)))),xMax=min(width-1,ceil(max(max(x1,x2),max(x3,x4))));");
  lines.push("    var yMin=max(0,floor(min(min(y1,y2),min(y3,y4)))),yMax=min(height-1,ceil(max(max(y1,y2),max(y3,y4))));");
  lines.push("    var cx=(x1+x2+x3+x4)/4,cy=(y1+y2+y3+y4)/4,maxD=1;");
  lines.push("    maxD=max(maxD,dist(cx,cy,x1,y1));maxD=max(maxD,dist(cx,cy,x2,y2));");
  lines.push("    maxD=max(maxD,dist(cx,cy,x3,y3));maxD=max(maxD,dist(cx,cy,x4,y4));");
  lines.push("    loadPixels();");
  lines.push("    for(var y=yMin;y<=yMax;y++){");
  lines.push("        for(var x=xMin;x<=xMax;x++){");
  lines.push("            var inside=_inPoly4(x,y,x1,y1,x2,y2,x3,y3,x4,y4);");
  lines.push("            if(inside){");
  lines.push("                var t=0;");
  lines.push("                if(mode===\"left-right\"||mode===\"horizontal\"){t=(x-xMin)/max(1,(xMax-xMin));}");
  lines.push("                else if(mode===\"right-left\"){t=1-(x-xMin)/max(1,(xMax-xMin));}");
  lines.push("                else if(mode===\"top-bottom\"||mode===\"vertical\"||mode===\"up-down\"){t=(y-yMin)/max(1,(yMax-yMin));}");
  lines.push("                else if(mode===\"bottom-top\"||mode===\"down-up\"){t=1-(y-yMin)/max(1,(yMax-yMin));}");
  lines.push("                else{var d=dist(x,y,cx,cy)/maxD;d=constrain(d,0,1);t=(mode===\"center-in\")?(1-d):d;}");
  lines.push("                set(x,y,lerpColor(c1,c2,constrain(t,0,1)));");
  lines.push("            }");
  lines.push("        }");
  lines.push("    }");
  lines.push("    updatePixels();");
  lines.push("};");
  lines.push("var gradientRectHQ=function(x,y,w,h,c1,c2,mode,cx,cy,rot){");
  lines.push("    mode=(mode||\"left-right\").toLowerCase();");
  lines.push("    if(w<=0||h<=0){return;}");
  lines.push("    cx=(typeof cx===\"number\")?cx:(x+w/2);");
  lines.push("    cy=(typeof cy===\"number\")?cy:(y+h/2);");
  lines.push("    rot=rot||0;");
  lines.push("    var cr=cos(-rot),sr=sin(-rot);");
  lines.push("    var p1={x:x,y:y},p2={x:x+w,y:y},p3={x:x+w,y:y+h},p4={x:x,y:y+h};");
  lines.push("    var rp=function(p){var dx=p.x-cx,dy=p.y-cy;return {x:cx+dx*cos(rot)-dy*sin(rot),y:cy+dx*sin(rot)+dy*cos(rot)};};");
  lines.push("    var q1=rp(p1),q2=rp(p2),q3=rp(p3),q4=rp(p4);");
  lines.push("    var xMin=max(0,floor(min(min(q1.x,q2.x),min(q3.x,q4.x)))),xMax=min(width-1,ceil(max(max(q1.x,q2.x),max(q3.x,q4.x))));");
  lines.push("    var yMin=max(0,floor(min(min(q1.y,q2.y),min(q3.y,q4.y)))),yMax=min(height-1,ceil(max(max(q1.y,q2.y),max(q3.y,q4.y))));");
  lines.push("    var lcX=x+w/2,lcY=y+h/2,maxD=1;");
  lines.push("    maxD=max(maxD,dist(lcX,lcY,x,y));maxD=max(maxD,dist(lcX,lcY,x+w,y));");
  lines.push("    maxD=max(maxD,dist(lcX,lcY,x+w,y+h));maxD=max(maxD,dist(lcX,lcY,x,y+h));");
  lines.push("    loadPixels();");
  lines.push("    for(var py=yMin;py<=yMax;py++){");
  lines.push("        for(var px=xMin;px<=xMax;px++){");
  lines.push("            var dx=px-cx,dy=py-cy;");
  lines.push("            var lx=cx+dx*cr-dy*sr,ly=cy+dx*sr+dy*cr;");
  lines.push("            if(lx<x||lx>x+w||ly<y||ly>y+h){continue;}");
  lines.push("            var t=0;");
  lines.push("            if(mode===\"left-right\"||mode===\"horizontal\"){t=(lx-x)/max(1,w);}");
  lines.push("            else if(mode===\"right-left\"){t=1-(lx-x)/max(1,w);}");
  lines.push("            else if(mode===\"top-bottom\"||mode===\"vertical\"||mode===\"up-down\"){t=(ly-y)/max(1,h);}");
  lines.push("            else if(mode===\"bottom-top\"||mode===\"down-up\"){t=1-(ly-y)/max(1,h);}");
  lines.push("            else{var d=dist(lx,ly,lcX,lcY)/maxD;d=constrain(d,0,1);t=(mode===\"center-in\")?(1-d):d;}");
  lines.push("            set(px,py,lerpColor(c1,c2,constrain(t,0,1)));");
  lines.push("        }");
  lines.push("    }");
  lines.push("    updatePixels();");
  lines.push("};");
  lines.push("var gradientEllipseHQ=function(cx,cy,w,h,c1,c2,mode){");
  lines.push("    mode=(mode||\"left-right\").toLowerCase();");
  lines.push("    var a=w/2,b=h/2;");
  lines.push("    if(a<=0||b<=0){return;}");
  lines.push("    var xMin=max(0,floor(cx-a)),xMax=min(width-1,ceil(cx+a));");
  lines.push("    var yMin=max(0,floor(cy-b)),yMax=min(height-1,ceil(cy+b));");
  lines.push("    loadPixels();");
  lines.push("    for(var py=yMin;py<=yMax;py++){");
  lines.push("        for(var px=xMin;px<=xMax;px++){");
  lines.push("            var nx=(px-cx)/a,ny=(py-cy)/b,r2=nx*nx+ny*ny;");
  lines.push("            if(r2<=1){");
  lines.push("                var t=0;");
  lines.push("                if(mode===\"left-right\"||mode===\"horizontal\"){t=(nx+1)*0.5;}");
  lines.push("                else if(mode===\"right-left\"){t=1-(nx+1)*0.5;}");
  lines.push("                else if(mode===\"top-bottom\"||mode===\"vertical\"||mode===\"up-down\"){t=(ny+1)*0.5;}");
  lines.push("                else if(mode===\"bottom-top\"||mode===\"down-up\"){t=1-(ny+1)*0.5;}");
  lines.push("                else{var d=sqrt(r2);t=(mode===\"center-in\")?(1-d):d;}");
  lines.push("                set(px,py,lerpColor(c1,c2,constrain(t,0,1)));");
  lines.push("            }");
  lines.push("        }");
  lines.push("    }");
  lines.push("    updatePixels();");
  lines.push("};");
  lines.push("");

  lines.push("var drawArt=function(x,y,w,h){");
  lines.push("    var __prevAngleMode=(typeof angleMode!==\"undefined\")?angleMode:\"degrees\";");
  lines.push("    angleMode=\"radians\";");
  lines.push("    var wa=max(0.001,(abs(w)+abs(h))*0.5);");
  lines.push("    var sx=function(nx){return x+nx*w;};");
  lines.push("    var sy=function(ny){return y+ny*h;};");
  lines.push("    var sw=function(nw){return abs(nw*w);};");
  lines.push("    var sh=function(nh){return abs(nh*h);};");
  lines.push("    var ss=function(ns){return max(0.001,ns*wa);};");
  lines.push("    strokeCap(ROUND);strokeJoin(ROUND);");
  lines.push("    for(var i=0;i<ART_DATA.length;i++){");
  lines.push("        var a=ART_DATA[i],t=a[0],p,j;");
  lines.push("        if(t===\"S\"){");
  lines.push("            stroke(a[1],a[2],a[3],a[4]);strokeWeight(ss(a[5]));p=a[6];");
  lines.push("            if(p.length===2){point(sx(p[0]),sy(p[1]));continue;}");
  lines.push("            for(j=2;j<p.length;j+=2){line(sx(p[j-2]),sy(p[j-1]),sx(p[j]),sy(p[j+1]));}");
  lines.push("        }else if(t===\"G\"){");
  lines.push("            p=a[6];");
  lines.push("            if(p.length===2){glowDot(sx(p[0]),sy(p[1]),a[1],a[2],a[3],a[4],ss(a[5]));continue;}");
  lines.push("            for(j=2;j<p.length;j+=2){glowLine(sx(p[j-2]),sy(p[j-1]),sx(p[j]),sy(p[j+1]),a[1],a[2],a[3],a[4],ss(a[5]));}");
  lines.push("        }else if(t===\"P\"){");
  lines.push("            stroke(a[1],a[2],a[3],a[4]);strokeWeight(ss(a[5]));p=a[6];");
  lines.push("            for(j=0;j<p.length;j+=2){point(sx(p[j]),sy(p[j+1]));}");
  lines.push("        }else if(t===\"L\"){");
  lines.push("            stroke(a[1],a[2],a[3],a[4]);strokeWeight(ss(a[5]));line(sx(a[6]),sy(a[7]),sx(a[8]),sy(a[9]));");
  lines.push("        }else if(t===\"A\"){");
  lines.push("            if(a[12]){pushMatrix();translate(sx(a[6]),sy(a[7]));rotate((a[12]));translate(-sx(a[6]),-sy(a[7]));}");
  lines.push("            if(a[13]){fill(a[14],a[15],a[16],a[4]);}else{noFill();}");
  lines.push("            stroke(a[1],a[2],a[3],a[4]);strokeWeight(ss(a[5]));");
  lines.push("            arc(sx(a[6]),sy(a[7]),sw(a[8]),sh(a[9]),(a[10]),(a[11]));");
  lines.push("            if(a[12]){popMatrix();}");
  lines.push("        }else if(t===\"R\"){");
  lines.push("            if(a[14]){pushMatrix();translate(sx(a[10])+sw(a[12])/2,sy(a[11])+sh(a[13])/2);rotate((a[14]));translate(-(sx(a[10])+sw(a[12])/2),-(sy(a[11])+sh(a[13])/2));}");
  lines.push("            if(a[9]){fill(a[1],a[2],a[3],a[7]);}else{noFill();}");
  lines.push("            stroke(a[4],a[5],a[6],a[7]);strokeWeight(ss(a[8]));rect(sx(a[10]),sy(a[11]),sw(a[12]),sh(a[13]));");
  lines.push("            if(a[14]){popMatrix();}");
  lines.push("        }else if(t===\"RG\"){");
  lines.push("            if(a[13]){gradientRectHQ(sx(a[14]),sy(a[15]),sw(a[16]),sh(a[17]),color(a[1],a[2],a[3],a[11]),color(a[4],a[5],a[6],a[11]),modeName(a[7]),sx(a[14])+sw(a[16])/2,sy(a[15])+sh(a[17])/2,a[18]||0);}");
  lines.push("            if(a[18]){pushMatrix();translate(sx(a[14])+sw(a[16])/2,sy(a[15])+sh(a[17])/2);rotate(a[18]);translate(-(sx(a[14])+sw(a[16])/2),-(sy(a[15])+sh(a[17])/2));}");
  lines.push("            stroke(a[8],a[9],a[10],a[11]);strokeWeight(ss(a[12]));");
  lines.push("            noFill();");
  lines.push("            rect(sx(a[14]),sy(a[15]),sw(a[16]),sh(a[17]));");
  lines.push("            if(a[18]){popMatrix();}");
  lines.push("        }else if(t===\"TG\"){");
  lines.push("            if(a[13]){gradientTriangleHQ(sx(a[14]),sy(a[15]),sx(a[16]),sy(a[17]),sx(a[18]),sy(a[19]),color(a[1],a[2],a[3],a[11]),color(a[4],a[5],a[6],a[11]),modeName(a[7]));}");
  lines.push("            stroke(a[8],a[9],a[10],a[11]);strokeWeight(ss(a[12]));");
  lines.push("            noFill();");
  lines.push("            triangle(sx(a[14]),sy(a[15]),sx(a[16]),sy(a[17]),sx(a[18]),sy(a[19]));");
  lines.push("        }else if(t===\"QG\"){");
  lines.push("            if(a[13]){gradientQuadHQ(sx(a[14]),sy(a[15]),sx(a[16]),sy(a[17]),sx(a[18]),sy(a[19]),sx(a[20]),sy(a[21]),color(a[1],a[2],a[3],a[11]),color(a[4],a[5],a[6],a[11]),modeName(a[7]));}");
  lines.push("            stroke(a[8],a[9],a[10],a[11]);strokeWeight(ss(a[12]));");
  lines.push("            noFill();");
  lines.push("            quad(sx(a[14]),sy(a[15]),sx(a[16]),sy(a[17]),sx(a[18]),sy(a[19]),sx(a[20]),sy(a[21]));");
  lines.push("        }else if(t===\"E\"){");
  lines.push("            if(a[14]){pushMatrix();translate(sx(a[10]),sy(a[11]));rotate((a[14]));translate(-sx(a[10]),-sy(a[11]));}");
  lines.push("            if(a[9]){fill(a[1],a[2],a[3],a[7]);}else{noFill();}");
  lines.push("            stroke(a[4],a[5],a[6],a[7]);strokeWeight(ss(a[8]));ellipse(sx(a[10]),sy(a[11]),sw(a[12]),sh(a[13]));");
  lines.push("            if(a[14]){popMatrix();}");
  lines.push("        }else if(t===\"EG\"){");
  lines.push("            if(a[13]){gradientEllipseHQ(sx(a[14]),sy(a[15]),sw(a[16]),sh(a[17]),color(a[1],a[2],a[3],a[11]),color(a[4],a[5],a[6],a[11]),modeName(a[7]));}");
  lines.push("            if(a[18]){pushMatrix();translate(sx(a[14]),sy(a[15]));rotate((a[18]));translate(-sx(a[14]),-sy(a[15]));}");
  lines.push("            stroke(a[8],a[9],a[10],a[11]);strokeWeight(ss(a[12]));");
  lines.push("            noFill();");
  lines.push("            ellipse(sx(a[14]),sy(a[15]),sw(a[16]),sh(a[17]));");
  lines.push("            if(a[18]){popMatrix();}");
  lines.push("        }else if(t===\"B\"){");
  lines.push("            stroke(a[1],a[2],a[3],a[4]);strokeWeight(ss(a[5]));noFill();");
  lines.push("            bezier(sx(a[6]),sy(a[7]),sx(a[8]),sy(a[9]),sx(a[10]),sy(a[11]),sx(a[12]),sy(a[13]));");
  lines.push("        }else if(t===\"VP\"){");
  lines.push("            if(a[11]){fill(a[1],a[2],a[3],a[7]);}else{noFill();}");
  lines.push("            stroke(a[4],a[5],a[6],a[7]);strokeWeight(ss(a[8]));");
  lines.push("            p=a[12];");
  lines.push("            if(a[9]){");
  lines.push("                beginShape();");
  lines.push("                if(a[10] && p.length>=6){");
  lines.push("                    curveVertex(sx(p[p.length-2]),sy(p[p.length-1]));");
  lines.push("                    for(j=0;j<p.length;j+=2){curveVertex(sx(p[j]),sy(p[j+1]));}");
  lines.push("                    curveVertex(sx(p[0]),sy(p[1]));");
  lines.push("                    curveVertex(sx(p[2]),sy(p[3]));");
  lines.push("                    endShape(CLOSE);");
  lines.push("                }else{");
  lines.push("                    curveVertex(sx(p[0]),sy(p[1]));");
  lines.push("                    for(j=0;j<p.length;j+=2){curveVertex(sx(p[j]),sy(p[j+1]));}");
  lines.push("                    curveVertex(sx(p[p.length-2]),sy(p[p.length-1]));");
  lines.push("                    endShape();");
  lines.push("                }");
  lines.push("            }else{");
  lines.push("                beginShape();");
  lines.push("                for(j=0;j<p.length;j+=2){vertex(sx(p[j]),sy(p[j+1]));}");
  lines.push("                if(a[10]){endShape(CLOSE);}else{endShape();}");
  lines.push("            }");
  lines.push("        }");
  lines.push("    }");
  lines.push("    angleMode=__prevAngleMode;");
  lines.push("};");
  lines.push("");
  lines.push("// Example: drawArt(50, 50, 300, 300);");

  if (skippedFillCount > 0) {
    lines.push(`// Note: ${skippedFillCount} fill-bucket action(s) were skipped in export.`);
  }

  return lines.join("\n");
}

function buildImageExportData() {
  redrawAll();
  const width = state.canvasWidth;
  const height = state.canvasHeight;
  const imageData = ctx.getImageData(0, 0, width, height);
  const data = imageData.data;
  const palette = [];
  const paletteIndex = new Map();
  const runs = [];

  const getPacked = (offset) => (
    (((data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]) >>> 0)
  );

  const getColorIndex = (packed) => {
    if (paletteIndex.has(packed)) {
      return paletteIndex.get(packed);
    }
    const r = (packed >>> 24) & 255;
    const g = (packed >>> 16) & 255;
    const b = (packed >>> 8) & 255;
    const a = packed & 255;
    const index = palette.length;
    palette.push([r, g, b, a]);
    paletteIndex.set(packed, index);
    return index;
  };

  for (let y = 0; y < height; y += 1) {
    let x = 0;
    while (x < width) {
      const start = x;
      const startOffset = (y * width + x) * 4;
      const packed = getPacked(startOffset);
      const colorId = getColorIndex(packed);
      x += 1;
      while (x < width) {
        const nextOffset = (y * width + x) * 4;
        if (getPacked(nextOffset) !== packed) {
          break;
        }
        x += 1;
      }
      runs.push([start, y, x - start, colorId]);
    }
  }

  return { width, height, palette, runs };
}

function generatePjsImageCode() {
  const img = buildImageExportData();
  const lines = [];
  lines.push("// Generated by Graphics Designer (Image Mode)");
  lines.push(`var ART_IMG_W=${img.width};`);
  lines.push(`var ART_IMG_H=${img.height};`);
  lines.push(`var ART_IMG_P=${JSON.stringify(img.palette)};`);
  lines.push(`var ART_IMG_R=${JSON.stringify(img.runs)};`);
  lines.push("var __artImgCache={};");
  lines.push("var _artBuildImg=function(w,h){");
  lines.push("    var key=w+\"x\"+h;");
  lines.push("    if(__artImgCache[key]){return __artImgCache[key];}");
  lines.push("    var backup=get(0,0,w,h);");
  lines.push("    pushStyle();");
  lines.push("    noStroke();");
  lines.push("    var sx=w/ART_IMG_W,sy=h/ART_IMG_H,rowH=max(1,sy);");
  lines.push("    for(var i=0;i<ART_IMG_R.length;i++){");
  lines.push("        var r=ART_IMG_R[i],c=ART_IMG_P[r[3]];");
  lines.push("        fill(c[0],c[1],c[2],c[3]);");
  lines.push("        rect(r[0]*sx,r[1]*sy,max(1,r[2]*sx),rowH);");
  lines.push("    }");
  lines.push("    popStyle();");
  lines.push("    __artImgCache[key]=get(0,0,w,h);");
  lines.push("    image(backup,0,0);");
  lines.push("    return __artImgCache[key];");
  lines.push("};");
  lines.push("var drawArt=function(x,y,w,h){");
  lines.push("    var aw=max(1,floor(abs(w))),ah=max(1,floor(abs(h)));");
  lines.push("    var img=_artBuildImg(aw,ah);");
  lines.push("    pushMatrix();");
  lines.push("    translate(x,y);");
  lines.push("    if(w<0||h<0){scale(w<0?-1:1,h<0?-1:1);}");
  lines.push("    image(img,w<0?-aw:0,h<0?-ah:0,aw,ah);");
  lines.push("    popMatrix();");
  lines.push("};");
  lines.push("");
  lines.push("// Example: drawArt(50, 50, 300, 300);");
  lines.push("// Image mode includes bucket-fill output.");
  return lines.join("\n");
}

function encodeAction(action) {
  if (action.type === "stroke") {
    const brush = action.brush;
    const baseColor = brush === "eraser" ? "#ffffff" : action.color;
    const [r, g, b] = hexToRgb(baseColor);
    const alpha = toAlpha(action.opacity);

    if (brush === "spray") {
      return [
        "P",
        r,
        g,
        b,
        alpha,
        normSize(Math.max(1, action.size * 0.5)),
        flattenPointList(action.dots)
      ];
    }

    if (brush === "glow") {
      return [
        "G",
        r,
        g,
        b,
        alpha,
        normSize(action.size),
        flattenPointList(action.points)
      ];
    }

    return [
      "S",
      r,
      g,
      b,
      alpha,
      normSize(brush === "pencil" ? Math.max(1, action.size * 0.55) : action.size),
      flattenPointList(action.points)
    ];
  }

  if (action.type === "line") {
    const [r, g, b] = hexToRgb(action.strokeColor || action.color || "#0d2238");
    return [
      "L",
      r,
      g,
      b,
      toAlpha(action.opacity),
      normSize(getShapeStrokeWeight(action)),
      normX(action.x1),
      normY(action.y1),
      normX(action.x2),
      normY(action.y2)
    ];
  }

  if (action.type === "arc") {
    const [fr, fg, fb] = hexToRgb(action.fillColor || action.color || "#00b7ff");
    const [r, g, b] = hexToRgb(action.strokeColor || action.color || "#0d2238");
    return [
      "A",
      r,
      g,
      b,
      toAlpha(action.opacity),
      normSize(getShapeStrokeWeight(action)),
      normX(action.x + action.w / 2),
      normY(action.y + action.h / 2),
      normX(action.w),
      normY(action.h),
      shortNumber(typeof action.start === "number" ? action.start : 0, 4),
      shortNumber(typeof action.stop === "number" ? action.stop : Math.PI, 4),
      shortNumber(getShapeRotation(action), 4),
      action.filled ? 1 : 0,
      fr,
      fg,
      fb
    ];
  }

  if (action.type === "rect") {
    const [fr, fg, fb] = hexToRgb(action.fillColor || action.color || "#00b7ff");
    const [sr, sg, sb] = hexToRgb(action.strokeColor || action.color || "#0d2238");
    const alpha = toAlpha(action.opacity);
    const strokeWeight = normSize(getShapeStrokeWeight(action));

    if (action.gradient && action.gradient.endColor) {
      const [gr, gg, gb] = hexToRgb(action.gradient.endColor);
      return [
        "RG",
        fr,
        fg,
        fb,
        gr,
        gg,
        gb,
        gradientModeCode(action.gradient.direction),
        sr,
        sg,
        sb,
        alpha,
        strokeWeight,
        action.filled ? 1 : 0,
        normX(action.x),
        normY(action.y),
        normX(action.w),
        normY(action.h),
        shortNumber(getShapeRotation(action), 4)
      ];
    }

    return [
      "R",
      fr,
      fg,
      fb,
      sr,
      sg,
      sb,
      alpha,
      strokeWeight,
      action.filled ? 1 : 0,
      normX(action.x),
      normY(action.y),
      normX(action.w),
      normY(action.h),
      shortNumber(getShapeRotation(action), 4)
    ];
  }

  if (action.type === "ellipse") {
    const [fr, fg, fb] = hexToRgb(action.fillColor || action.color || "#00b7ff");
    const [sr, sg, sb] = hexToRgb(action.strokeColor || action.color || "#0d2238");
    const alpha = toAlpha(action.opacity);
    const strokeWeight = normSize(getShapeStrokeWeight(action));

    if (action.gradient && action.gradient.endColor) {
      const [gr, gg, gb] = hexToRgb(action.gradient.endColor);
      return [
        "EG",
        fr,
        fg,
        fb,
        gr,
        gg,
        gb,
        gradientModeCode(action.gradient.direction),
        sr,
        sg,
        sb,
        alpha,
        strokeWeight,
        action.filled ? 1 : 0,
        normX(action.x + action.w / 2),
        normY(action.y + action.h / 2),
        normX(action.w),
        normY(action.h),
        shortNumber(getShapeRotation(action), 4)
      ];
    }

    return [
      "E",
      fr,
      fg,
      fb,
      sr,
      sg,
      sb,
      alpha,
      strokeWeight,
      action.filled ? 1 : 0,
      normX(action.x + action.w / 2),
      normY(action.y + action.h / 2),
      normX(action.w),
      normY(action.h),
      shortNumber(getShapeRotation(action), 4)
    ];
  }

  if (action.type === "polygonShape") {
    const [fr, fg, fb] = hexToRgb(action.fillColor || action.color || "#00b7ff");
    const [sr, sg, sb] = hexToRgb(action.strokeColor || action.color || "#0d2238");
    const points = getPolygonShapePoints(action);

    if (action.shape === "triangle" && action.gradient && action.gradient.endColor && points.length >= 3) {
      const [gr, gg, gb] = hexToRgb(action.gradient.endColor);
      return [
        "TG",
        fr,
        fg,
        fb,
        gr,
        gg,
        gb,
        gradientModeCode(action.gradient.direction),
        sr,
        sg,
        sb,
        toAlpha(action.opacity),
        normSize(getShapeStrokeWeight(action)),
        action.filled ? 1 : 0,
        normX(points[0].x),
        normY(points[0].y),
        normX(points[1].x),
        normY(points[1].y),
        normX(points[2].x),
        normY(points[2].y)
      ];
    }

    if (action.shape === "quad" && action.gradient && action.gradient.endColor && points.length >= 4) {
      const [gr, gg, gb] = hexToRgb(action.gradient.endColor);
      return [
        "QG",
        fr,
        fg,
        fb,
        gr,
        gg,
        gb,
        gradientModeCode(action.gradient.direction),
        sr,
        sg,
        sb,
        toAlpha(action.opacity),
        normSize(getShapeStrokeWeight(action)),
        action.filled ? 1 : 0,
        normX(points[0].x),
        normY(points[0].y),
        normX(points[1].x),
        normY(points[1].y),
        normX(points[2].x),
        normY(points[2].y),
        normX(points[3].x),
        normY(points[3].y)
      ];
    }

    return [
      "VP",
      fr,
      fg,
      fb,
      sr,
      sg,
      sb,
      toAlpha(action.opacity),
      normSize(getShapeStrokeWeight(action)),
      0,
      1,
      action.filled ? 1 : 0,
      flattenPointList(points)
    ];
  }

  if (action.type === "vertexPath") {
    const [fr, fg, fb] = hexToRgb(action.fillColor || action.color || "#00b7ff");
    const [sr, sg, sb] = hexToRgb(action.strokeColor || action.color || "#0d2238");
    return [
      "VP",
      fr,
      fg,
      fb,
      sr,
      sg,
      sb,
      toAlpha(action.opacity),
      normSize(getShapeStrokeWeight(action)),
      action.curved ? 1 : 0,
      action.closed ? 1 : 0,
      action.closed && action.filled ? 1 : 0,
      flattenPointList(action.points)
    ];
  }

  if (action.type === "bezier") {
    const [sr, sg, sb] = hexToRgb(action.strokeColor || action.color || "#0d2238");
    return [
      "B",
      sr,
      sg,
      sb,
      toAlpha(action.opacity),
      normSize(getShapeStrokeWeight(action)),
      normX(action.x1),
      normY(action.y1),
      normX(action.cx1),
      normY(action.cy1),
      normX(action.cx2),
      normY(action.cy2),
      normX(action.x2),
      normY(action.y2)
    ];
  }

  return null;
}

function flattenPointList(points) {
  if (!Array.isArray(points)) {
    return [];
  }

  const flat = [];
  for (const point of points) {
    flat.push(normX(point.x), normY(point.y));
  }
  return flat;
}

async function copyPjsOutput() {
  const content = pjsOutput.value.trim();
  if (!content) {
    updateSaveStatus("No PJS code to copy");
    return;
  }

  try {
    await navigator.clipboard.writeText(content);
    updateSaveStatus("PJS copied to clipboard");
  } catch (error) {
    pjsOutput.focus();
    pjsOutput.select();
    const copied = document.execCommand("copy");
    updateSaveStatus(copied ? "PJS copied to clipboard" : "Clipboard blocked by browser", !copied);
  }
}

function downloadPng() {
  const href = drawCanvas.toDataURL("image/png");
  triggerDownload(href, "khan-graphics-design.png");
  updateSaveStatus("PNG downloaded");
}

function downloadPjs() {
  const code = pjsOutput.value;
  if (!code.trim()) {
    updateSaveStatus("No PJS code to download");
    return;
  }

  const blob = new Blob([code], { type: "text/plain;charset=utf-8" });
  const href = URL.createObjectURL(blob);
  triggerDownload(href, "khan-graphics-design.js");
  URL.revokeObjectURL(href);
  updateSaveStatus("PJS file downloaded");
}

function triggerDownload(href, fileName) {
  const link = document.createElement("a");
  link.href = href;
  link.download = fileName;
  link.click();
}

function ensureActionIds(actions) {
  for (const action of actions) {
    if (!action.id) {
      action.id = createActionId();
    }
    if (action.type === "polygonShape" && (!Array.isArray(action.points) || action.points.length < 3)) {
      action.points = getPolygonShapePoints(action);
    }
  }
}

function createActionId() {
  const id = `a_${Date.now().toString(36)}_${actionCounter}`;
  actionCounter += 1;
  return id;
}

function cloneAction(action) {
  return JSON.parse(JSON.stringify(action));
}

function cloneActions(actions) {
  return JSON.parse(JSON.stringify(actions));
}

function clonePointList(points) {
  if (!Array.isArray(points)) {
    return [];
  }
  return points.map((point) => ({ x: point.x, y: point.y }));
}

function getBoundsFromPoints(points, margin) {
  if (!Array.isArray(points) || points.length === 0) {
    return null;
  }

  let xMin = points[0].x;
  let yMin = points[0].y;
  let xMax = points[0].x;
  let yMax = points[0].y;
  for (let i = 1; i < points.length; i += 1) {
    xMin = Math.min(xMin, points[i].x);
    yMin = Math.min(yMin, points[i].y);
    xMax = Math.max(xMax, points[i].x);
    yMax = Math.max(yMax, points[i].y);
  }
  return { xMin: xMin - margin, yMin: yMin - margin, xMax: xMax + margin, yMax: yMax + margin };
}

function hitTestPolyline(points, point, weight, closed) {
  const tol = Math.max(5, weight * 0.75);
  for (let i = 1; i < points.length; i += 1) {
    if (distanceToSegment(point.x, point.y, points[i - 1].x, points[i - 1].y, points[i].x, points[i].y) <= tol) {
      return true;
    }
  }
  if (closed && points.length > 2) {
    const last = points[points.length - 1];
    const first = points[0];
    if (distanceToSegment(point.x, point.y, last.x, last.y, first.x, first.y) <= tol) {
      return true;
    }
  }
  return false;
}

function isPointInPolygon(points, point) {
  let inside = false;
  for (let i = 0, j = points.length - 1; i < points.length; j = i++) {
    const xi = points[i].x;
    const yi = points[i].y;
    const xj = points[j].x;
    const yj = points[j].y;
    const intersect = ((yi > point.y) !== (yj > point.y)) &&
      (point.x < ((xj - xi) * (point.y - yi)) / (yj - yi + 0.000001) + xi);
    if (intersect) {
      inside = !inside;
    }
  }
  return inside;
}

function hitTestPointInPolyOrEdge(points, point, filled, weight) {
  if (filled && isPointInPolygon(points, point)) {
    return true;
  }
  return hitTestPolyline(points, point, weight, true);
}

function sampleBezier(action, t) {
  const u = 1 - t;
  const tt = t * t;
  const uu = u * u;
  const uuu = uu * u;
  const ttt = tt * t;
  return {
    x: uuu * action.x1 + 3 * uu * t * action.cx1 + 3 * u * tt * action.cx2 + ttt * action.x2,
    y: uuu * action.y1 + 3 * uu * t * action.cy1 + 3 * u * tt * action.cy2 + ttt * action.y2
  };
}

function rgbaString(hex, opacity) {
  const [r, g, b] = hexToRgb(hex);
  return `rgba(${r}, ${g}, ${b}, ${clamp(opacity, 0, 1)})`;
}

function hexToRgb(hex) {
  const value = String(hex || "#000000").replace("#", "");
  const normalized = value.length === 3
    ? value.split("").map((char) => char + char).join("")
    : value;

  const number = Number.parseInt(normalized, 16);
  return [
    (number >> 16) & 255,
    (number >> 8) & 255,
    number & 255
  ];
}

function normalizeGradientDirection(direction) {
  const value = String(direction || "").toLowerCase();
  if (value === "horizontal") {
    return "left-right";
  }
  if (value === "vertical") {
    return "top-bottom";
  }
  if (
    value === "left-right" ||
    value === "right-left" ||
    value === "top-bottom" ||
    value === "bottom-top" ||
    value === "center-out" ||
    value === "center-in"
  ) {
    return value;
  }
  return "left-right";
}

function gradientModeCode(direction) {
  const mode = normalizeGradientDirection(direction);
  if (mode === "right-left") {
    return 1;
  }
  if (mode === "top-bottom") {
    return 2;
  }
  if (mode === "bottom-top") {
    return 3;
  }
  if (mode === "center-out") {
    return 4;
  }
  if (mode === "center-in") {
    return 5;
  }
  return 0;
}

function toAlpha(opacity) {
  return Math.round(clamp(opacity, 0, 1) * 255);
}

function normX(value) {
  return shortNumber(value / state.canvasWidth, 4);
}

function normY(value) {
  return shortNumber(value / state.canvasHeight, 4);
}

function normSize(value) {
  const base = (state.canvasWidth + state.canvasHeight) / 2;
  return shortNumber(value / base, 4);
}

function shortNumber(value, decimals = 1) {
  return Number(Number(value).toFixed(decimals));
}

function distance(a, b) {
  const dx = a.x - b.x;
  const dy = a.y - b.y;
  return Math.sqrt(dx * dx + dy * dy);
}

function angleBetween(center, point) {
  return Math.atan2(point.y - center.y, point.x - center.x);
}

function normalizeAngle(angle) {
  const twoPi = Math.PI * 2;
  let out = angle % twoPi;
  if (out < 0) {
    out += twoPi;
  }
  return out;
}

function getShapeRotation(action) {
  return action && typeof action.rotation === "number" ? action.rotation : 0;
}

function toLocalRotatedPoint(point, center, rotation) {
  if (!rotation) {
    return { x: point.x, y: point.y };
  }
  return rotatePoint(point, center, -rotation);
}

function getArcCenter(action) {
  return {
    x: action.x + action.w / 2,
    y: action.y + action.h / 2
  };
}

function arcPointAtAngle(action, angle, includeRotation = true) {
  const center = getArcCenter(action);
  const rx = Math.max(1, Math.abs(action.w) / 2);
  const ry = Math.max(1, Math.abs(action.h) / 2);
  const point = {
    x: center.x + Math.cos(angle) * rx,
    y: center.y + Math.sin(angle) * ry
  };
  if (!includeRotation) {
    return point;
  }
  const rotation = getShapeRotation(action);
  return rotation ? rotatePoint(point, center, rotation) : point;
}

function rotatePoint(point, center, angle) {
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  const dx = point.x - center.x;
  const dy = point.y - center.y;
  return {
    x: center.x + dx * cos - dy * sin,
    y: center.y + dx * sin + dy * cos
  };
}

function distanceToSegment(px, py, x1, y1, x2, y2) {
  const dx = x2 - x1;
  const dy = y2 - y1;

  if (dx === 0 && dy === 0) {
    const ddx = px - x1;
    const ddy = py - y1;
    return Math.sqrt(ddx * ddx + ddy * ddy);
  }

  const t = clamp(((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy), 0, 1);
  const projX = x1 + t * dx;
  const projY = y1 + t * dy;
  const distX = px - projX;
  const distY = py - projY;
  return Math.sqrt(distX * distX + distY * distY);
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function capitalize(text) {
  if (!text) {
    return "";
  }
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function formatTime(isoString) {
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) {
    return "now";
  }

  return date.toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });
}
