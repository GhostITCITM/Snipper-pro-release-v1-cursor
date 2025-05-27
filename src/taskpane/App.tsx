import React, { useState, useEffect, useCallback } from "react";
import { PdfViewer, Rectangle } from "../viewer/PdfViewer";
import { OCREngine, SumCalculator } from "../ocr/ocr";
import { parseTableFromOCRText } from "../ocr/table";
import {
  write,
  writeTick,
  writeCross,
  writeTableToCell,
  logSnip,
  getCurrentCellAddress,
  findSnipByCell
} from "../excel/excel";
import { modeManager, SnipMode } from "../helpers/mode";
import {
  FluentProvider,
  webLightTheme,
  Button,
  Text,
  Spinner
} from "@fluentui/react-components";

interface AppState {
  currentDocument: ArrayBuffer | null;
  documentName: string;
  currentMode: SnipMode | null;
  isSelectionMode: boolean;
  selectedCellAddress: string | null;
  isProcessing: boolean;
  statusMessage: string;
  highlightRect: Rectangle | null;
  highlightPage: number | null;
}

const App: React.FC = () => {
  const [state, setState] = useState<AppState>({
    currentDocument: null,
    documentName: "",
    currentMode: null,
    isSelectionMode: false,
    selectedCellAddress: null,
    isProcessing: false,
    statusMessage: "Ready",
    highlightRect: null,
    highlightPage: null
  });

  // Initialize OCR engine on mount
  useEffect(() => {
    OCREngine.initialize().catch(console.error);

    // Selection change listener to reveal stored snips
    const selectionHandler = async () => {
      try {
        // If we are in an active snip mode, ignore (handled elsewhere)
        if (await modeManager.isActive()) return;

        const cell = await getCurrentCellAddress();
        const record = await findSnipByCell(cell);

        if (record) {
          setState((prev) => ({
            ...prev,
            highlightRect: JSON.parse(record.rect),
            highlightPage: record.page
          }));
        } else {
          setState((prev) => ({ ...prev, highlightRect: null, highlightPage: null }));
        }
      } catch (err) {
        console.error("Error handling highlight:", err);
      }
    };

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      selectionHandler
    );

    return () => {
      // remove handler when taskpane unloads
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: selectionHandler }
      );
      OCREngine.terminate().catch(console.error);
    };
  }, []);

  // Monitor mode changes from ribbon commands
  useEffect(() => {
    const checkModeChanges = async () => {
      try {
        const currentMode = await modeManager.current();
        const isActive = await modeManager.isActive();

        if (currentMode && isActive) {
          setState((prev) => ({
            ...prev,
            currentMode,
            isSelectionMode:
              currentMode === "text" || currentMode === "sum" || currentMode === "table",
            statusMessage: `${currentMode.toUpperCase()} mode active - Select a cell then draw rectangle`
          }));

          // Get current cell selection
          try {
            const cellAddress = await getCurrentCellAddress();
            setState((prev) => ({
              ...prev,
              selectedCellAddress: cellAddress
            }));
          } catch (error) {
            console.warn("Could not get cell address:", error);
          }
        } else {
          setState((prev) => ({
            ...prev,
            currentMode: null,
            isSelectionMode: false,
            selectedCellAddress: null,
            statusMessage: "Ready"
          }));
        }
      } catch (error) {
        console.error("Error checking mode:", error);
      }
    };

    // Check immediately and then periodically
    checkModeChanges();
    const interval = setInterval(checkModeChanges, 1000);

    return () => clearInterval(interval);
  }, []);

  const handleFileImport = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (file.type !== "application/pdf" && !file.type.startsWith("image/")) {
      setState((prev) => ({
        ...prev,
        statusMessage: "Please select a PDF file or image"
      }));
      return;
    }

    file
      .arrayBuffer()
      .then((buffer) => {
        setState((prev) => ({
          ...prev,
          currentDocument: buffer,
          documentName: file.name,
          statusMessage: `Loaded: ${file.name}`
        }));
      })
      .catch((error) => {
        console.error("Error reading file:", error);
        setState((prev) => ({
          ...prev,
          statusMessage: "Error loading file"
        }));
      });

    // Clear the input
    event.target.value = "";
  }, []);

  const handleRectangleSelection = useCallback(
    async (imageData: ImageData, pageNumber: number, rectangle: Rectangle) => {
      if (!state.currentMode || !state.selectedCellAddress) {
        setState((prev) => ({
          ...prev,
          statusMessage: "Please select a cell first"
        }));
        return;
      }

      setState((prev) => ({
        ...prev,
        isProcessing: true,
        statusMessage: "Processing selection..."
      }));

      try {
        let resultValue = "";
        let snippedText = "";

        const activeMode = state.currentMode!; // non-null here

        switch (activeMode) {
          case "text":
            const textResult = await OCREngine.recognizeText(imageData);
            resultValue = textResult.text;
            snippedText = textResult.text;
            await write(resultValue);
            break;

          case "sum":
            const sumResult = await SumCalculator.extractAndSum(imageData);
            resultValue = sumResult.sum.toString();
            snippedText = sumResult.text;
            await write(resultValue);
            break;

          case "table":
            const ocrResult = await OCREngine.recognizeText(imageData);
            const tableData = parseTableFromOCRText(ocrResult.text);
            if (tableData.length > 0) {
              await writeTableToCell(tableData);
              resultValue = `Table (${tableData.length}x${tableData[0]?.length || 0})`;
              snippedText = tableData.map((row) => row.join("\t")).join("\n");
            } else {
              resultValue = "No table detected";
              snippedText = ocrResult.text;
              await write(resultValue);
            }
            break;

          case "validation":
            await writeTick();
            resultValue = "✓";
            snippedText = "Validation mark";
            break;

          case "exception":
            await writeCross();
            resultValue = "✗";
            snippedText = "Exception mark";
            break;

          default:
            throw new Error(`Unknown mode: ${activeMode}`);
        }

        // Save snip record
        await logSnip({
          cell: state.selectedCellAddress,
          page: pageNumber,
          rect: JSON.stringify(rectangle),
          mode: activeMode,
          text: snippedText
        });

        // Clear mode and reset state
        await modeManager.clearMode();

        setState((prev) => ({
          ...prev,
          currentMode: null,
          isSelectionMode: false,
          selectedCellAddress: null,
          isProcessing: false,
          statusMessage: `${activeMode.toUpperCase()} snip completed successfully`
        }));
      } catch (error) {
        console.error("Error processing snip:", error);
        const err = error as any;
        setState((prev) => ({
          ...prev,
          isProcessing: false,
          statusMessage: `Error: ${err?.message || "Processing failed"}`
        }));
      }
    },
    [state.currentMode, state.selectedCellAddress, state.documentName]
  );

  const clearCurrentMode = useCallback(async () => {
    await modeManager.clearMode();
    setState((prev) => ({
      ...prev,
      currentMode: null,
      isSelectionMode: false,
      selectedCellAddress: null,
      statusMessage: "Ready"
    }));
  }, []);

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ height: "100vh", display: "flex", flexDirection: "column" }}>
        {/* Header */}
        <div style={{ padding: "16px", borderBottom: "1px solid #ddd", background: "#fff" }}>
          <Text size={500} weight="semibold">
            SnipperClone - Document Analysis
          </Text>

          {/* Import Section */}
          <div style={{ marginTop: "12px" }}>
            <input
              type="file"
              accept=".pdf,image/*"
              onChange={handleFileImport}
              style={{ display: "none" }}
              id="file-input"
            />
            <Button
              onClick={() => document.getElementById("file-input")?.click()}
              appearance="primary"
              size="small"
            >
              Import Document
            </Button>

            {state.documentName && (
              <Text size={200} style={{ marginLeft: "8px", color: "#666" }}>
                {state.documentName}
              </Text>
            )}
          </div>
        </div>

        {/* Mode indicator */}
        {state.currentMode && (
          <div
            style={{
              padding: "8px 16px",
              background: "#e6f3ff",
              borderBottom: "1px solid #ccc",
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center"
            }}
          >
            <Text size={300}>
              <strong>{state.currentMode.toUpperCase()} MODE:</strong>
              {state.selectedCellAddress
                ? ` Target cell: ${state.selectedCellAddress}`
                : " Select a cell first"}
            </Text>
            <Button size="small" onClick={clearCurrentMode}>
              Cancel
            </Button>
          </div>
        )}

        {/* PDF Viewer */}
        <div style={{ flex: 1, overflow: "hidden" }}>
          <PdfViewer
            buffer={state.currentDocument}
            onSelect={handleRectangleSelection}
            isSelectionMode={state.isSelectionMode && !!state.selectedCellAddress}
            currentMode={state.currentMode}
            highlightRect={state.highlightRect}
            highlightPage={state.highlightPage}
          />
        </div>

        {/* Status Bar */}
        <div
          style={{
            padding: "8px 16px",
            background: "#f8f9fa",
            borderTop: "1px solid #ddd",
            fontSize: "12px",
            color: "#666",
            display: "flex",
            alignItems: "center",
            gap: "8px"
          }}
        >
          {state.isProcessing && <Spinner size="tiny" label="" />}
          <span>Status: {state.statusMessage}</span>
        </div>
      </div>
    </FluentProvider>
  );
};

export default App;
