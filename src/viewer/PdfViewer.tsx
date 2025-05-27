import React, { useEffect, useRef, useState } from "react";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist/legacy/build/pdf";
import { Button, Text } from "@fluentui/react-components";
import { ArrowLeft16Regular, ArrowRight16Regular } from "@fluentui/react-icons";
import type { PDFDocumentProxy } from "pdfjs-dist/types/src/display/api";

// Use local worker file with dynamic path resolution
GlobalWorkerOptions.workerSrc =
  Office.context.document.url.replace(/[^/]+$/, "") + "pdf.worker.min.js";

export interface Rectangle {
  x: number;
  y: number;
  width: number;
  height: number;
}

type Props = {
  buffer: ArrayBuffer | null;
  onSelect: (img: ImageData, page: number, rect: Rectangle) => void;
  isSelectionMode: boolean;
  currentMode: string | null;
  highlightRect?: Rectangle | null;
  highlightPage?: number | null;
};

export const PdfViewer: React.FC<Props> = ({
  buffer,
  onSelect,
  isSelectionMode,
  currentMode,
  highlightRect = null,
  highlightPage = null
}) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [pdf, setPdf] = useState<PDFDocumentProxy | null>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [totalPages, setTotalPages] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Selection state
  const [isSelecting, setIsSelecting] = useState(false);
  const [selectionStart, setSelectionStart] = useState<{ x: number; y: number } | null>(null);
  const [selectionRect, setSelectionRect] = useState<Rectangle | null>(null);

  // Load PDF
  useEffect(() => {
    if (!buffer) {
      setPdf(null);
      setTotalPages(0);
      setCurrentPage(1);
      return;
    }

    setIsLoading(true);
    setError(null);

    getDocument({ data: buffer })
      .promise.then((pdfDoc) => {
        setPdf(pdfDoc);
        setTotalPages(pdfDoc.numPages);
        setCurrentPage(1);
        setIsLoading(false);
      })
      .catch((err) => {
        console.error("Error loading PDF:", err);
        setError("Failed to load PDF document");
        setIsLoading(false);
      });
  }, [buffer]);

  // If highlightPage provided, go to that page
  useEffect(() => {
    if (highlightPage && highlightPage !== currentPage) {
      setCurrentPage(highlightPage);
    }
  }, [highlightPage]);

  // Render current page whenever currentPage or pdf changes
  useEffect(() => {
    if (!pdf || !canvasRef.current) return;

    (async () => {
      setIsLoading(true);
      try {
        const page = await pdf.getPage(currentPage);
        const viewport = page.getViewport({ scale: 1.25 });
        const canvas = canvasRef.current!;
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext("2d")!;
        await page.render({ canvasContext: ctx, viewport }).promise;
        setIsLoading(false);
      } catch (err) {
        console.error("Error rendering page:", err);
        setError("Failed to render page");
        setIsLoading(false);
      }
    })();
  }, [pdf, currentPage]);

  // Mouse event handlers for rectangle selection
  const handleMouseDown = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isSelectionMode || !canvasRef.current) return;

    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    setIsSelecting(true);
    setSelectionStart({ x, y });
    setSelectionRect(null);
  };

  const handleMouseMove = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isSelecting || !selectionStart || !canvasRef.current) return;

    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const currentX = e.clientX - rect.left;
    const currentY = e.clientY - rect.top;

    const selection: Rectangle = {
      x: Math.min(selectionStart.x, currentX),
      y: Math.min(selectionStart.y, currentY),
      width: Math.abs(currentX - selectionStart.x),
      height: Math.abs(currentY - selectionStart.y)
    };

    setSelectionRect(selection);
  };

  const handleMouseUp = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isSelecting || !selectionStart || !canvasRef.current) return;

    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const currentX = e.clientX - rect.left;
    const currentY = e.clientY - rect.top;

    const selection: Rectangle = {
      x: Math.min(selectionStart.x, currentX),
      y: Math.min(selectionStart.y, currentY),
      width: Math.abs(currentX - selectionStart.x),
      height: Math.abs(currentY - selectionStart.y)
    };

    // Only process if selection has meaningful size
    if (selection.width > 5 && selection.height > 5) {
      const context = canvas.getContext("2d")!;
      const imageData = context.getImageData(
        selection.x,
        selection.y,
        selection.width,
        selection.height
      );

      onSelect(imageData, currentPage, selection);
    }

    // Reset selection state
    setIsSelecting(false);
    setSelectionStart(null);
    setSelectionRect(null);
  };

  // Navigation handlers
  const goToPreviousPage = () => {
    if (currentPage > 1) {
      setCurrentPage(currentPage - 1);
    }
  };

  const goToNextPage = () => {
    if (currentPage < totalPages) {
      setCurrentPage(currentPage + 1);
    }
  };

  if (!buffer) {
    return (
      <div style={{ padding: "20px", textAlign: "center", color: "#666" }}>
        <p>No document loaded</p>
        <small>Click "Import Docs" to load a PDF file</small>
      </div>
    );
  }

  if (error) {
    return <div style={{ padding: "20px", color: "#d13438" }}>Error: {error}</div>;
  }

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      {/* Toolbar */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "8px 12px",
          background: "#fff",
          borderBottom: "1px solid #ddd"
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          <Button
            size="small"
            appearance="subtle"
            icon={<ArrowLeft16Regular />}
            disabled={currentPage <= 1}
            onClick={goToPreviousPage}
          >
            Prev
          </Button>

          <Text size={200} style={{ margin: "0 8px" }}>
            Page {currentPage} of {totalPages}
          </Text>

          <Button
            size="small"
            appearance="subtle"
            icon={<ArrowRight16Regular />}
            iconPosition="after"
            disabled={currentPage >= totalPages}
            onClick={goToNextPage}
          >
            Next
          </Button>
        </div>

        {currentMode && (
          <div
            style={{
              padding: "4px 8px",
              background: "#0078d4",
              color: "white",
              borderRadius: "3px",
              fontSize: "11px",
              fontWeight: 500,
              textTransform: "uppercase"
            }}
          >
            {currentMode} Mode
          </div>
        )}
      </div>

      {/* Content */}
      <div style={{ flex: 1, overflow: "auto", background: "#f0f0f0" }}>
        {isLoading ? (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              padding: "40px",
              color: "#666",
              fontSize: "14px"
            }}
          >
            Loading...
          </div>
        ) : (
          <div
            style={{
              position: "relative",
              display: "flex",
              justifyContent: "center",
              minHeight: "100%",
              alignItems: "flex-start",
              padding: "20px"
            }}
          >
            <div style={{ position: "relative" }}>
              <canvas
                ref={canvasRef}
                style={{
                  border: "1px solid #ccc",
                  boxShadow: "0 2px 10px rgba(0, 0, 0, 0.1)",
                  background: "white",
                  cursor: isSelectionMode ? "crosshair" : "default"
                }}
                onMouseDown={handleMouseDown}
                onMouseMove={handleMouseMove}
                onMouseUp={handleMouseUp}
              />

              {/* Selection overlay while drawing */}
              {selectionRect && isSelecting && (
                <div
                  style={{
                    position: "absolute",
                    left: selectionRect.x,
                    top: selectionRect.y,
                    width: selectionRect.width,
                    height: selectionRect.height,
                    border: "2px dashed #0078d4",
                    background: "rgba(0, 120, 212, 0.1)",
                    pointerEvents: "none",
                    zIndex: 10
                  }}
                />
              )}

              {/* Highlight overlay for stored snip */}
              {highlightRect && !isSelecting && (
                <div
                  style={{
                    position: "absolute",
                    left: highlightRect.x,
                    top: highlightRect.y,
                    width: highlightRect.width,
                    height: highlightRect.height,
                    border: "2px solid #e81123",
                    background: "rgba(232, 17, 35, 0.15)",
                    pointerEvents: "none",
                    zIndex: 9
                  }}
                />
              )}
            </div>
          </div>
        )}
      </div>

      {/* Status bar */}
      <div
        style={{
          padding: "4px 12px",
          background: "#f8f9fa",
          borderTop: "1px solid #ddd",
          fontSize: "11px",
          color: "#666"
        }}
      >
        {isSelectionMode ? `Ready to select area for ${currentMode} snip` : "Ready"}
        {pdf && ` • Document loaded: ${totalPages} pages`}
      </div>
    </div>
  );
};
