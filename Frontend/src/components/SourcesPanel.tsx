import {
  ExternalLink,
  X,
  Menu,
  FileSpreadsheet,
  MessageSquarePlus,
  History,
  Settings,
  Upload,
} from "lucide-react";
import { useState, useRef, useEffect } from "react";
import { useSpreadsheet } from "../contexts/SpreadsheetContext";
import { WOPIViewer } from "./WOPIViewer";

// Update this with your server's IP address
const SERVER_IP = "10.172.168.165";
const BASE_URL = `http://${SERVER_IP}:5000`;

export const SourcesPanel = () => {
  const [isMenuExpanded, setIsMenuExpanded] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<string>("");
  const [isDragging, setIsDragging] = useState(false);
  const {
    highlightedCells,
    focusedCells,
    currentFile,
    setCurrentFile,
    setFileId,
  } = useSpreadsheet();
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Add event listener for highlighted file updates
  useEffect(() => {
    const handleHighlightedFile = (
      event: CustomEvent<{ fileName: string; fileId: string }>
    ) => {
      console.log("Loading highlighted file:", event.detail.fileName);
      setCurrentFile(event.detail.fileName);
      setFileId(event.detail.fileId);
    };

    window.addEventListener(
      "loadHighlightedFile",
      handleHighlightedFile as EventListener
    );
    return () => {
      window.removeEventListener(
        "loadHighlightedFile",
        handleHighlightedFile as EventListener
      );
    };
  }, [setCurrentFile, setFileId]);

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setUploadStatus("Uploading file...");
      const formData = new FormData();
      formData.append("file", file);

      // First upload to local server
      const localResponse = await fetch("http://localhost:3001/upload", {
        method: "POST",
        body: formData,
      });

      if (!localResponse.ok) {
        throw new Error(`Local upload failed: ${localResponse.statusText}`);
      }

      const localData = await localResponse.json();
      setCurrentFile(localData.fileName);

      // Then upload to backend server
      const backendResponse = await fetch(`${BASE_URL}/upload`, {
        method: "POST",
        body: formData,
      });

      if (!backendResponse.ok) {
        throw new Error(`Backend upload failed: ${backendResponse.statusText}`);
      }

      const backendData = await backendResponse.json();
      if (!backendData.file_id) {
        throw new Error("No file ID received from backend");
      }

      setFileId(backendData.file_id);
      setUploadStatus("File uploaded successfully!");

      // Clear the input
      e.target.value = "";
    } catch (error) {
      console.error("Upload error:", error);
      setUploadStatus(`Upload failed: ${error.message}`);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);

    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".xls"))) {
      const fakeEvent = {
        target: {
          files: [file],
          value: "",
        },
      } as unknown as React.ChangeEvent<HTMLInputElement>;
      await handleExcelUpload(fakeEvent);
    }
  };

  const handleOpenInNewTab = () => {
    if (currentFile) {
      const wopiUrl = `http://localhost:3000?file=${currentFile}`;
      window.open(wopiUrl, "_blank");
    }
  };

  // Handle cell highlighting through WOPI
  useEffect(() => {
    if (highlightedCells && highlightedCells.length > 0) {
      // The sheet name would need to be determined from your context or state
      const sheetName = "Sheet1"; // This should be dynamic based on your needs
      const cellRanges = highlightedCells.join(",");

      // The WOPIViewer component will handle the highlighting
      // This is just to keep track of the highlighted state
      console.log(`Highlighting cells: ${cellRanges} in sheet: ${sheetName}`);
    }
  }, [highlightedCells]);

  return (
    <div className="h-full flex">
      {/* Vertical Panel */}
      <div
        className={`bg-[#1f1f1f] flex-shrink-0 flex flex-col transition-all duration-200 ${
          isMenuExpanded ? "w-48" : "w-14"
        }`}
        onMouseEnter={() => setIsMenuExpanded(true)}
        onMouseLeave={() => setIsMenuExpanded(false)}
      >
        {/* Icons section */}
        <div className="flex-1 py-4">
          <div className="space-y-4">
            <label className="w-full flex items-center px-3 py-2 text-gray-400 hover:text-white hover:bg-[#2d2d2d] transition-colors whitespace-nowrap overflow-hidden cursor-pointer">
              <FileSpreadsheet className="w-5 h-5 flex-shrink-0" />
              {isMenuExpanded && (
                <span className="ml-3 text-sm">Upload Excel</span>
              )}
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleExcelUpload}
                className="hidden"
              />
            </label>
            {uploadStatus && isMenuExpanded && (
              <div className="px-3 text-xs text-gray-400">{uploadStatus}</div>
            )}
            <button className="w-full flex items-center px-3 py-2 text-gray-400 hover:text-white hover:bg-[#2d2d2d] transition-colors whitespace-nowrap overflow-hidden">
              <MessageSquarePlus className="w-5 h-5 flex-shrink-0" />
              {isMenuExpanded && <span className="ml-3 text-sm">New Chat</span>}
            </button>

            {/* History section */}
            <div className="pt-4">
              <div className="px-3 mb-2 flex items-center whitespace-nowrap overflow-hidden">
                <History className="w-5 h-5 text-gray-400 flex-shrink-0" />
                {isMenuExpanded && (
                  <span className="ml-3 text-sm text-gray-400">History</span>
                )}
              </div>
              {/* Placeholder for chat history items */}
              <div className="mt-2 space-y-1">
                {isMenuExpanded && (
                  <>
                    <div className="px-3 py-1 text-xs text-gray-500 hover:bg-[#2d2d2d] cursor-pointer whitespace-nowrap overflow-hidden">
                      Previous chat 1
                    </div>
                    <div className="px-3 py-1 text-xs text-gray-500 hover:bg-[#2d2d2d] cursor-pointer whitespace-nowrap overflow-hidden">
                      Previous chat 2
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Settings at bottom */}
        <div className="p-3">
          <button className="w-full flex items-center text-gray-400 hover:text-white transition-colors whitespace-nowrap overflow-hidden">
            <Settings className="w-5 h-5 flex-shrink-0" />
            {isMenuExpanded && <span className="ml-3 text-sm">Settings</span>}
          </button>
        </div>
      </div>

      {/* Spreadsheet View */}
      <div className="flex-1 bg-[#2d2d2d] flex flex-col">
        {/* Header */}
        <div className="p-4 flex items-center justify-between flex-shrink-0">
          <div className="flex items-center space-x-3">
            <button className="text-gray-400 hover:text-white">
              <Menu className="w-5 h-5" />
            </button>
            <h2 className="text-white font-medium">Sources</h2>
          </div>
          {currentFile && (
            <div className="flex items-center space-x-2">
              <button
                className="text-gray-400 hover:text-white p-2"
                onClick={handleOpenInNewTab}
              >
                <ExternalLink className="w-4 h-4" />
              </button>
              <button
                className="text-gray-400 hover:text-white"
                onClick={() => {
                  console.log("Closing file:", currentFile);
                  setCurrentFile(null);
                }}
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          )}
        </div>

        {/* Main Content Area */}
        <div className="flex-1 relative">
          {currentFile ? (
            <WOPIViewer
              fileName={currentFile}
              onHighlightCells={(cells) => {
                console.log("Cells highlighted:", cells);
              }}
            />
          ) : (
            <div
              className={`h-full flex flex-col items-center justify-center p-8 ${
                isDragging ? "bg-gray-500/10" : ""
              }`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <div className="bg-[#1f1f1f] p-8 rounded-2xl text-center max-w-md w-full">
                <FileSpreadsheet className="w-16 h-16 mx-auto mb-4 text-gray-500" />
                <h3 className="text-xl font-semibold text-white mb-2">
                  Upload Excel File
                </h3>
                <p className="text-gray-400 mb-6">
                  Drag and drop your Excel file here, or click to select
                </p>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleExcelUpload}
                  className="hidden"
                />
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="bg-gray-600 text-white px-6 py-3 rounded-lg hover:bg-gray-700 transition-colors"
                >
                  Choose File
                </button>
                {uploadStatus && (
                  <p className="mt-4 text-sm text-gray-400">{uploadStatus}</p>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
