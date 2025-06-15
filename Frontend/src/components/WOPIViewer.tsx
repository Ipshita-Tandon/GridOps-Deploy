import { useEffect, useState, useMemo } from "react";
import { useSpreadsheet } from "../contexts/SpreadsheetContext";
import config from "../config";

interface WOPIViewerProps {
  fileName: string;
  onHighlightCells?: (cells: string[]) => void;
}

export const WOPIViewer: React.FC<WOPIViewerProps> = ({
  fileName,
  onHighlightCells,
}) => {
  const [status, setStatus] = useState<string>("Loading...");
  const [hostUrl, setHostUrl] = useState<string>("");
  const [error, setError] = useState<string | null>(null);
  const { setFileId } = useSpreadsheet();

  // Memoize the viewer URL to prevent unnecessary re-renders
  const viewerUrl = useMemo(() => {
    if (!hostUrl) return "";
    const wopiSrc = encodeURIComponent(`${hostUrl}/wopi/files/${fileName}`);
    const token = "12345"; // This should be handled more securely in production
    return `https://excel.officeapps.live.com/x/_layouts/xlviewerinternal.aspx?WOPISrc=${wopiSrc}&access_token=${token}`;
  }, [hostUrl, fileName]);

  useEffect(() => {
    const loadExcelViewer = async () => {
      try {
        setError(null);
        console.log("Starting file upload process...");
        setStatus(`Loading ${fileName}...`);

        // First check server health
        console.log("Checking remote server health...");
        const healthResponse = await fetch(`${config.apiUrl}/health`);
        console.log("Health check response:", healthResponse.status);
        if (!healthResponse.ok) {
          throw new Error("Remote server is not available");
        }

        // Upload the file to the remote server
        console.log("Preparing file for upload...");
        const formData = new FormData();

        // Get file from local server
        console.log("Fetching file from local server...");
        const fileResponse = await fetch(
          `${config.wopiUrl}/public/${fileName}`
        );
        if (!fileResponse.ok) {
          throw new Error(
            `Failed to get file from local server: ${fileResponse.status}`
          );
        }

        const fileBlob = await fileResponse.blob();
        console.log("File blob size:", fileBlob.size);
        formData.append("file", fileBlob, fileName);

        console.log("Uploading to remote server...");
        const uploadResponse = await fetch(`${config.apiUrl}/upload`, {
          method: "POST",
          body: formData,
        });

        console.log("Upload response status:", uploadResponse.status);
        if (!uploadResponse.ok) {
          const errorText = await uploadResponse.text();
          throw new Error(
            `Failed to upload file to remote server: ${errorText}`
          );
        }

        const uploadResult = await uploadResponse.json();
        console.log("Upload result:", uploadResult);

        if (!uploadResult.file_id) {
          throw new Error("No file ID received from server");
        }

        console.log("Setting fileId to:", uploadResult.file_id);
        setFileId(uploadResult.file_id);
        setStatus(
          `File uploaded successfully. File ID: ${uploadResult.file_id}`
        );

        // Get the WOPI host URL for local preview
        console.log("Getting WOPI host URL...");
        const res = await fetch(`${config.wopiUrl}/host-url`);
        const data = await res.json();
        const hostUrl = data.hostUrl;
        setHostUrl(hostUrl);
        console.log("WOPI Host URL:", hostUrl);

        if (!hostUrl || hostUrl.includes("localhost")) {
          setStatus(
            "WOPI server not ready. Please ensure the server is running and accessible."
          );
          return;
        }

        setStatus(`Displaying ${fileName}`);
      } catch (error) {
        console.error("Error in loadExcelViewer:", error);
        setError(error.message);
        setStatus(`Error: ${error.message}`);
      }
    };

    loadExcelViewer();
  }, [fileName, setFileId]);

  // Function to highlight cells
  const highlightCells = async (sheetName: string, cellRanges: string) => {
    try {
      setError(null);
      setStatus(`Highlighting cells in ${fileName}, sheet: ${sheetName}...`);

      const response = await fetch(`${config.wopiUrl}/highlight-cells`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          fileName,
          sheetName,
          cellRanges,
        }),
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Server error: ${response.status} - ${errorText}`);
      }

      const result = await response.json();
      if (result.highlightedFileName) {
        setStatus(`Successfully highlighted cells in sheet '${sheetName}'`);
        return result.highlightedFileName;
      }
    } catch (error) {
      setError(error.message);
      setStatus(`Error highlighting cells: ${error.message}`);
      throw error;
    }
  };

  return (
    <div className="flex flex-col h-full">
      <div className="text-sm text-gray-400 p-2">
        {error ? <div className="text-red-400">{error}</div> : status}
      </div>
      <div className="flex-1 relative">
        {viewerUrl && (
          <iframe
            key={viewerUrl} // Add key to force re-render only when URL changes
            src={viewerUrl}
            className="absolute inset-0 w-full h-full border-0"
            title="Excel Viewer"
            sandbox="allow-scripts allow-same-origin allow-popups allow-forms"
          />
        )}
      </div>
    </div>
  );
};
