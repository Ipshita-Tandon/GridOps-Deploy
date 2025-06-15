import { useEffect, useRef } from "react";
import { SpreadsheetComponent } from "@syncfusion/ej2-react-spreadsheet";
import "../styles/spreadsheet.css";

export const StandaloneSpreadsheet = () => {
  const spreadsheetRef = useRef(null);

  const onCreated = async () => {
    const savedState = localStorage.getItem("spreadsheetState");
    if (savedState && spreadsheetRef.current) {
      try {
        const { workbookData } = JSON.parse(savedState);
        await spreadsheetRef.current.loadFromJson({ file: workbookData });

        // Clear the saved state to prevent stale data on refresh
        localStorage.removeItem("spreadsheetState");
      } catch (error) {
        console.error("Error loading spreadsheet data:", error);
      }
    }
  };

  return (
    <div className="h-screen w-screen bg-[#2d2d2d]">
      <SpreadsheetComponent
        ref={spreadsheetRef}
        openUrl="https://services.syncfusion.com/react/production/api/spreadsheet/open"
        saveUrl="https://services.syncfusion.com/react/production/api/spreadsheet/save"
        height="100%"
        width="100%"
        showRibbon={true}
        showFormulaBar={true}
        allowChart={true}
        created={onCreated}
      />
    </div>
  );
};
