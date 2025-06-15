import React, { createContext, useContext, useState } from "react";

interface SpreadsheetContextType {
  currentFile: string | null;
  setCurrentFile: (file: string | null) => void;
  highlightedCells: string[] | null;
  setHighlightedCells: (cells: string[] | null) => void;
  focusedCells: { cells: string[]; sheet?: string } | null;
  setFocusedCells: (cells: { cells: string[]; sheet?: string } | null) => void;
  fileId: string | null;
  setFileId: (id: string | null) => void;
}

const SpreadsheetContext = createContext<SpreadsheetContextType | undefined>(
  undefined
);

export const SpreadsheetProvider: React.FC<{ children: React.ReactNode }> = ({
  children,
}) => {
  const [currentFile, setCurrentFile] = useState<string | null>(null);
  const [highlightedCells, setHighlightedCells] = useState<string[] | null>(
    null
  );
  const [focusedCells, setFocusedCells] = useState<{
    cells: string[];
    sheet?: string;
  } | null>(null);
  const [fileId, setFileId] = useState<string | null>(null);

  return (
    <SpreadsheetContext.Provider
      value={{
        currentFile,
        setCurrentFile,
        highlightedCells,
        setHighlightedCells,
        focusedCells,
        setFocusedCells,
        fileId,
        setFileId,
      }}
    >
      {children}
    </SpreadsheetContext.Provider>
  );
};

export const useSpreadsheet = () => {
  const context = useContext(SpreadsheetContext);
  if (context === undefined) {
    throw new Error("useSpreadsheet must be used within a SpreadsheetProvider");
  }
  return context;
};
