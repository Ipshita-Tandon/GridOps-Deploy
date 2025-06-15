export interface ResponseSegment {
  text: string;
  citation: {
    cells: string[];  // Array of cell references like ["A1", "B2:B5"]
    sheet?: string;   // Sheet name if applicable
  };
}

export interface StructuredResponse {
  text: string;
  segments: ResponseSegment[];
}

// Single mock response with precise cell references
export const MOCK_RESPONSES: Record<string, StructuredResponse> = {
  default: {
    text: "The quarterly report shows significant trends in our key metrics. Our revenue growth [1] has exceeded expectations, particularly in Q4. The customer acquisition data [2] indicates strong market penetration. Looking at the satisfaction scores [3], we've maintained high standards throughout the year. The comprehensive market analysis [4] provides detailed insights across all segments.",
    segments: [
      {
        text: "has exceeded expectations, particularly in Q4",
        citation: {
          cells: ["D15"],  // Single cell reference
          sheet: "Sheet1"
        }
      },
      {
        text: "indicates strong market penetration",
        citation: {
          cells: ["B8"],  // Single merged cell reference
          sheet: "Sheet1"
        }
      },
      {
        text: "we've maintained high standards throughout the year",
        citation: {
          cells: ["B13:B15"],  // Column range
          sheet: "Sheet1"
        }
      },
      {
        text: "provides detailed insights across all segments",
        citation: {
          cells: ["B9:K9"],  // Full row range
          sheet: "Sheet1"
        }
      }
    ]
  }
}; 