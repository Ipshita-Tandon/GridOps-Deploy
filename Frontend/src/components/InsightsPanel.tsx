import { useState, useEffect, useCallback, useRef } from "react";
import {
  Send,
  Square,
  Mic,
  Menu,
  FileSpreadsheet,
  MessageSquarePlus,
  History,
  Settings,
  X,
} from "lucide-react";
import { useSpreadsheet } from "../contexts/SpreadsheetContext";
import {
  ResponseSegment as MockResponseSegment,
  MOCK_RESPONSES,
} from "../mocks/responses";
import config from "../config";

// TypeScript declarations for Web Speech API
interface SpeechRecognitionEvent extends Event {
  results: SpeechRecognitionResultList;
}

interface SpeechRecognitionResultList {
  length: number;
  item(index: number): SpeechRecognitionResult;
  [index: number]: SpeechRecognitionResult;
}

interface SpeechRecognitionResult {
  isFinal: boolean;
  [index: number]: SpeechRecognitionAlternative;
}

interface SpeechRecognitionAlternative {
  transcript: string;
  confidence: number;
}

interface SpeechRecognition extends EventTarget {
  continuous: boolean;
  interimResults: boolean;
  start(): void;
  stop(): void;
  onresult: (event: SpeechRecognitionEvent) => void;
  onend: () => void;
}

declare global {
  interface Window {
    SpeechRecognition: new () => SpeechRecognition;
    webkitSpeechRecognition: new () => SpeechRecognition;
  }
}

interface Message {
  id: string;
  role: "user" | "assistant";
  content: string;
  isTyping?: boolean;
  segments?: ResponseSegment[];
  attributions?: ResponseData[];
}

interface ResponseSegment {
  text: string;
  citation: {
    cells: string[]; // Array of cell references like ["A1", "B2:B5"]
    sheet?: string; // Sheet name if applicable
  };
}

interface CitationProps {
  index: number;
  citation: MockResponseSegment["citation"];
  onHover: (cells: string[] | null) => void;
  onMouseLeave: () => void;
  onClick: (cells: string[], sheet?: string) => void;
}

const Citation: React.FC<CitationProps> = ({
  index,
  citation,
  onHover,
  onMouseLeave,
  onClick,
}) => {
  return (
    <span
      className="inline-flex items-center justify-center w-5 h-5 text-xs font-medium bg-blue-600/20 text-blue-400 rounded cursor-pointer mx-1 hover:bg-blue-600/30 transition-colors"
      onMouseEnter={() => onHover(citation.cells)}
      onMouseLeave={onMouseLeave}
      onClick={() => onClick(citation.cells, citation.sheet)}
    >
      {index + 1}
    </span>
  );
};

const useTypewriter = (
  text: string,
  speed: number = 10,
  shouldStop: boolean = false,
  messageId: string
) => {
  const [displayedText, setDisplayedText] = useState("");
  const [isComplete, setIsComplete] = useState(false);
  const intervalRef = useRef<NodeJS.Timeout | null>(null);
  const messageIdRef = useRef(messageId);

  useEffect(() => {
    // Reset if message ID changes
    if (messageId !== messageIdRef.current) {
      messageIdRef.current = messageId;
      setDisplayedText("");
      setIsComplete(false);
    }

    // If already complete or stopped, show full text
    if (shouldStop || isComplete) {
      setDisplayedText(text);
      setIsComplete(true);
      return;
    }

    let i = 0;
    intervalRef.current = setInterval(() => {
      if (i < text.length) {
        const charsToAdd = text.slice(i, i + 3);
        setDisplayedText((prev) => prev + charsToAdd);
        i += 3;
      } else {
        setIsComplete(true);
        if (intervalRef.current) {
          clearInterval(intervalRef.current);
        }
      }
    }, speed);

    return () => {
      if (intervalRef.current) {
        clearInterval(intervalRef.current);
      }
    };
  }, [text, speed, shouldStop, messageId]);

  return { displayedText, isComplete };
};

const formatMarkdown = (text: string): string => {
  // Ensure text is a string
  if (typeof text !== "string") {
    console.warn("formatMarkdown received non-string input:", text);
    return String(text);
  }

  // Split into paragraphs and format each one
  console.log("Text:", text);
  const paragraphs = text.split("\n\n");
  const formattedParagraphs = paragraphs.map((paragraph) => {
    // Add bullet point if not already present
    if (!paragraph.trim().startsWith("•")) {
      paragraph = `• ${paragraph}`;
    }
    return paragraph;
  });

  return formattedParagraphs.join("\n\n");
};

// Add this helper function at the top level
const extractHighlightCommand = (message: string) => {
  // Match patterns like "highlight A1:B2 in Sheet1" or "highlight A1,B2,C3 in Sheet1"
  const regex = /highlight\s+([A-Z0-9:,\s]+)\s+in\s+(.+)/i;
  const match = message.match(regex);
  if (match) {
    return {
      cellRanges: match[1].trim(),
      sheetName: match[2].trim(),
    };
  }
  return null;
};

interface ResponseData {
  answer: string;
  attribution?: string[];
}

const formatResponse = (
  result: any
): { text: string; references: string[] } => {
  if (typeof result === "string") {
    try {
      result = JSON.parse(result);
    } catch (e) {
      console.warn("Failed to parse response as JSON:", e);
      return { text: result, references: [] };
    }
  }

  let references: string[] = [];
  let formattedText = "";

  if (Array.isArray(result)) {
    formattedText = result
      .map((item) => {
        if (item.answer) {
          references.push(item.answer);
          return `• ${item.answer.trim()}`;
        }
        return "";
      })
      .filter(Boolean)
      .join("\n");
  } else if (typeof result === "object" && result.answer) {
    references.push(result.answer);
    formattedText = `• ${result.answer.trim()}`;
  } else {
    formattedText = String(result);
  }

  return { text: formattedText, references };
};

export const InsightsPanel = () => {
  const [message, setMessage] = useState("");
  const [messages, setMessages] = useState<Message[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [stopTyping, setStopTyping] = useState(false);
  const [isListening, setIsListening] = useState(false);
  const recognitionRef = useRef<SpeechRecognition | null>(null);
  const {
    setHighlightedCells,
    setFocusedCells,
    currentFile,
    fileId,
    setCurrentFile,
    setFileId,
  } = useSpreadsheet();
  const [attributions, setAttributions] = useState<ResponseData[]>([]);

  // Function to get mock response based on query
  const getMockResponse = (query: string) => {
    return MOCK_RESPONSES.default;
  };

  // Function to handle citation hover
  const handleCitationHover = (cells: string[] | null) => {
    setHighlightedCells(cells);
  };

  // Function to handle citation click
  const handleCitationClick = async (cells: string[], sheet?: string) => {
    if (!cells || cells.length === 0 || !currentFile || !fileId) return;

    try {
      // Find the matching attribution to get the correct sheet name
      const matchingAttribution = attributions.find(
        (attr) => attr.attribution && attr.attribution.includes(cells[0])
      );

      if (!matchingAttribution) {
        throw new Error("Could not find matching attribution for these cells");
      }

      // Get the sheet name from the attribution (it's usually the first element)
      const sheetName = matchingAttribution.attribution[0];
      if (!sheetName) {
        throw new Error("No sheet name found in attribution");
      }

      console.log("Highlighting cells:", {
        fileId,
        sheetName,
        cellRanges: cells.join(","),
      });

      // Call the highlight endpoint
      const response = await fetch(`${config.apiUrl}/highlight`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          file_id: fileId,
          sheet_name: sheetName,
          cell_ranges: cells.join(","),
        }),
      });

      if (!response.ok) {
        throw new Error(`Failed to highlight cells: ${await response.text()}`);
      }

      const result = await response.json();
      if (result.success && result.file_id && result.filename) {
        // Update the current file ID and name
        setFileId(result.file_id);
        setCurrentFile(result.filename);

        // Notify the SourcesPanel to load the new file
        const event = new CustomEvent("loadHighlightedFile", {
          detail: {
            fileName: result.filename,
            fileId: result.file_id,
          },
        });
        window.dispatchEvent(event);

        const assistantMessage: Message = {
          id: Date.now().toString() + "-assistant",
          role: "assistant",
          content: `I've highlighted the cells ${cells.join(
            ", "
          )} in sheet "${sheetName}". The Excel viewer should update momentarily.`,
          isTyping: true,
        };
        setMessages((prev) => [...prev, assistantMessage]);
      } else {
        throw new Error("Invalid response from server");
      }
    } catch (error) {
      console.error("Error highlighting cells:", error);
      const errorMessage: Message = {
        id: Date.now().toString() + "-error",
        role: "assistant",
        content: `Sorry, there was an error highlighting the cells: ${
          error instanceof Error ? error.message : "Unknown error"
        }`,
        isTyping: true,
      };
      setMessages((prev) => [...prev, errorMessage]);
    }
  };

  // Initialize speech recognition
  useEffect(() => {
    if ("SpeechRecognition" in window || "webkitSpeechRecognition" in window) {
      const SpeechRecognition =
        window.SpeechRecognition || window.webkitSpeechRecognition;
      recognitionRef.current = new SpeechRecognition();
      recognitionRef.current.continuous = true;
      recognitionRef.current.interimResults = true;

      recognitionRef.current.onresult = (event) => {
        const transcript = Array.from(event.results)
          .map((result) => result[0].transcript)
          .join("");
        setMessage(transcript);
      };

      recognitionRef.current.onend = () => {
        setIsListening(false);
      };
    }
  }, []);

  const toggleListening = () => {
    if (!recognitionRef.current) return;

    if (isListening) {
      recognitionRef.current.stop();
      setIsListening(false);
    } else {
      recognitionRef.current.start();
      setIsListening(true);
    }
  };

  const handleStopResponse = () => {
    setStopTyping(true);
    setIsLoading(false);
  };

  const sendMessage = async (messageToSend: string = message) => {
    if (!messageToSend.trim()) return;

    const userMessage: Message = {
      id: Date.now().toString() + "-user",
      role: "user",
      content: messageToSend,
    };

    setMessage("");
    setMessages((prev) => [...prev, userMessage]);
    setIsLoading(true);
    setStopTyping(false);

    try {
      // Check if this is a highlight command
      const highlightCommand = extractHighlightCommand(messageToSend);
      if (highlightCommand) {
        if (!currentFile) {
          const assistantMessage: Message = {
            id: Date.now().toString() + "-assistant",
            role: "assistant",
            content:
              "Please upload an Excel file first before using highlighting commands.",
            isTyping: true,
          };
          setMessages((prev) => [...prev, assistantMessage]);
          return;
        }

        // Call the highlight endpoint
        const response = await fetch("http://localhost:3001/highlight-cells", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            fileName: currentFile,
            sheetName: highlightCommand.sheetName,
            cellRanges: highlightCommand.cellRanges,
          }),
        });

        if (!response.ok) {
          throw new Error(
            `Failed to highlight cells: ${await response.text()}`
          );
        }

        const result = await response.json();
        if (result.highlightedFileName) {
          // Notify the SourcesPanel to load the new file
          const event = new CustomEvent("loadHighlightedFile", {
            detail: { fileName: result.highlightedFileName },
          });
          window.dispatchEvent(event);

          const assistantMessage: Message = {
            id: Date.now().toString() + "-assistant",
            role: "assistant",
            content: `I've highlighted the cells ${highlightCommand.cellRanges} in sheet "${highlightCommand.sheetName}". The Excel viewer should update momentarily.`,
            isTyping: true,
          };
          setMessages((prev) => [...prev, assistantMessage]);
        }
      } else {
        // Send query to remote server
        console.log("Current file:", currentFile);
        console.log("Current fileId:", fileId);
        if (!currentFile || !fileId) {
          console.log("Missing file or fileId");
          const assistantMessage: Message = {
            id: Date.now().toString() + "-assistant",
            role: "assistant",
            content:
              "Please upload an Excel file first before asking questions.",
            isTyping: true,
          };
          setMessages((prev) => [...prev, assistantMessage]);
          return;
        }

        console.log("Sending question to backend:", messageToSend);
        console.log("File ID:", fileId);
        const requestBody = {
          file_id: fileId,
          question: messageToSend,
        };
        console.log("Request payload:", requestBody);
        const response = await fetch(`${config.apiUrl}/qna`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Accept: "application/json",
            "Access-Control-Allow-Origin": "*",
          },
          body: JSON.stringify(requestBody),
        });

        console.log("Response status:", response.status);
        console.log(
          "Response headers:",
          Object.fromEntries(response.headers.entries())
        );
        if (!response.ok) {
          const errorText = await response.text();
          console.error("Error response:", errorText);
          throw new Error(`Query failed: ${response.status} - ${errorText}`);
        }

        const result = await response.json();
        console.log("Response data:", result);

        // Store attributions for later use
        const attributions: ResponseData[] = [];
        let formattedText = "";

        // Check if result has the answer field from the backend response
        if (result.answer && Array.isArray(result.answer)) {
          formattedText = result.answer
            .map((item) => {
              if (item.answer) {
                if (item.attribution) {
                  attributions.push({
                    answer: item.answer,
                    attribution: item.attribution,
                  });
                }
                return item.answer.trim();
              }
              return "";
            })
            .filter(Boolean)
            .join("\n");
        } else {
          console.error("Unexpected response format:", result);
          formattedText = "Sorry, I couldn't process the response properly.";
        }

        setAttributions(attributions);

        const assistantMessage: Message = {
          id: Date.now().toString() + "-assistant",
          role: "assistant",
          content: formattedText,
          isTyping: true,
          segments: attributions.map((item, index) => ({
            text: item.answer,
            citation: {
              cells: item.attribution || [],
              sheet: "",
            },
          })),
        };
        setMessages((prev) => [...prev, assistantMessage]);
      }
    } catch (error) {
      console.error("Error:", error);
      const errorMessage: Message = {
        id: Date.now().toString() + "-error",
        role: "assistant",
        content: `Sorry, there was an error: ${
          error instanceof Error ? error.message : "Unknown error"
        }`,
        isTyping: true,
      };
      setMessages((prev) => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  };

  const handlePromptClick = (prompt: string) => {
    sendMessage(prompt);
  };

  console.log("Message:", messages);
  const MessageContent = ({
    message,
    onCitationHover,
    onCitationClick,
  }: {
    message: Message;
    onCitationHover?: (cells: string[] | null) => void;
    onCitationClick?: (cells: string[], sheet?: string) => void;
  }) => {
    const [highlightedRef, setHighlightedRef] = useState<number | null>(null);

    if (message.role === "assistant" && message.segments) {
      let lastIndex = 0;
      const parts: JSX.Element[] = [];

      message.segments.forEach((segment, index) => {
        const segmentStart = message.content.indexOf(segment.text, lastIndex);
        if (segmentStart > lastIndex) {
          parts.push(
            <span key={`text-${index}`}>
              {message.content.slice(lastIndex, segmentStart)}
            </span>
          );
        }
        parts.push(
          <span
            key={`segment-${index}`}
            className={`px-1 rounded transition-colors ${
              highlightedRef === index ? "bg-blue-600/10" : ""
            }`}
          >
            {segment.text}
            <Citation
              index={index}
              citation={segment.citation}
              onHover={(cells) => {
                setHighlightedRef(index);
                onCitationHover?.(cells);
              }}
              onMouseLeave={() => {
                setHighlightedRef(null);
                onCitationHover?.(null);
              }}
              onClick={onCitationClick || (() => {})}
            />
          </span>
        );
        lastIndex = segmentStart + segment.text.length;
      });

      if (lastIndex < message.content.length) {
        parts.push(
          <span key="text-end">{message.content.slice(lastIndex)}</span>
        );
      }

      return <div className="whitespace-pre-wrap">{parts}</div>;
    }

    // Split text into parts based on reference numbers
    const parts = message.content.split(/(\[\d+\])/).map((part, index) => {
      const refMatch = part.match(/\[(\d+)\]/);
      if (refMatch) {
        const refNumber = parseInt(refMatch[1]);
        return (
          <span key={index} className="inline-flex items-center">
            <span
              className={`px-1 rounded transition-colors ${
                highlightedRef === refNumber ? "bg-blue-600/10" : ""
              }`}
            >
              {part.replace(/\[\d+\]/, "")}
            </span>
            <span
              className="inline-flex items-center justify-center w-5 h-5 text-xs font-medium bg-blue-600/20 text-blue-400 rounded cursor-pointer mx-1 hover:bg-blue-600/30 transition-colors"
              onMouseEnter={() => setHighlightedRef(refNumber)}
              onMouseLeave={() => setHighlightedRef(null)}
              onClick={() => {
                // Find the matching attribution and call onCitationClick
                const matchingAttribution = message.attributions?.find(
                  (attr) =>
                    attr.attribution &&
                    attr.attribution.includes(refNumber.toString())
                );
                if (matchingAttribution) {
                  onCitationClick?.(
                    matchingAttribution.attribution.slice(1),
                    matchingAttribution.attribution[0]
                  );
                }
              }}
            >
              {refNumber}
            </span>
          </span>
        );
      }
      return part;
    });

    return <div className="whitespace-pre-wrap">{parts}</div>;
  };

  return (
    <div className="h-full bg-[#1a1a1a] flex flex-col">
      {/* Main Content - Scrollable */}
      <div className="flex-1 overflow-y-auto overflow-x-hidden">
        <div className="max-w-3xl mx-auto p-6">
          {messages.length === 0 ? (
            <>
              {/* Main Question */}
              <div className="text-center mb-8">
                <h1 className="text-2xl font-semibold text-white mb-4 mt-8">
                  What would you like to do next?
                </h1>
                <p className="text-gray-400">
                  Ask a question, pick a focus, or explore key insights from
                  your sources.{" "}
                  <span className="text-blue-400 cursor-pointer hover:underline">
                    See tips
                  </span>
                </p>
              </div>

              {/* Action Cards */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-8">
                <div
                  onClick={() =>
                    handlePromptClick(
                      "Identify patterns across different document types"
                    )
                  }
                  className="bg-[#2d2d2d] p-4 rounded-2xl hover:bg-[#3a3a3a] transition-colors cursor-pointer border border-gray-700"
                >
                  <div className="flex items-start space-x-3">
                    <div className="w-2 h-2 bg-pink-500 rounded-full mt-2"></div>
                    <div>
                      <p className="text-white font-medium mb-1">
                        Identify patterns across different document types
                      </p>
                    </div>
                  </div>
                </div>

                <div
                  onClick={() =>
                    handlePromptClick(
                      "Draft a personal reflection on the themes presented"
                    )
                  }
                  className="bg-[#2d2d2d] p-4 rounded-2xl hover:bg-[#3a3a3a] transition-colors cursor-pointer border border-gray-700"
                >
                  <div className="flex items-start space-x-3">
                    <div className="w-2 h-2 bg-pink-500 rounded-full mt-2"></div>
                    <div>
                      <p className="text-white font-medium mb-1">
                        Draft a personal reflection on the themes presented
                      </p>
                    </div>
                  </div>
                </div>

                <div
                  onClick={() =>
                    handlePromptClick(
                      "Create a report outlining insights from the collection"
                    )
                  }
                  className="bg-[#2d2d2d] p-4 rounded-2xl hover:bg-[#3a3a3a] transition-colors cursor-pointer border border-gray-700"
                >
                  <div className="flex items-start space-x-3">
                    <div className="w-2 h-2 bg-pink-500 rounded-full mt-2"></div>
                    <div>
                      <p className="text-white font-medium mb-1">
                        Create a report outlining insights from the collection
                      </p>
                    </div>
                  </div>
                </div>

                {/* Key Insights Card */}
                <div
                  onClick={() =>
                    handlePromptClick("Summarize this document for me.")
                  }
                  className="bg-[#2d2d2d] p-4 rounded-2xl hover:bg-[#3a3a3a] transition-colors cursor-pointer border border-gray-700"
                >
                  <div className="flex items-start space-x-3">
                    <div className="w-2 h-2 bg-pink-500 rounded-full mt-2"></div>
                    <div>
                      <p className="text-white font-medium mb-1">
                        Summarize this document for me.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </>
          ) : (
            <div className="space-y-4">
              {messages.map((msg, index) => (
                <div
                  key={index}
                  className={`flex ${
                    msg.role === "user" ? "justify-end" : "justify-start"
                  }`}
                >
                  <div
                    className={`max-w-3xl p-4 rounded-3xl ${
                      msg.role === "user"
                        ? "bg-[#2d2d2d] text-white"
                        : "bg-[#28282B] text-gray-100"
                    }`}
                  >
                    <MessageContent
                      message={msg}
                      onCitationHover={handleCitationHover}
                      onCitationClick={handleCitationClick}
                    />
                  </div>
                </div>
              ))}
              {isLoading && (
                <div className="flex justify-start">
                  <div className="bg-[#2d2d2d] text-gray-100 border border-gray-700 p-4 rounded-3xl">
                    <div className="flex items-center space-x-2">
                      <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-500"></div>
                      <span>Thinking...</span>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Chat input section */}
      <div className="flex-shrink-0">
        <div className="max-w-3xl mx-auto p-6">
          <div className="relative">
            <textarea
              value={message}
              onChange={(e) => setMessage(e.target.value)}
              onKeyPress={handleKeyPress}
              placeholder="Ask a question or describe what you'd like to work on"
              className="w-full bg-transparent text-white placeholder-gray-400 border border-gray-600 rounded-3xl px-4 py-3 resize-none focus:outline-none focus:border-blue-500 transition-colors pr-24"
              rows={3}
              disabled={isLoading}
            />

            <div className="absolute bottom-3 right-4 flex items-center space-x-2">
              <button
                onClick={toggleListening}
                className={`p-2 rounded-full transition-colors ${
                  isListening
                    ? "bg-pink-600 text-white"
                    : "bg-[#4a4a4a] text-gray-300 hover:bg-[#5a5a5a] hover:text-white"
                }`}
              >
                <Mic className="w-4 h-4" />
              </button>

              <button
                onClick={isLoading ? handleStopResponse : () => sendMessage()}
                disabled={(!message.trim() && !isLoading) || stopTyping}
                className={`p-2 rounded-full transition-colors ${
                  isLoading
                    ? "bg-red-600 text-white hover:bg-red-700"
                    : message.trim()
                    ? "bg-blue-600 text-white hover:bg-blue-700"
                    : "bg-gray-600 text-gray-400 cursor-not-allowed"
                }`}
              >
                {isLoading ? (
                  <Square className="w-4 h-4" />
                ) : (
                  <Send className="w-4 h-4" />
                )}
              </button>
            </div>
          </div>

          <div className="mt-4 text-xs text-gray-500 text-center">
            Be sure to double-check responses as they may be inaccurate.{" "}
            <span className="text-blue-400 cursor-pointer hover:underline">
              User Disclosures
            </span>
            {" | "}
            <span className="text-blue-400 cursor-pointer hover:underline">
              Generative AI User Guidelines
            </span>
          </div>
        </div>
      </div>
    </div>
  );
};
