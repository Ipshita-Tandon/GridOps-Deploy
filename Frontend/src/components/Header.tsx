import { Bell, HelpCircle, Share, User } from "lucide-react";

export const Header = () => {
  return (
    <header className="bg-[#2d2d2d] border-b border-gray-700 px-4 py-3 flex items-center justify-between">
      <div className="flex items-center space-x-4">
        <div className="w-6 h-6 flex items-center justify-center">
          <img
            src="/sheets.png"
            alt="Sheets Logo"
            className="w-full h-full object-contain"
          />
        </div>
        <h1 className="text-white font-medium">
          Key Insights from Your Document
        </h1>
        <button className="text-gray-400 hover:text-white">
          <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
            <path
              fillRule="evenodd"
              d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z"
              clipRule="evenodd"
            />
          </svg>
        </button>
      </div>

      <div className="flex items-center space-x-5">
        <button className="bg-[#3a3a3a] text-white px-3 py-1.5 rounded text-sm hover:bg-[#4a4a4a] transition-colors flex items-center space-x-2">
          <Share className="w-4 h-4" />
          <span>Share</span>
        </button>
        {/* <button className="text-gray-400 hover:text-white">
          <HelpCircle className="w-5 h-5" />
        </button>
        <button className="text-gray-400 hover:text-white">
          <Bell className="w-5 h-5" />
        </button> */}
        <button className="w-8 h-8 bg-purple-500 rounded-full flex items-center justify-center">
          <User className="w-4 h-4 text-white" />
        </button>
      </div>
    </header>
  );
};
