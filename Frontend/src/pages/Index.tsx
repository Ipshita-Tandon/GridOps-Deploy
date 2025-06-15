import { SourcesPanel } from "../components/SourcesPanel";
import { InsightsPanel } from "../components/InsightsPanel";
import { Header } from "../components/Header";
import {
  ResizablePanelGroup,
  ResizablePanel,
  ResizableHandle,
} from "../components/ui/resizable";

const Index = () => {
  return (
    <div className="h-screen bg-[#1a1a1a] text-white flex flex-col">
      <Header />
      <div className="flex-1 min-h-0">
        <ResizablePanelGroup direction="horizontal" className="h-full">
          <ResizablePanel
            defaultSize={43}
            minSize={15}
            maxSize={50}
            className="h-full"
          >
            <SourcesPanel />
          </ResizablePanel>
          <ResizableHandle withHandle />
          <ResizablePanel defaultSize={67} className="h-full">
            <InsightsPanel />
          </ResizablePanel>
        </ResizablePanelGroup>
      </div>
    </div>
  );
};

export default Index;
