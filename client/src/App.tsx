import { Toaster } from "@/components/ui/sonner";
import ErrorBoundary from "./components/ErrorBoundary";
import Home from "./pages/Home";

function App() {
  return (
    <ErrorBoundary>
      <Toaster position="top-left" richColors />
      <Home />
    </ErrorBoundary>
  );
}

export default App;
