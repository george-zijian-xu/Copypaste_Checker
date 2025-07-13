"use client";

import { InputSection } from "@/components/feature/copy-checker";
import { CardsSection } from "@/components/feature/copy-checker";
import { TabsSection } from "@/components/feature/copy-checker";
import { useDocumentStore } from "@/store/document-store";
import { Progress } from "@/components/ui/progress";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Terminal } from "lucide-react";

export default function PlagiarismCheckerPage() {
  const { status, error, progress, reset } = useDocumentStore();

  return (
    <div className="w-full max-w-6xl mx-auto px-4 py-12 space-y-16">
      {/* Hero Section */}
      <section className="text-center space-y-4">
        <h1 className="text-5xl font-bold tracking-tight text-gray-900">Copy Forensics</h1>
        <p className="text-lg text-gray-600 max-w-2xl mx-auto">
        Uncover every single paste event with Word's own revision metadata—no black-box guesses, just clear evidence of what was pasted.
        </p>
      </section>

      {/* Status and Progress */}
      {(status === 'uploading' || status === 'analyzing') && (
        <div className="space-y-4">
          <div className="text-center">
            <p className="text-sm text-gray-600">
              {status === 'uploading' ? 'Uploading document...' : 'Analyzing document...'}
            </p>
          </div>
          <Progress value={progress} className="w-full" />
        </div>
      )}

      {status === 'error' && error && (
        <Alert variant="destructive">
          <Terminal className="h-4 w-4" />
          <AlertTitle>Error</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      )}

      {/* Main Content */}
      <InputSection />
      
      {/* Results Section - only show if we have results */}
      {status === 'done' && (
        <>
          <CardsSection />
          <TabsSection />
        </>
      )}
    </div>
  );
} 