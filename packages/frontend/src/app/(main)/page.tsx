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
        <h1 className="text-5xl font-bold tracking-tight text-gray-900">Plagiarism Checker</h1>
        <p className="text-lg text-gray-600 max-w-2xl mx-auto">
          Ensure every word is your own with Grammarly's AI-powered plagiarism checker, which uses advanced AI to detect
          plagiarism in your text and check for other writing issues.
        </p>
      </section>

      {/* Main Input and Guide Section */}
      <section className="mb-24">
        {/* Increased bottom margin */}
        <InputSection />
        {(status === "uploading" || status === "analyzing") && (
          <div className="mt-8 space-y-2 text-center">
            <Progress value={progress} className="w-full max-w-md mx-auto" />
            <p className="text-sm text-gray-500">
              {status === "uploading" ? `Uploading... ${progress}%` : "Analyzing document..."}
            </p>
          </div>
        )}
        {status === "error" && error && (
          <Alert variant="destructive" className="mt-8 max-w-md mx-auto">
            <Terminal className="h-4 w-4" />
            <AlertTitle>Error</AlertTitle>
            <AlertDescription>{error}</AlertDescription>
            <button
              onClick={reset}
              className="mt-2 inline-flex items-center justify-center rounded-md text-sm font-medium ring-offset-background transition-colors hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:pointer-events-none disabled:opacity-50 h-10 px-4 py-2"
            >
              Try Again
            </button>
          </Alert>
        )}
      </section>

      {/* Feature Section */}
      <section className="space-y-8 py-16 bg-gray-50">
        {/* Added grey background and padding */}
        <div className="text-center space-y-4">
          <h2 className="text-4xl font-bold tracking-tight text-gray-900">
            Make the grade with our AI plagiarism checker
          </h2>
          <p className="text-lg text-gray-600 max-w-3xl mx-auto">
            Your ideas are unique, and your writing should reflect that. Grammarly's AI-powered plagiarism detector makes
            it easy to express your thoughts in a way that's clear, original, and full of academic integrity.
          </p>
        </div>
        <CardsSection />
      </section>

      {/* Tabbed Section */}
      <section className="space-y-8">
        <div className="text-center space-y-4">
          <h2 className="text-4xl font-bold tracking-tight text-gray-900">
            Beyond plagiarism detection: Speed up your work
          </h2>
          <p className="text-lg text-gray-600 max-w-3xl mx-auto">
            Go beyond plagiarism detection to make your writing shine. From final papers to internship applications,
            Grammarly's AI writing assistance improves your writing and teaches you how to use generative AI responsibly
            so you're a step ahead at school and when entering the workforce.
          </p>
        </div>
        <TabsSection />
      </section>
    </div>
  );
} 