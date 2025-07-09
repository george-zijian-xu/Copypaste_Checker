"use client";

import { Upload, FileText, Trash2 } from "lucide-react";
import {
  Card,
  CardContent,
  CardFooter,
} from "@/components/ui/card";
import { Textarea } from "@/components/ui/textarea";
import { Button } from "@/components/ui/button";
import { useInput, MAX_CHARACTERS } from "@/hooks/copy-checker";
import { useDocumentStore } from "@/store/document-store";

export default function InputSection() {
  const {
    file,
    textInput,
    charCount,
    getRootProps,
    getInputProps,
    handleTextChange,
    handleFileUploadClick,
    handleFileSelect,
    handleScanClick,
  } = useInput();

  const { status, highlightedHtml, reset } = useDocumentStore();
  const isDone = status === "done";

  return (
    <div className="grid grid-cols-1">
      {/* Text / file input card */}
      <Card className="p-0 border-gray-200 shadow-sm" {...getRootProps()}>
        <input {...getInputProps()} />
        <CardContent className="p-0">
          {isDone && highlightedHtml ? (
            <div
              className="min-h-[300px] w-full overflow-auto rounded-lg border-0 p-6 text-base"
              style={{ whiteSpace: "pre-wrap" }}
              dangerouslySetInnerHTML={{ __html: highlightedHtml }}
            />
          ) : (
            <Textarea
              placeholder="Enter text or upload file to check for plagiarism and writing errors."
              className="w-full h-64 resize-none border-none focus-visible:ring-0 focus-visible:ring-offset-0 rounded-b-none text-base p-6"
              value={textInput}
              onChange={handleTextChange}
              maxLength={MAX_CHARACTERS}
              aria-label="Enter text for plagiarism check"
            />
          )}
        </CardContent>
        <CardFooter className="flex items-center justify-between p-4 border-t border-gray-200 bg-gray-50 rounded-b-lg">
          <div className="flex items-center space-x-4">
            {isDone ? (
              <Button
                className="bg-gray-500 hover:bg-gray-600/90 text-white px-6 py-2 rounded-md font-semibold transition-colors"
                onClick={reset}
              >
                <Trash2 className="mr-2 h-4 w-4" />
                Start New Check
              </Button>
            ) : (
              <>
                <Button
                  className="bg-green-050 hover:bg-green-060/90 active:bg-green-060 text-white px-6 py-2 rounded-md font-semibold transition-colors"
                  onClick={handleScanClick}
                >
                  Scan for plagiarism
                </Button>
                <input
                  id="file-upload-input"
                  type="file"
                  className="hidden"
                  onChange={handleFileSelect}
                  accept=".pdf,.doc,.docx,.txt"
                />
                <Button
                  variant="ghost"
                  className="text-gray-600 hover:text-green-060 hover:bg-transparent px-0 transition-colors"
                  onClick={handleFileUploadClick}
                >
                  <Upload className="mr-2 h-4 w-4" />
                  Upload a File
                </Button>
              </>
            )}
          </div>
          <div className="text-sm text-gray-500">
            {file && !isDone ? (
              <span className="flex items-center">
                <FileText className="h-4 w-4 mr-1" /> {file.name} ({(file.size / 1024 / 1024).toFixed(2)} MB)
              </span>
            ) : (
              !isDone && `${charCount}/${MAX_CHARACTERS}`
            )}
          </div>
        </CardFooter>
      </Card>
    </div>
  );
} 