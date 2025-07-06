"use client";

import { Upload, FileText, CheckCircle2 } from "lucide-react";
import {
  Card,
  CardContent,
  CardFooter,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Textarea } from "@/components/ui/textarea";
import { Button } from "@/components/ui/button";
import { useInput, MAX_CHARACTERS } from "@/hooks/copy-checker";

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

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
      {/* Text / file input card */}
      <Card className="lg:col-span-2 p-0 border-gray-200 shadow-sm" {...getRootProps()}>
        <input {...getInputProps()} />
        <CardContent className="p-0">
          <Textarea
            placeholder="Enter text or upload file to check for plagiarism and writing errors."
            className="w-full h-64 resize-none border-none focus-visible:ring-0 focus-visible:ring-offset-0 rounded-b-none text-base p-6 shadow-inner ring-1 ring-gray-200"
            value={textInput}
            onChange={handleTextChange}
            maxLength={MAX_CHARACTERS}
            aria-label="Enter text for plagiarism check"
          />
        </CardContent>
        <CardFooter className="flex items-center justify-between p-4 border-t border-gray-200 bg-gray-50 rounded-b-lg">
          <div className="flex items-center space-x-4">
            <Button
              className="bg-green-050 hover:bg-green-060/90 active:bg-green-060 text-white px-6 py-2 rounded-md font-semibold transition-colors"
              onClick={handleScanClick}
            >
              Scan for plagiarism
            </Button>
            {/* hidden manual file upload */}
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
          </div>
          <div className="text-sm text-gray-500">
            {file ? (
              <span className="flex items-center">
                <FileText className="h-4 w-4 mr-1" /> {file.name} ({(file.size / 1024 / 1024).toFixed(2)} MB)
              </span>
            ) : (
              `${charCount}/${MAX_CHARACTERS}`
            )}
          </div>
        </CardFooter>
      </Card>

      {/* Instruction / promotion card */}
      <Card className="lg:col-span-1 border-gray-200 shadow-sm bg-white">
        <CardHeader className="pb-4">
          <CardTitle className="text-lg font-semibold flex items-center">
            <CheckCircle2 className="h-5 w-5 text-green-060 mr-2" />
            Let's get started.
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-3 text-sm text-gray-700">
          <p>
            <span className="font-semibold">Step 1:</span> Add your text or upload a file.
          </p>
          <p>
            <span className="font-semibold">Step 2:</span> Click to scan for plagiarism.
          </p>
          <p>
            <span className="font-semibold">Step 3:</span> Review the results for instances of potential plagiarism, plus additional writing issues.
          </p>
        </CardContent>
        <CardFooter className="pt-4">
          <Button className="w-full bg-green-050 hover:bg-green-060/90 active:bg-green-060 text-white px-6 py-2 rounded-md font-semibold transition-colors">
            Get Grammarly Pro
          </Button>
        </CardFooter>
      </Card>
    </div>
  );
} 