"use client";

import type React from "react";
import { UploadCloud, FileText } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { useUpload } from "@/hooks/copy-checker";

interface UploadProps {
  onFileChange: (file: File | null) => void;
  onSubmit: () => void;
}

export default function Upload({ onFileChange, onSubmit }: UploadProps) {
  const {
    selectedFile,
    isDragActive,
    getRootProps,
    getInputProps,
    handleFileSelect,
    handleRemoveFile,
  } = useUpload(onFileChange);

  return (
    <div className="space-y-6">
      <div
        {...getRootProps()}
        className="flex flex-col items-center justify-center p-6 border-2 border-dashed border-gray-300 rounded-lg cursor-pointer hover:border-gray-400 transition-colors duration-200"
      >
        <Input {...getInputProps()} id="file-upload" className="hidden" />
        <UploadCloud className="h-12 w-12 text-gray-400" />
        <p className="mt-2 text-sm text-gray-600">
          {isDragActive ? "Drop the file here ..." : "Drag 'n' drop a document here, or click to select one"}
        </p>
        <p className="text-xs text-gray-500">PDF, DOCX, TXT up to 10MB</p>
      </div>

      {selectedFile && (
        <div className="flex items-center justify-between p-3 border border-gray-200 rounded-md bg-gray-50">
          <div className="flex items-center space-x-2">
            <FileText className="h-5 w-5 text-gray-500" />
            <span className="text-sm font-medium text-gray-700">{selectedFile.name}</span>
            <span className="text-xs text-gray-500">({(selectedFile.size / 1024 / 1024).toFixed(2)} MB)</span>
          </div>
          <Button variant="ghost" size="sm" onClick={handleRemoveFile}>
            Remove
          </Button>
        </div>
      )}

      <div className="flex justify-center">
        <Button onClick={onSubmit} disabled={!selectedFile} className="w-full sm:w-auto">
          Upload Document
        </Button>
      </div>
    </div>
  );
} 