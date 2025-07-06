import type React from "react"
import { UploadCloud, FileText } from "lucide-react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { useFileUpload } from "@/hooks/use-file-upload"

interface FileUploadProps {
  onFileChange: (file: File | null) => void
  onSubmit: () => void
}

export function FileUpload({ onFileChange, onSubmit }: FileUploadProps) {
  const {
    selectedFile,
    isDragActive,
    getRootProps,
    getInputProps,
    handleFileSelect,
    handleRemoveFile,
  } = useFileUpload(onFileChange)

  return (
    <div className="flex items-center justify-center w-full">
      <div
        {...getRootProps()}
        className="flex flex-col items-center justify-center w-full p-6 border-2 border-dashed border-gray-300 rounded-lg cursor-pointer bg-gray-50 hover:border-gray-400 transition-colors"
      >
        <Input {...getInputProps()} id="file-upload" className="hidden" />
        <UploadCloud className="h-12 w-12 text-gray-400" />
        <p className="mt-2 text-sm text-gray-600">
          {isDragActive ? "Drop the file here ..." : "Drag 'n' drop a document here, or click to select one"}
        </p>
        <p className="text-xs text-gray-500">PDF, DOCX, TXT up to 10MB</p>
      </div>

      {selectedFile && (
        <div className="flex items-center justify-between p-3 border border-gray-200 rounded-md bg-gray-50 mt-4 w-full">
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

      <div className="flex justify-center w-full mt-6">
        <Button onClick={onSubmit} disabled={!selectedFile} className="w-full sm:w-auto">
          Upload Document
        </Button>
      </div>
    </div>
  )
} 