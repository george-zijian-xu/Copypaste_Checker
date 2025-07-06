import { useCallback, useEffect, useState } from "react";
import { useDropzone } from "react-dropzone";
import { useDocumentStore } from "@/store/document-store";
import { useDocumentAnalysis } from "@/hooks/use-document-analysis";

export const MAX_CHARACTERS = 30000;

interface UsePlagiarismInputReturn {
  file: File | null;
  textInput: string;
  charCount: number;
  getRootProps: ReturnType<typeof useDropzone>["getRootProps"];
  getInputProps: ReturnType<typeof useDropzone>["getInputProps"];
  isDragActive: boolean;
  handleTextChange: (event: React.ChangeEvent<HTMLTextAreaElement>) => void;
  handleFileUploadClick: () => void;
  handleFileSelect: (event: React.ChangeEvent<HTMLInputElement>) => void;
  handleScanClick: () => void;
}

export function usePlagiarismInput(): UsePlagiarismInputReturn {
  const { file, setFile, setStatus, setProgress } = useDocumentStore();
  const { handleSubmit } = useDocumentAnalysis();

  const [textInput, setTextInput] = useState<string>("");
  const [charCount, setCharCount] = useState<number>(0);

  const onDrop = useCallback(
    (acceptedFiles: File[]) => {
      if (acceptedFiles.length > 0) {
        setFile(acceptedFiles[0]);
        setTextInput("");
      }
    },
    [setFile]
  );

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    multiple: false,
    noClick: true,
    accept: {
      "application/pdf": [".pdf"],
      "application/msword": [".doc"],
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
      "text/plain": [".txt"],
    },
  });

  const handleTextChange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    const text = event.target.value;
    if (text.length <= MAX_CHARACTERS) {
      setTextInput(text);
      setCharCount(text.length);
      setFile(null);
    }
  };

  const handleFileUploadClick = () => {
    const fileInput = document.getElementById("file-upload-input");
    if (fileInput) {
      fileInput.click();
    }
  };

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files.length > 0) {
      setFile(event.target.files[0]);
      setTextInput("");
    } else {
      setFile(null);
    }
  };

  const handleScanClick = () => {
    if (textInput.trim() || file) {
      const input = textInput.trim() || file;
      if (input) {
        handleSubmit(input);
      }
    }
  };

  // Reset relevant state when component using this hook unmounts
  useEffect(() => {
    return () => {
      setFile(null);
      setTextInput("");
      setCharCount(0);
      setStatus("idle");
      setProgress(0);
    };
  }, [setFile, setStatus, setProgress]);

  return {
    file,
    textInput,
    charCount,
    getRootProps,
    getInputProps,
    isDragActive,
    handleTextChange,
    handleFileUploadClick,
    handleFileSelect,
    handleScanClick,
  };
} 
 
 
 
 
 
 
 
 
 
 
 