import { useCallback, useEffect, useState } from "react";
import { useDropzone } from "react-dropzone";
import { useDocumentStore } from "@/store/document-store";
import { useDocumentAnalysis } from "@/hooks/copy-checker/useDocumentAnalysis";

export const MAX_CHARACTERS = 30000;

export function useInput() {
  const { file, setFile, setStatus, setProgress } = useDocumentStore();
  const { handleSubmit } = useDocumentAnalysis();

  const [textInput, setTextInput] = useState("");
  const [charCount, setCharCount] = useState(0);

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

  const handleTextChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const text = e.target.value;
    if (text.length <= MAX_CHARACTERS) {
      setTextInput(text);
      setCharCount(text.length);
      setFile(null);
    }
  };

  const handleFileUploadClick = () => {
    const fileInput = document.getElementById("file-upload-input");
    fileInput?.click();
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
      setTextInput("");
    } else {
      setFile(null);
    }
  };

  const handleScanClick = () => {
    if (file) {
      handleSubmit(file);
    } else if (textInput.trim()) {
      // For now, text input analysis is not implemented
      // TODO: Implement text analysis or create a text-to-file conversion
      console.warn("Text input analysis not yet implemented");
    }
  };

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