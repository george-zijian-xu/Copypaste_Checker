import { useCallback, useState } from "react";
import { useDropzone } from "react-dropzone";

export function useFileUpload(onFileChange: (file: File | null) => void) {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const onDrop = useCallback(
    (acceptedFiles: File[]) => {
      if (acceptedFiles.length > 0) {
        setSelectedFile(acceptedFiles[0]);
        onFileChange(acceptedFiles[0]);
      }
    },
    [onFileChange]
  );

  const {
    getRootProps,
    getInputProps,
    isDragActive,
  } = useDropzone({
    onDrop,
    multiple: false,
    accept: {
      "application/pdf": [".pdf"],
      "application/msword": [".doc"],
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
      "text/plain": [".txt"],
    },
  });

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files.length > 0) {
      setSelectedFile(event.target.files[0]);
      onFileChange(event.target.files[0]);
    } else {
      setSelectedFile(null);
      onFileChange(null);
    }
  };

  const handleRemoveFile = () => {
    setSelectedFile(null);
    onFileChange(null);
  };

  return {
    selectedFile,
    isDragActive,
    getRootProps,
    getInputProps,
    handleFileSelect,
    handleRemoveFile,
  };
} 