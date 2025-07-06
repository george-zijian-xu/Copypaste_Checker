export function useDocumentAnalysis() {
  const handleSubmit = async (input: string | File) => {
    // TODO: implement actual API call to backend
    console.log("Submitting for analysis", input);
  };

  return { handleSubmit };
} 