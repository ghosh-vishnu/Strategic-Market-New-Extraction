import React, { useCallback, useMemo, useRef, useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { useTheme } from "./src/useTheme";
import { useAuth } from "./src/AuthContext";
import ProfileDropdown from "./src/ProfileDropdown";
import { API_BASE_URL, API_ENDPOINTS } from "./src/config";

type FileItem = {
  id: string;
  file: File;
  name: string;
  size: number;
  status: "pending" | "queued" | "converting" | "success" | "done" | "error";
  errorMessage?: string | undefined;
};

type ConversionResult = {
  downloadUrl: string;
  openUrl?: string;
};

// Legacy endpoint paths (for backward compatibility)
const UPLOAD_PATH = (import.meta as any).env?.VITE_UPLOAD_PATH || "/api/upload/";
const CONVERT_PATH = (import.meta as any).env?.VITE_CONVERT_PATH || "/api/convert/";
const PROGRESS_PATH = (import.meta as any).env?.VITE_PROGRESS_PATH || "/api/progress/";
const RESULT_PATH = (import.meta as any).env?.VITE_RESULT_PATH || "/api/result/";


async function uploadFolderToBackend(files: File[], jobId?: string): Promise<{ jobId: string }>{
  const formData = new FormData();
  files.forEach((file) => formData.append("files", file, (file as any).webkitRelativePath || file.name));
  const url = jobId ? `${API_BASE_URL}${UPLOAD_PATH}?jobId=${encodeURIComponent(jobId)}` : `${API_BASE_URL}${UPLOAD_PATH}`;
  const res = await fetch(url , { method: "POST", body: formData });
  if (!res.ok) throw new Error("Upload failed");
  return res.json();
}

async function uploadExcelSheet(excelFile: File, jobId: string): Promise<{ success: boolean; message: string; entries: number }>{
  const formData = new FormData();
  formData.append("excelFile", excelFile);
  formData.append("jobId", jobId);
  
  const res = await fetch(API_ENDPOINTS.UPLOAD_EXCEL, { 
    method: "POST", 
    body: formData 
  });
  if (!res.ok) throw new Error("Excel upload failed");
  return res.json();
}

async function uploadExtractExcel(excelFile: File, jobId: string): Promise<{ success: boolean; message: string; entries: number }>{
  const formData = new FormData();
  formData.append("excelFile", excelFile);
  formData.append("jobId", jobId);
  
  const res = await fetch(API_ENDPOINTS.UPLOAD_EXTRACT_EXCEL, { 
    method: "POST", 
    body: formData 
  });
  if (!res.ok) throw new Error("Extract Excel upload failed");
  return res.json();
}


function chunkArray<T>(arr: T[], size: number): T[][] {
  const chunks: T[][] = [];
  for (let i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}

async function startBackendConversion(jobId: string): Promise<{ started: boolean }>{
  const url = `${API_BASE_URL}${CONVERT_PATH}?jobId=${encodeURIComponent(jobId)}`;
  const res = await fetch(url, { method: "POST" });
  if (!res.ok) throw new Error("Failed to start conversion");
  return res.json();
}

async function pollConversionProgress(jobId: string): Promise<{ progress: number; done: boolean; error?: string; status_message?: string }>{
  const url = `${API_BASE_URL}${PROGRESS_PATH}?jobId=${encodeURIComponent(jobId)}`;
  const res = await fetch(url);
  if (!res.ok) {
    const errorText = await res.text();
    console.error(`Progress check failed: ${res.status} - ${errorText}`);
    throw new Error(`Progress check failed: ${res.status}`);
  }
  return res.json();
}

async function fetchConversionResult(jobId: string): Promise<ConversionResult>{
  // Prefer CSV text if user wants to open in a new tab easily; keep xlsx fallback by toggling query
  const endpointUrl = `${API_BASE_URL}${RESULT_PATH}?jobId=${encodeURIComponent(jobId)}`;
  const res = await fetch(endpointUrl);
  if (!res.ok) {
    const errorText = await res.text();
    console.error(`Result fetch failed: ${res.status} - ${errorText}`);
    throw new Error(`Result not ready: ${res.status}`);
  }
  const blob = await res.blob();
  const resultUrl = URL.createObjectURL(blob);
  return { downloadUrl: resultUrl, openUrl: resultUrl };
}

async function fetchConversionCsv(jobId: string): Promise<string> {
  const endpointUrl = `${API_BASE_URL}${RESULT_PATH}?jobId=${encodeURIComponent(jobId)}&format=csv`;
  const res = await fetch(endpointUrl);
  if (!res.ok) throw new Error("CSV not ready");
  return res.text();
}

function bytesToReadable(size: number): string {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  if (size < 1024 * 1024 * 1024) return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  return `${(size / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}
export default function WordToExcel(): React.ReactElement {
  const [files, setFiles] = useState<FileItem[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [jobId, setJobId] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [statusMessage, setStatusMessage] = useState<string>("");
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [result, setResult] = useState<ConversionResult | null>(null);
  const [excelUploaded, setExcelUploaded] = useState<boolean>(false);
  const [mappingApplied, setMappingApplied] = useState<boolean>(false);
  const [extractFileUploaded, setExtractFileUploaded] = useState<boolean>(false);
  const [mappingFileUploaded, setMappingFileUploaded] = useState<boolean>(false);
  const [showApplyMappingButton, setShowApplyMappingButton] = useState<boolean>(false);
  const { isDarkMode, toggleTheme } = useTheme();

  const inputRef = useRef<HTMLInputElement | null>(null);
  const excelInputRef = useRef<HTMLInputElement | null>(null);
  const extractFileInputRef = useRef<HTMLInputElement | null>(null);
  const mappingFileInputRef = useRef<HTMLInputElement | null>(null);
  const pollingRef = useRef<number | null>(null);

  const hasFiles = files.length > 0 || extractFileUploaded;

  const acceptedExtensions = React.useMemo(() => [
    ".doc",
    ".docx",
    ".rtf",
    ".odt",
  ], []);

  const onFileInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList) return;
    const newFiles: FileItem[] = [];
    for (let i = 0; i < fileList.length; i += 1) {
      const file = fileList.item(i);
      if (!file) continue;
      const isAccepted = acceptedExtensions.some((ext) => file.name.toLowerCase().endsWith(ext));
      if (!isAccepted) continue;
      newFiles.push({
        id: `${file.name}-${file.size}-${file.lastModified}-${i}`,
        file,
        name: (file as any).webkitRelativePath || file.name,
        size: file.size,
        status: "pending",
      });
    }
    setFiles((prev) => [...prev, ...newFiles]);
  }, [acceptedExtensions]);

  const onDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const dt = e.dataTransfer;
    const items = dt.items;
    const collected: File[] = [];
    if (items && items.length > 0) {
      for (let i = 0; i < items.length; i += 1) {
        const item = items[i];
        if (!item) continue;
        const entry = (item as any).webkitGetAsEntry?.();
        if (entry && entry.isDirectory) {
          continue;
        } else {
          const file = item.getAsFile();
          if (file) collected.push(file);
        }
      }
    } else {
      if (dt.files && dt.files.length > 0) {
        for (let i = 0; i < dt.files.length; i += 1) {
          const f = dt.files.item(i);
          if (f) collected.push(f);
        }
      }
    }
    const filtered = collected.filter((f) => acceptedExtensions.some((ext) => f.name.toLowerCase().endsWith(ext)));
    const mapped: FileItem[] = filtered.map((file, idx) => ({
      id: `${file.name}-${file.size}-${file.lastModified}-${idx}`,
      file,
      name: file.name,
      size: file.size,
      status: "pending",
    }));
    setFiles((prev) => [...prev, ...mapped]);
  }, [acceptedExtensions]);

  const removeFile = useCallback((id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  }, []);

  const updateFileStatus = useCallback((id: string, status: FileItem["status"], errorMessage?: string) => {
    setFiles((prev) => prev.map((f) => 
      f.id === id ? { ...f, status, errorMessage } : f
    ));
  }, []);

  const resetAll = useCallback(async () => {
    try {
      // Inform backend to cancel current job and clear any queued files
      const url = jobId ? `${API_BASE_URL}${RESULT_PATH.replace('/result/', '/reset/')}?jobId=${encodeURIComponent(jobId)}` : `${API_BASE_URL}${RESULT_PATH.replace('/result/', '/reset/')}`;
      await fetch(url, { method: "POST" }).catch(() => {});
    } catch {
      // ignore network errors on reset
    } finally {
      setFiles([]);
      setJobId(null);
      setIsUploading(false);
      setIsConverting(false);
      setProgress(0);
      setStatusMessage("");
      setErrorMessage("");
      setResult(null);
      setExcelUploaded(false);
      setMappingApplied(false);
      setExtractFileUploaded(false);
      setMappingFileUploaded(false);
      setShowApplyMappingButton(false);
      if (pollingRef.current) {
        window.clearInterval(pollingRef.current);
        pollingRef.current = null;
      }
    }
  }, [jobId]);

  const handleExtractFileUpload = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
      setErrorMessage("Only Excel files are allowed (.xlsx or .xls). Please select a valid Excel file.");
      return;
    }
    
    try {
      setStatusMessage("Uploading extract Excel file...");
      
      // If no jobId exists, create a new one for direct Excel upload
      let currentJobId = jobId;
      if (!currentJobId) {
        // Create a new job for direct Excel upload
        const response = await fetch(API_ENDPOINTS.UPLOAD_DIRECT_EXCEL, {
          method: 'POST',
          body: (() => {
            const formData = new FormData();
            formData.append('excelFile', file);
            return formData;
          })(),
        });
        
        if (!response.ok) {
          throw new Error('Failed to create job for direct Excel upload');
        }
        
        const data = await response.json();
        currentJobId = data.jobId;
        setJobId(currentJobId);
        setExtractFileUploaded(true);
        setStatusMessage(`${data.entries} entries - Ready for mapping`);
        return;
      } else {
        // Use existing jobId for regular extract Excel upload
        const result = await uploadExtractExcel(file, currentJobId);
        setExtractFileUploaded(true);
        setStatusMessage(`${result.entries} entries`);
      }
    } catch (error) {
      console.error(`Extract Excel upload error:`, error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      setErrorMessage(`Extract Excel upload failed: ${errorMessage}`);
    }
  }, [jobId]);

  const handleMappingFileUpload = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file || !jobId) return;
    
    if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
      setErrorMessage("Only Excel files are allowed (.xlsx or .xls). Please select a valid Excel file.");
      return;
    }
    
    try {
      setStatusMessage("Uploading mapping Excel file...");
      const result = await uploadExcelSheet(file, jobId);
      setMappingFileUploaded(true);
      // Simple logic: show "Applying mapping" button after mapping Excel is uploaded
      setShowApplyMappingButton(true);
      setStatusMessage(`Mapping Excel uploaded: ${result.entries} entries.`);
    } catch (error) {
      console.error(`Mapping Excel upload error:`, error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      setErrorMessage(`Mapping Excel upload failed: ${errorMessage}`);
    }
  }, [jobId]);

  const handleExcelUpload = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file || !jobId) return;
    
    if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
      setErrorMessage("Only Excel files are allowed (.xlsx or .xls). Please select a valid Excel file.");
      return;
    }
    
    try {
      setStatusMessage("Uploading Excel sheet...");
      const result = await uploadExcelSheet(file, jobId);
      setStatusMessage(`Excel uploaded: ${result.entries} entries.`);
      setExcelUploaded(true);
    } catch (error) {
      console.error(`Excel upload error:`, error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      setErrorMessage(`Excel upload failed: ${errorMessage}`);
    }
  }, [jobId]);

  const applyMapping = useCallback(async () => {
    // Check if both files are uploaded
    if (!extractFileUploaded || !mappingFileUploaded) {
      setErrorMessage("Please upload both Extract Excel file and Mapping Excel file before applying mapping.");
      return;
    }
    
    try {
      setErrorMessage("");
      setStatusMessage("Applying Excel mapping... Please wait, this may take a few seconds.");
      setIsConverting(true);
      
      // Call the new apply mapping API
      const formData = new FormData();
      formData.append('jobId', jobId!);
      
      const response = await fetch(API_ENDPOINTS.APPLY_MAPPING, {
        method: "POST",
        body: formData,
      });
      
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.message || errorData.error || "Failed to apply mapping");
      }
      
      const result = await response.json();
      setExcelUploaded(true);
      setMappingApplied(true);
      setShowApplyMappingButton(false);
      setStatusMessage("Mapping applied! Ready for download.");
      
      const updatedResult = await fetchConversionResult(jobId!);
      setResult(updatedResult);
      
      setIsConverting(false);
    } catch (error) {
      setErrorMessage(`Mapping failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setIsConverting(false);
    }
  }, [jobId, extractFileUploaded, mappingFileUploaded]);

  const beginConversion = useCallback(async () => {
    try {
      setErrorMessage("");
      setStatusMessage("Uploading files...");
      setIsUploading(true);
      
      // Set all files to converting status
      files.forEach((file) => {
        updateFileStatus(file.id, "converting");
      });
      
      // Batch upload in chunks
      const batchSize = 200;
      const fileChunks = chunkArray(files.map((f) => f.file), batchSize);
      let createdJobId: string | null = null;
      for (let i = 0; i < fileChunks.length; i += 1) {
        const res = await uploadFolderToBackend(fileChunks[i] || [], createdJobId || undefined);
        createdJobId = res.jobId;
        setStatusMessage(`Uploading batch ${i + 1}/${fileChunks.length}...`);
        setProgress(Math.min(4, 4));
      }
      setIsUploading(false);
      setJobId(createdJobId);

      setStatusMessage("Starting conversion...");
      setIsConverting(true);
      if (!createdJobId) throw new Error("Failed to initialize job");
      await startBackendConversion(createdJobId);

      setStatusMessage("Converting...");
      setProgress(5);

      // Track individual file processing
      let processedFiles = 0;
      const totalFiles = files.length;
      let lastProgress = 0;

      pollingRef.current = window.setInterval(async () => {
        try {
          const p = await pollConversionProgress(createdJobId!);
          if (p.error) throw new Error(p.error);
          
          // Get progress from backend
          const backendProgress = p.progress || 0;
          const currentProgress = Math.min(100, Math.max(0, backendProgress));
          
          // Only update progress if it has actually increased (no reset/flicker)
          if (currentProgress > lastProgress) {
            setProgress(currentProgress);
            lastProgress = currentProgress;
            
            // Update file statuses based on progress
            // Backend now provides progress from 5% to 85% for file processing
            // Map this to individual file completion
            const fileProcessingProgress = Math.max(0, currentProgress - 5); // Remove initial 5%
            const fileProgressRatio = Math.min(1, fileProcessingProgress / 80); // 80% for file processing
            const expectedCompletedFiles = Math.floor(fileProgressRatio * totalFiles);
            
            // Mark files as success if they should be completed
            for (let i = processedFiles; i < expectedCompletedFiles && i < totalFiles; i++) {
              const file = files[i];
              if (file) {
                updateFileStatus(file.id, "success");
              }
            }
            processedFiles = Math.max(processedFiles, expectedCompletedFiles);
          }
          
          if (p.done) {
            // Mark all remaining files as success
            files.forEach((file) => {
              if (file && file.status === "converting") {
                updateFileStatus(file.id, "success");
              }
            });
            
            if (pollingRef.current) {
              window.clearInterval(pollingRef.current);
              pollingRef.current = null;
            }
            setStatusMessage("Finalizing...");
            // const res = await fetchConversionResult(uploadRes.jobId);
            const res = await fetchConversionResult(createdJobId!);
            setResult(res);
            setIsConverting(false);
            setStatusMessage("Conversion complete! Ready for download.");
          }
        } catch (err: any) {
          if (pollingRef.current) {
            window.clearInterval(pollingRef.current);
            pollingRef.current = null;
          }
          setIsConverting(false);
          const errorMessage = err?.message || "An error occurred during conversion.";
          console.error("Conversion polling error:", errorMessage);
          setErrorMessage(errorMessage);
          setStatusMessage("");
        }
      }, 300); // Even more responsive updates
    } catch (err: any) {
      setIsUploading(false);
      setIsConverting(false);
      setStatusMessage("");
      setErrorMessage(err?.message || "Failed to start conversion.");
    }
  }, [files, updateFileStatus]);

  return (
    <div className={`min-h-screen ${isDarkMode ? 'bg-gray-950 text-gray-100' : 'bg-gray-50 text-gray-800'}`}>
      {/* Header */}
      <header className={`sticky top-0 z-10 backdrop-blur shadow-sm ${isDarkMode ? 'bg-gray-900/60 border-gray-800' : 'bg-white/70 border-gray-100'} border-b`}>
        <div className="mx-auto max-w-3xl px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-9 w-9 rounded-xl bg-indigo-600 text-white grid place-items-center font-bold">W</div>
            <div className="leading-tight">
              <p className={`text-2xl font-bold ${isDarkMode ? 'text-gray-100' : 'text-gray-900'}`}>Word → Excel Converter</p>
            </div>
          </div>
          
          <div className="flex items-center gap-3">
            <button
              type="button"
              aria-label="Toggle theme"
              onClick={toggleTheme}
              className={`inline-flex items-center gap-2 rounded-full border px-3 py-1.5 text-sm transition-colors ${isDarkMode ? 'border-gray-700 text-gray-300 hover:bg-gray-800' : 'border-gray-200 text-gray-600 hover:bg-gray-100'}`}
            >
              <span className={`h-2.5 w-2.5 rounded-full ${isDarkMode ? "bg-yellow-400" : "bg-gray-400"}`} />
              {isDarkMode ? "Dark" : "Light"}
            </button>
            <ProfileDropdown isDarkMode={isDarkMode} />
          </div>
        </div>
      </header>

      {/* Main */}
      <main className="mx-auto max-w-3xl p-6">
        <div className="mb-6 md:mb-6">
          <h1 className={`text-2xl md:text-3xl font-bold ${isDarkMode ? 'text-gray-100' : 'text-gray-900'}`}>Convert Word documents to Excel</h1>
          <p className={`mt-2 ${isDarkMode ? 'text-gray-400' : 'text-gray-600'}`}>Upload a folder of Word files and we will convert them into a single Excel file. Simple, fast, and secure.</p>
        </div>

        

        {/* Upload Section */}
        <section>
          <motion.div
            layout
            className={`rounded-xl border-2 border-dashed shadow-md p-8 md:p-10 transition-colors ${isDragging ? (isDarkMode ? "border-indigo-500 bg-indigo-900/20" : "border-indigo-500 bg-indigo-50") : (isDarkMode ? "border-gray-300 bg-gray-900" : "border-gray-300 bg-white")}`}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={(e) => {
              e.preventDefault();
              setIsDragging(false);
            }}
            onDrop={onDrop}
          >
            <div className="flex flex-col items-center text-center gap-4">
              <div className={`h-14 w-14 rounded-2xl grid place-items-center ${isDragging ? (isDarkMode ? "bg-indigo-900/40 text-indigo-600" : "bg-indigo-100 text-indigo-600") : (isDarkMode ? "bg-gray-800 text-gray-300" : "bg-gray-100 text-gray-500")}`}>
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-7 w-7">
                  <path d="M12 3a5 5 0 0 0-5 5v2H6a4 4 0 0 0 0 8h12a4 4 0 0 0 0-8h-1V8a5 5 0 0 0-5-5Zm-1 9V8a1 1 0 1 1 2 0v4h2.5a.75.75 0 0 1 .53 1.28l-3.5 3.5a.75.75 0 0 1-1.06 0l-3.5-3.5A.75.75 0 0 1 8.5 12H11Z" />
                </svg>
              </div>

              <div>
                <p className={`text-base md:text-lg font-medium ${isDarkMode ? 'text-gray-100' : 'text-gray-900'}`}>Drop folder here or click to browse</p>
                <p className={`text-sm ${isDarkMode ? 'text-gray-400' : 'text-gray-500'}`}>Accepted: DOC, DOCX, RTF, ODT</p>
              </div>

              {/* Action row: browse + start + reset */}
              <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                <div className="flex flex-wrap items-center justify-center gap-3">
                  <button
                    type="button"
                    onClick={() => inputRef.current?.click()}
                    className="inline-flex items-center gap-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-6 py-3 shadow focus-visible:outline focus-visible:outline-2 focus-visible:outline-indigo-600"
                  >
                    Browse Folder
                  </button>
                  <button
                    type="button"
                    disabled={!hasFiles || isUploading || isConverting}
                    onClick={showApplyMappingButton ? applyMapping : beginConversion}
                    className="inline-flex justify-center items-center gap-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-6 py-3 shadow disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {isUploading ? "Uploading..." : isConverting ? "Converting..." : showApplyMappingButton ? "Apply Mapping" : "Start Conversion"}
                  </button>
                  <button
                    type="button"
                    onClick={() => extractFileInputRef.current?.click()}
                    className="inline-flex items-center gap-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-medium px-5 py-2.5 shadow"
                  >
                    {jobId ? "Extract Excel" : "Upload Excel"}
                  </button>
                  <button
                    type="button"
                    onClick={() => mappingFileInputRef.current?.click()}
                    className="inline-flex items-center gap-2 rounded-xl bg-green-600 hover:bg-green-700 text-white font-medium px-5 py-2.5 shadow"
                  >
                    Mapping Excel
                  </button>
                  <button
                    type="button"
                    onClick={resetAll}
                    className={`inline-flex items-center gap-2 rounded-xl border px-6 py-3 font-medium ${isDarkMode ? 'border-gray-700 text-gray-300 hover:bg-gray-800' : 'border-gray-300 text-gray-600 hover:bg-gray-100'}`}
                  >
                    Reset
                  </button>
                </div>
              </div>

              <input
                ref={inputRef}
                type="file"
                multiple
                // @ts-expect-error - non-standard but widely supported in Chromium-based browsers
                webkitdirectory="true"
                directory="true"
                className="hidden"
                onChange={onFileInputChange}
              />
              <input
                ref={excelInputRef}
                type="file"
                onChange={handleExcelUpload}
                accept=".xlsx,.xls"
                className="hidden"
              />
              <input
                ref={extractFileInputRef}
                type="file"
                onChange={handleExtractFileUpload}
                accept=".xlsx,.xls"
                className="hidden"
              />
              <input
                ref={mappingFileInputRef}
                type="file"
                onChange={handleMappingFileUpload}
                accept=".xlsx,.xls"
                className="hidden"
              />
            </div>
          </motion.div>
        </section>

        {/* Progress & Download (directly below upload box) */}
        <AnimatePresence>
          {(isUploading || isConverting || progress > 0 || statusMessage || result) && (
            <motion.section
              layout
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -12 }}
              transition={{ duration: 0.2 }}
              className="mt-8"
              aria-live="polite"
            >
              <div className={`rounded-2xl shadow-sm p-6 ${isDarkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'} border`}>
                <div className={`mb-2 text-sm ${isDarkMode ? 'text-gray-400' : 'text-gray-600'}`}>{progress}%</div>
                <div
                  className={`w-full h-3 rounded-full overflow-hidden ${isDarkMode ? 'bg-gray-700' : 'bg-gray-200'}`}
                  role="progressbar"
                  aria-valuemin={0}
                  aria-valuemax={100}
                  aria-valuenow={progress}
                >
                  <motion.div
                    className="h-full bg-indigo-600"
                    initial={{ width: "0%" }}
                    animate={{ width: `${progress}%` }}
                    transition={{ ease: "easeInOut", duration: 0.5 }}
                    style={{ borderRadius: 9999 }}
                  />
                </div>
                <div className="mt-3 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                  <div className={`text-sm ${isDarkMode ? 'text-gray-400' : 'text-gray-500'}`}>
                    {statusMessage || (progress > 0 ? `Progress: ${progress}%` : "Idle")}
                  </div>
                  {result && (
                    <div className="flex items-center gap-3">
                      <button
                        onClick={() => {
                          // Create a temporary link to trigger download
                          const link = document.createElement('a');
                          link.href = result.downloadUrl;
                          link.download = '';
                          document.body.appendChild(link);
                          link.click();
                          document.body.removeChild(link);
                          
                          // Reset after download
                          setTimeout(() => {
                            resetAll();
                          }, 1000); // Wait 1 second after download starts
                        }}
                        className="inline-flex items-center gap-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-5 py-2.5 shadow"
                      >
                        {mappingApplied ? "Download Mapped Excel File" : "Download Excel File"}
                      </button>
                      {result.openUrl && (
                        <a
                          href={result.openUrl}
                          target="_blank"
                          rel="noreferrer"
                          className={`inline-flex items-center gap-2 rounded-xl border px-5 py-2.5 font-medium ${isDarkMode ? 'border-gray-700 text-gray-300 hover:bg-gray-800' : 'border-gray-300 text-gray-700 hover:bg-gray-100'}`}
                        >
                          {mappingApplied ? "Open Mapped Excel in New Tab" : "Open in New Tab"}
                        </a>
                      )}
                    </div>
                  )}
                </div>
              </div>
            </motion.section>
          )}
        </AnimatePresence>

        {/* Uploaded Files List */}
        <AnimatePresence>
          {hasFiles && (
            <motion.section
              layout
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -12 }}
              transition={{ duration: 0.2 }}
              className="mt-8"
            >
              <div className={`rounded-2xl shadow-sm ${isDarkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'} border`}>
                <div className={`p-5 md:p-6 border-b flex items-center justify-between ${isDarkMode ? 'border-gray-700' : 'border-gray-100'}`}>
                  <h2 className={`text-lg font-semibold ${isDarkMode ? 'text-gray-100' : 'text-gray-900'}`}>Files</h2>
                  <span className={`text-sm ${isDarkMode ? 'text-gray-400' : 'text-gray-500'}`}>{files.length} selected</span>
                </div>
                <div className="max-h-96 overflow-y-auto">
                <ul className={`divide-y ${isDarkMode ? 'divide-gray-700' : 'divide-gray-100'}`}>
                  {files.map((item) => (
                    <li key={item.id} className="px-5 md:px-6 py-2 flex items-center gap-4">
                        <div className={`h-9 w-9 rounded-lg grid place-items-center relative ${
                          item.status === "success" 
                            ? (isDarkMode ? "bg-green-900/30 text-green-400" : "bg-green-50 text-green-600")
                            : item.status === "converting"
                            ? (isDarkMode ? "bg-blue-900/30 text-blue-400" : "bg-blue-50 text-blue-600")
                            : item.status === "error"
                            ? (isDarkMode ? "bg-red-900/30 text-red-400" : "bg-red-50 text-red-600")
                            : (isDarkMode ? "bg-indigo-900/30 text-indigo-400" : "bg-indigo-50 text-indigo-600")
                        }`}>
                          {item.status === "success" ? (
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5">
                              <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
                            </svg>
                          ) : item.status === "converting" ? (
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5 animate-spin">
                              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                            </svg>
                          ) : (
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5">
                          <path d="M6 2a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h9.5a2 2 0 0 0 2-2V8.5L13.5 2H6Zm7 1.5L18.5 9H13a.5.5 0 0 1-.5-.5V3.5Z" />
                        </svg>
                          )}
                      </div>
                      <div className="flex-1 min-w-0">
                        <p className={`truncate font-medium ${isDarkMode ? 'text-gray-100' : 'text-gray-900'}`}>{item.name}</p>
                        <p className={`text-sm ${isDarkMode ? 'text-gray-400' : 'text-gray-500'}`}>{bytesToReadable(item.size)}</p>
                      </div>
                      <div className="hidden sm:block">
                          <span className={`text-xs rounded-full px-2 py-1 border ${
                            item.status === "success" 
                              ? (isDarkMode ? "border-green-700 text-green-400 bg-green-900/30" : "border-green-200 text-green-600 bg-green-50")
                              : item.status === "converting"
                              ? (isDarkMode ? "border-blue-700 text-blue-400 bg-blue-900/30" : "border-blue-200 text-blue-600 bg-blue-50")
                              : item.status === "error"
                              ? (isDarkMode ? "border-red-700 text-red-400 bg-red-900/30" : "border-red-200 text-red-600 bg-red-50")
                              : (isDarkMode ? "border-gray-700 text-gray-300" : "border-gray-200 text-gray-600")
                          }`}>
                            {item.status === "success" ? "✓ Success" : item.status}
                        </span>
                      </div>
                      <button
                        type="button"
                        onClick={() => removeFile(item.id)}
                        className={`ml-2 inline-flex items-center justify-center h-8 w-8 rounded-lg text-red-500 hover:text-red-600 transition-colors ${isDarkMode ? 'hover:bg-red-900/20' : 'hover:bg-red-50'}`}
                        aria-label={`Remove ${item.name}`}
                      >
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5">
                          <path d="M9 3a1 1 0 0 0-1 1v1H5.5a1 1 0 1 0 0 2H6v12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V7h.5a1 1 0 1 0 0-2H16V4a1 1 0 0 0-1-1H9Zm2 4a1 1 0 0 0-1 1v9a1 1 0 1 0 2 0V8a1 1 0 0 0-1-1Zm4 0a1 1 0 0 0-1 1v9a1 1 0 1 0 2 0V8a1 1 0 0 0-1-1Z" />
                        </svg>
                      </button>
                    </li>
                  ))}
                </ul>
                </div>
              </div>
            </motion.section>
          )}
        </AnimatePresence>

        

        {/* Error */}
        <AnimatePresence>
          {errorMessage && (
            <motion.section
              layout
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -12 }}
              transition={{ duration: 0.2 }}
              className="mt-8"
              aria-live="assertive"
            >
              <div className={`rounded-2xl border p-5 flex items-start gap-3 ${isDarkMode ? 'border-red-900 bg-red-950 text-red-300' : 'border-red-200 bg-red-50 text-red-800'}`}>
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5 mt-0.5">
                  <path d="M12 2a10 10 0 1 0 10 10A10.011 10.011 0 0 0 12 2Zm1 15h-2v-2h2Zm0-4h-2V7h2Z" />
                </svg>
                <div>
                  <p className="font-semibold">Conversion failed</p>
                  <p className="text-sm">{errorMessage}</p>
                </div>
              </div>
            </motion.section>
          )}
        </AnimatePresence>

        {/* Result card removed in favor of inline download above */}
      </main>
    </div>
  );
}