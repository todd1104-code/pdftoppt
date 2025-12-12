import React, { useState, useRef, useEffect } from 'react';
import PptxGenJS from 'pptxgenjs';
import * as pdfjsLib from 'pdfjs-dist';
// Vite handles this import by giving us the URL to the worker file
// import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';

// Set up the worker source
// pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

// 使用內建 SVG 圖示取代外部套件，確保複製貼上即用
const Icons = {
  Upload: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-blue-500">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" />
    </svg>
  ),
  FileText: ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" /><polyline points="14 2 14 8 20 8" />
    </svg>
  ),
  CheckCircle: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-green-500">
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" /><polyline points="22 4 12 14.01 9 11.01" />
    </svg>
  ),
  AlertCircle: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-red-500">
      <circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" />
    </svg>
  ),
  Loader2: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="animate-spin">
      <path d="M21 12a9 9 0 1 1-6.219-8.56" />
    </svg>
  ),
  RefreshCw: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8" /><path d="M21 3v5h-5" /><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16" /><path d="M8 16H3v5" />
    </svg>
  )
};

const PDFToPPTX = () => {
  const [file, setFile] = useState(null);
  const [isConverting, setIsConverting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState('idle');
  const [statusMessage, setStatusMessage] = useState('');
  const canvasRef = useRef(null);
  const libsLoaded = useRef(false);

  // === 狀態與邏輯 ===
  // 移除自動載入樣式與 CDN 函式庫的 useEffect，因為已改為本地安裝與 import


  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile && selectedFile.type === 'application/pdf') {
      setFile(selectedFile);
      setStatus('idle');
      setProgress(0);
      setStatusMessage('');
    } else {
      alert('請上傳 PDF 檔案');
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile && droppedFile.type === 'application/pdf') {
      setFile(droppedFile);
      setStatus('idle');
      setProgress(0);
      setStatusMessage('');
    } else {
      alert('請拖曳 PDF 檔案');
    }
  };

  const convertToPPTX = async () => {
    if (!file) return;

    // 直接使用 import 的核心，無須檢查 window 對象
    if (!pdfjsLib || !PptxGenJS) {
      setStatus('error');
      setStatusMessage('轉換核心載入失敗，請重新整理頁面。');
      return;
    }

    setIsConverting(true);
    setStatus('converting');
    setProgress(0);
    setStatusMessage('正在讀取 PDF...');

    try {
      const arrayBuffer = await file.arrayBuffer();
      // 使用 imported pdfjsLib
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const totalPages = pdf.numPages;


      // 使用 imported PptxGenJS
      const pres = new PptxGenJS();
      pres.layout = 'LAYOUT_16x9';

      for (let i = 1; i <= totalPages; i++) {
        setStatusMessage(`正在轉換第 ${i} / ${totalPages} 頁...`);

        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 2.0 }); // 2.0 解析度確保清晰
        const canvas = canvasRef.current;
        const context = canvas.getContext('2d');

        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({
          canvasContext: context,
          viewport: viewport
        }).promise;

        const imgData = canvas.toDataURL('image/jpeg', 0.8);

        const slide = pres.addSlide();
        slide.addImage({
          data: imgData,
          x: 0,
          y: 0,
          w: '100%',
          h: '100%',
          sizing: { type: 'contain', w: '100%', h: '100%' }
        });

        setProgress(Math.round((i / totalPages) * 100));
      }

      setStatusMessage('正在建立 PPTX 檔案...');

      const fileName = file.name.replace('.pdf', '');
      await pres.writeFile({ fileName: `${fileName}.pptx` });

      setStatus('completed');
      setStatusMessage('轉換完成！下載已開始。');

    } catch (error) {
      console.error("Conversion Error:", error);
      setStatus('error');
      setStatusMessage(`錯誤: ${error.message || '轉換失敗，請確認 PDF 未加密且未損毀。'}`);
    } finally {
      setIsConverting(false);
    }
  };

  const reset = () => {
    setFile(null);
    setStatus('idle');
    setProgress(0);
    setStatusMessage('');
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 flex flex-col items-center justify-center p-4 font-sans text-slate-800">

      <div className="max-w-xl w-full bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-200">

        {/* Header */}
        <div className="bg-blue-600 p-6 text-white text-center">
          <h1 className="text-2xl font-bold flex items-center justify-center gap-2">
            <Icons.FileText className="w-8 h-8" />
            PDF 轉 PPTX 轉換器
          </h1>
          <p className="text-blue-100 mt-2 text-sm opacity-90">
            安全、快速、無須上傳伺服器
          </p>
        </div>

        {/* Main Content */}
        <div className="p-8">

          <canvas ref={canvasRef} style={{ display: 'none' }} />

          {!file && (
            <div
              className="border-3 border-dashed border-slate-300 rounded-xl p-10 flex flex-col items-center justify-center text-center cursor-pointer hover:bg-slate-50 hover:border-blue-400 transition-all duration-300 group"
              onDragOver={handleDragOver}
              onDrop={handleDrop}
              onClick={() => document.getElementById('fileInput').click()}
            >
              <input
                type="file"
                id="fileInput"
                accept=".pdf"
                className="hidden"
                onChange={handleFileChange}
              />
              <div className="bg-blue-50 p-4 rounded-full mb-4 group-hover:bg-blue-100 transition-colors">
                <Icons.Upload />
              </div>
              <p className="text-lg font-medium text-slate-700 mb-1">點擊或拖曳 PDF 檔案至此</p>
              <p className="text-sm text-slate-400">支援所有標準 PDF 文件</p>
            </div>
          )}

          {file && status !== 'completed' && (
            <div className="flex flex-col items-center">
              <div className="w-full bg-slate-50 border border-slate-200 rounded-lg p-4 mb-6 flex items-center gap-4">
                <div className="bg-red-100 p-3 rounded-lg">
                  <Icons.FileText className="w-6 h-6 text-red-500" />
                </div>
                <div className="flex-1 min-w-0">
                  <p className="font-semibold text-slate-800 truncate">{file.name}</p>
                  <p className="text-xs text-slate-500">{(file.size / 1024 / 1024).toFixed(2)} MB</p>
                </div>
                {status === 'idle' && (
                  <button onClick={reset} className="text-slate-400 hover:text-red-500">
                    ✕
                  </button>
                )}
              </div>

              {status === 'idle' && (
                <div className="w-full">
                  <button
                    onClick={convertToPPTX}
                    className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-6 rounded-lg shadow-md hover:shadow-lg transition-all flex items-center justify-center gap-2 active:scale-95"
                  >
                    <Icons.RefreshCw />
                    開始轉換
                  </button>
                  <p className="text-xs text-center text-slate-400 mt-3">
                    注意：轉換後的 PPTX 將以圖片形式呈現，以確保格式不變。
                  </p>
                </div>
              )}

              {status === 'converting' && (
                <div className="w-full text-center">
                  <div className="w-full bg-slate-200 rounded-full h-2.5 mb-4 overflow-hidden">
                    <div
                      className="bg-blue-600 h-2.5 rounded-full transition-all duration-300 ease-out"
                      style={{ width: `${progress}%` }}
                    ></div>
                  </div>
                  <p className="text-blue-600 font-medium flex items-center justify-center gap-2">
                    <Icons.Loader2 />
                    {statusMessage}
                  </p>
                </div>
              )}

              {status === 'error' && (
                <div className="text-red-500 flex items-center gap-2 bg-red-50 p-3 rounded-lg w-full justify-center">
                  <Icons.AlertCircle />
                  {statusMessage}
                </div>
              )}
            </div>
          )}

          {status === 'completed' && (
            <div className="text-center py-4 animate-in fade-in zoom-in duration-300">
              <div className="w-20 h-20 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <Icons.CheckCircle />
              </div>
              <h2 className="text-2xl font-bold text-slate-800 mb-2">轉換成功！</h2>
              <p className="text-slate-500 mb-6">您的檔案已自動下載。</p>

              <button
                onClick={reset}
                className="text-blue-600 hover:text-blue-800 font-medium hover:underline flex items-center justify-center gap-1 mx-auto"
              >
                轉換另一個檔案
              </button>
            </div>
          )}

        </div>

        {/* Footer */}
        <div className="bg-slate-50 p-4 text-center border-t border-slate-200">
          <p className="text-xs text-slate-400">
            Client-side processing powered by PDF.js & PptxGenJS
          </p>
        </div>
      </div>
    </div>
  );
};

export default PDFToPPTX;