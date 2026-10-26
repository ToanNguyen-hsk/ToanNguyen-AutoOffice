import React, { useState } from 'react';

function App() {
  const [coverFile, setCoverFile] = useState(null);
  const [carFiles, setCarFiles] = useState([]);
  const [carStatuses, setCarStatuses] = useState({});
  const [processedResults, setProcessedResults] = useState([]);
  const [isExporting, setIsExporting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [currentProcessingFile, setCurrentProcessingFile] = useState('');
  const [processMode, setProcessMode] = useState('auto_bump');
  const [qtiNo, setQtiNo] = useState('QTI-PECC1-PMB2-26-10-111');

  const handleCoverChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      setCoverFile(e.target.files[0]);
    }
  };

  const handleCarFilesChange = (e) => {
    const files = Array.from(e.target.files);
    setCarFiles(files);
    const initialStatuses = {};
    files.forEach(f => { initialStatuses[f.name] = 'APP'; });
    setCarStatuses(initialStatuses);
  };

  const handleStatusChange = (fileName, status) => {
    setCarStatuses(prev => ({ ...prev, [fileName]: status }));
  };

  const handleExportZipFull = async () => {
    if (carFiles.length === 0) return alert("Vui lòng nạp danh sách file CaRs ở Cột 2 trước!");
    setIsExporting(true);
    setProgress(5);
    setCurrentProcessingFile('Đang bóc tách danh mục tài liệu & Ngày đệ trình từ Thư Cover EPC...');

    const totalFiles = carFiles.length;
    for (let i = 0; i < totalFiles; i++) {
      setCurrentProcessingFile(
        processMode === 'auto_bump' 
          ? `Đang định dạng & nhân bản Sheet: ${carFiles[i].name}` 
          : `Đang trích xuất tự động Status từ file: ${carFiles[i].name}`
      );
      setProgress(10 + Math.round(((i + 1) / totalFiles) * 75));
      await new Promise(resolve => setTimeout(resolve, 30));
    }

    setCurrentProcessingFile('Đang điền dữ liệu vào Form Word gốc & Đóng gói ZIP...');
    setProgress(90);

    const formData = new FormData();
    if (coverFile) formData.append('cover_file', coverFile);
    carFiles.forEach(file => formData.append('files', file));
    formData.append('mode', processMode);
    formData.append('qti_no', qtiNo.trim());

    if (processMode === 'auto_bump') {
      formData.append('metadata', JSON.stringify(carStatuses));
    } else {
      formData.append('metadata', JSON.stringify({})); 
    }

    try {
      // Đã trỏ tới Backend Render của Toàn Nguyễn và sửa lỗi cú pháp fetch
      const response = await fetch('https://toannguyen-autooffice.onrender.com/api/export-zip-full', {
        method: 'POST',
        body: formData
      });
      
      if (!response.ok) {
        const errText = await response.text();
        throw new Error(errText);
      }
      
      const resultData = await response.json();
      
      if (resultData.processed_items) {
        setProcessedResults(resultData.processed_items);
      }

      if (resultData.zip_base64) {
        const linkSource = `data:application/zip;base64,${resultData.zip_base64}`;
        const downloadLink = document.createElement("a");
        downloadLink.href = linkSource;
        
        // Tên file xuất ra chỉ chứa đúng Mã QTI
        const safeQtiName = qtiNo.replace(/[\\/*?:"<>|]/g, '-').trim();
        downloadLink.download = `${safeQtiName}.zip`;
        
        downloadLink.click();
        downloadLink.remove();
      }
      
      setProgress(100);
      setCurrentProcessingFile('✅ Đã xuất trọn bộ hồ sơ thành công!');
    } catch (err) {
      alert("❌ LỖI HỆ THỐNG:\n\n" + err.message);
      setProgress(0);
      setCurrentProcessingFile('');
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div style={{ fontFamily: 'Segoe UI, Arial, sans-serif', backgroundColor: '#f1f5f9', minHeight: '100vh', padding: '20px' }}>
      <header style={{ backgroundColor: '#0f172a', color: '#fff', padding: '16px 24px', borderRadius: '8px', marginBottom: '20px' }}>
        <h2 style={{ margin: 0, fontSize: '20px' }}>⚡ QT1 - PECC1 AUTOMATED CaR REVISION SYSTEM</h2>
      </header>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.3fr 1.2fr', gap: '20px' }}>
        
        {/* CỘT 1 */}
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
          <h3 style={{ borderBottom: '2px solid #0284c7', paddingBottom: '8px', color: '#0369a1', marginTop: 0 }}>📄 CỘT 1: THƯ COVER SHEET EPC</h3>
          <p style={{ fontSize: '13px', color: '#64748b' }}>Nạp Thư Cover đệ trình từ EPC dạng <b>.pdf</b> hoặc <b>.xlsx</b>:</p>
          <input type="file" accept=".pdf,.xlsx,.xls" onChange={handleCoverChange} style={{ width: '100%', marginBottom: '10px' }} />
          {coverFile && (
            <div style={{ backgroundColor: '#f0fdf4', border: '1px solid #86efac', padding: '8px', borderRadius: '4px', fontSize: '12px', color: '#166534' }}>✓ Đã nạp: <b>{coverFile.name}</b></div>
          )}
        </div>

        {/* CỘT 2 */}
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
          <h3 style={{ borderBottom: '2px solid #0d9488', paddingBottom: '8px', color: '#0f766e', marginTop: 0 }}>⚙️ CỘT 2: CARs & TÙY CHỌN KẾT QUẢ</h3>
          
          <div style={{ marginBottom: '15px', padding: '12px', backgroundColor: '#f8fafc', borderRadius: '6px', border: '1px solid #cbd5e1' }}>
            <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: 'bold', color: '#0f766e' }}>CHẾ ĐỘ XỬ LÝ EXCEL CaRs:</label>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '14px' }}>
              <label style={{ cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
                <input type="radio" name="mode" value="auto_bump" checked={processMode === 'auto_bump'} onChange={() => setProcessMode('auto_bump')} style={{ marginRight: '8px' }} />
                🔄 Tự động tạo Round mới (R+1) & Cập nhật Form
              </label>
              <label style={{ cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
                <input type="radio" name="mode" value="pre_prepared" checked={processMode === 'pre_prepared'} onChange={() => setProcessMode('pre_prepared')} style={{ marginRight: '8px' }} />
                📁 Đã có sẵn CaR (Chỉ bóc tách tạo Thư Cover)
              </label>
            </div>
          </div>

          <input type="file" multiple accept=".xlsx,.xls" onChange={handleCarFilesChange} style={{ width: '100%', marginBottom: '15px' }} />
          
          {carFiles.length > 0 && (
            <div style={{ maxHeight: '420px', overflowY: 'auto', marginBottom: '15px' }}>
              {carFiles.map((file, idx) => (
                <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#f8fafc', padding: '8px', marginBottom: '6px', borderRadius: '4px', border: '1px solid #e2e8f0' }}>
                  <span style={{ fontSize: '12px', width: '70%', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={file.name}>{file.name}</span>
                  
                  {processMode === 'auto_bump' ? (
                    <select 
                      value={carStatuses[file.name] || 'APP'} 
                      onChange={(e) => handleStatusChange(file.name, e.target.value)} 
                      style={{ padding: '4px', borderRadius: '4px', fontWeight: 'bold', fontSize: '12px', border: '1px solid #cbd5e1' }}
                    >
                      <option value="APP">🟢 APP</option>
                      <option value="RFC">🔴 RFC</option>
                      <option value="AWC">🟡 AWC</option>
                      <option value="RJ">⬛ RJ</option>
                    </select>
                  ) : (
                    <span style={{ fontSize: '11px', color: '#0369a1', fontWeight: 'bold', padding: '4px 8px', backgroundColor: '#e0f2fe', borderRadius: '4px', border: '1px solid #bae6fd' }}>
                      🔍 Auto Extract
                    </span>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>

        {/* CỘT 3 */}
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
          <h3 style={{ borderBottom: '2px solid #16a34a', paddingBottom: '8px', color: '#15803d', marginTop: 0 }}>📋 CỘT 3: KẾT QUẢ & XUẤT HỒ SƠ</h3>
          
          <div style={{ marginBottom: '15px' }}>
            <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', color: '#0f766e', marginBottom: '6px' }}>
              Số Ref/Tên file xuất (Tùy chỉnh):
            </label>
            <input 
              type="text" 
              value={qtiNo} 
              onChange={(e) => setQtiNo(e.target.value)} 
              style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '14px', fontWeight: 'bold', color: '#0f172a', boxSizing: 'border-box' }} 
            />
          </div>

          <button onClick={handleExportZipFull} disabled={isExporting || carFiles.length === 0} style={{ width: '100%', backgroundColor: isExporting || carFiles.length === 0 ? '#94a3b8' : '#16a34a', color: '#fff', border: 'none', padding: '13px', borderRadius: '6px', fontWeight: 'bold', cursor: isExporting || carFiles.length === 0 ? 'not-allowed' : 'pointer', marginBottom: '15px' }}>
            {isExporting ? `⏳ ĐANG XỬ LÝ (${progress}%)...` : '⚡ XUẤT ZIP BỘ HỒ SƠ D OFFICE'}
          </button>
          
          {(isExporting || progress > 0) && (
            <div style={{ backgroundColor: '#f8fafc', border: '1px solid #cbd5e1', padding: '12px', borderRadius: '6px', marginBottom: '15px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', fontWeight: 'bold', marginBottom: '6px' }}>
                <span>Tiến độ xử lý:</span><span style={{ color: '#2563eb' }}>{progress}%</span>
              </div>
              <div style={{ width: '100%', backgroundColor: '#e2e8f0', borderRadius: '4px', height: '10px', overflow: 'hidden' }}>
                <div style={{ width: `${progress}%`, backgroundColor: '#16a34a', height: '100%', transition: 'width 0.2s ease-in-out' }}></div>
              </div>
              <p style={{ fontSize: '11px', color: '#475569', margin: '8px 0