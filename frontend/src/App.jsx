import React, { useState, useEffect, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

const COLUMN_ALIASES = {
  'DAFTAR PERTELAAN ARSIP': 'Kode Kopel',
  'daftar pertelaan arsip': 'Kode Kopel',
  'DAFTAR PERTELAAN ARSIP CENTRALFILE KANTOR PUSAT': 'Kode Kopel',
  'daftar pertelaan arsip centralfile kantor pusat': 'Kode Kopel'
};

const getColumnLabel = (key) => {
  if (!key) return '';
  const trimmed = key.trim();
  
  // Rule-based check for "pertelaan arsip" variations (case-insensitive)
  if (trimmed.toLowerCase().includes('pertelaan arsip')) {
    return 'Kode Kopel';
  }
  
  if (COLUMN_ALIASES[trimmed]) {
    return COLUMN_ALIASES[trimmed];
  }
  return trimmed;
};

function App() {
  // Navigation Tabs: 'search' | 'upload' | 'files' | 'logs' | 'dashboard' | 'bookmarks'
  const [activeTab, setActiveTab] = useState('search');
  
  // Dark/Light Theme
  const [isDarkMode, setIsDarkMode] = useState(() => {
    const saved = localStorage.getItem('app_theme');
    return saved ? saved === 'dark' : true;
  });

  // Toast notifications
  const [toasts, setToasts] = useState([]);

  // Dashboard stats
  const [dashboardStats, setDashboardStats] = useState(null);
  const [loadingDashboard, setLoadingDashboard] = useState(false);
  
  // Database status
  const [dbConnected, setDbConnected] = useState(false);
  const [stats, setStats] = useState({ totalFiles: 0, totalRows: 0 });
  
  // Search state
  const [searchQuery, setSearchQuery] = useState('');
  const [searchResults, setSearchResults] = useState([]);
  const [searchUIState, setSearchUIState] = useState('loading'); // 'welcome' | 'preview' | 'loading' | 'results' | 'empty'
  const [selectedFileId, setSelectedFileId] = useState(null); // Active file tab in results
  
  // Advanced filters state
  const [filterSheet, setFilterSheet] = useState('');
  const [filterUnit, setFilterUnit] = useState('');
  const [showAdvancedFilters, setShowAdvancedFilters] = useState(false);

  // Search History state
  const [searchHistory, setSearchHistory] = useState(() => {
    try {
      const saved = localStorage.getItem('spreadsheet_search_history');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  // Bookmarks state
  const [bookmarksList, setBookmarksList] = useState([]);
  const [loadingBookmarks, setLoadingBookmarks] = useState(false);
  
  // Bulk Delete selection state
  const [selectedFileIds, setSelectedFileIds] = useState([]);

  // Column visibility filter state (persisted in localStorage)
  const [visibleColumnsOrder, setVisibleColumnsOrder] = useState(() => {
    try {
      const saved = localStorage.getItem('spreadsheet_visible_columns_order');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });
  const [showColumnFilter, setShowColumnFilter] = useState(false);
  const [columnSearchQuery, setColumnSearchQuery] = useState('');

  // Bookmark Column visibility filter state
  const [bookmarkVisibleColumns, setBookmarkVisibleColumns] = useState(() => {
    try {
      const saved = localStorage.getItem('spreadsheet_bookmark_visible_columns');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });
  const [showBookmarkColumnFilter, setShowBookmarkColumnFilter] = useState(false);
  const [bookmarkColumnSearchQuery, setBookmarkColumnSearchQuery] = useState('');
  
  // Files state
  const [filesList, setFilesList] = useState([]);
  const [loadingFiles, setLoadingFiles] = useState(false);
  
  // Logs state
  const [logs, setLogs] = useState([]);
  const [loadingLogs, setLoadingLogs] = useState(false);
  const [logsError, setLogsError] = useState(null);
  const [logSearch, setLogSearch] = useState('');
  const [logPage, setLogPage] = useState(1);
  const [logLimit, setLogLimit] = useState(50);
  const [logTotal, setLogTotal] = useState(0);
  const [logTotalPages, setLogTotalPages] = useState(1);
  
  // Upload state
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [uploadStatusMessage, setUploadStatusMessage] = useState('');
  const [uploadSuccess, setUploadSuccess] = useState(null);
  const [uploadError, setUploadError] = useState(null);
  const [dragActive, setDragActive] = useState(false);
  const [copySuccess, setCopySuccess] = useState(false);
  
  const fileInputRef = useRef(null);
  const searchInputRef = useRef(null);
  const searchAbortControllerRef = useRef(null);

  // Initial checks and loads
  useEffect(() => {
    checkHealth();
    fetchFiles();
  }, []);

  // Sync dark/light mode to body class and localStorage
  useEffect(() => {
    if (isDarkMode) {
      document.body.classList.remove('light-mode');
    } else {
      document.body.classList.add('light-mode');
    }
    localStorage.setItem('app_theme', isDarkMode ? 'dark' : 'light');
  }, [isDarkMode]);

  // Sync URL search param on mount (share link feature)
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const q = params.get('q');
    if (q) {
      setSearchQuery(q);
      setActiveTab('search');
    }
  }, []);

  // Keyboard shortcut: focus search input on '/' key
  useEffect(() => {
    const handleKeyDown = (e) => {
      if (e.key === '/' && document.activeElement !== searchInputRef.current) {
        e.preventDefault();
        setActiveTab('search');
        setTimeout(() => searchInputRef.current?.focus(), 50);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, []);

  // Close filter dropdown and reset search input when selected file changes
  useEffect(() => {
    setShowColumnFilter(false);
    setColumnSearchQuery('');
  }, [selectedFileId]);

  // Save column visibility order changes to localStorage
  useEffect(() => {
    localStorage.setItem('spreadsheet_visible_columns_order', JSON.stringify(visibleColumnsOrder));
  }, [visibleColumnsOrder]);

  // Save bookmark column visibility changes to localStorage
  useEffect(() => {
    localStorage.setItem('spreadsheet_bookmark_visible_columns', JSON.stringify(bookmarkVisibleColumns));
  }, [bookmarkVisibleColumns]);

  // Fetch logs when logs tab is selected, or when search/page/limit changes
  useEffect(() => {
    if (activeTab === 'logs') {
      fetchLogs(logSearch, logPage, logLimit);
    }
  }, [activeTab, logPage, logLimit]);

  // Debounce log search
  useEffect(() => {
    if (activeTab !== 'logs') return;
    const timer = setTimeout(() => {
      setLogPage(1);
      fetchLogs(logSearch, 1, logLimit);
    }, 300);
    return () => clearTimeout(timer);
  }, [logSearch]);

  // Initial preview trigger when files list is first loaded
  useEffect(() => {
    if (searchQuery.trim() === '' && filesList.length > 0 && !selectedFileId) {
      fetchFilePreview(filesList[0].id);
    }
  }, [filesList]);

  // Fetch bookmarks when bookmarks tab is active
  useEffect(() => {
    if (activeTab === 'bookmarks') {
      fetchBookmarks();
    }
  }, [activeTab]);

  // Debounced search trigger / Auto preview trigger (re-runs when searchQuery, filterSheet, or filterUnit changes)
  useEffect(() => {
    // Sync URL with search query
    const url = new URL(window.location);
    if (searchQuery.trim()) {
      url.searchParams.set('q', searchQuery.trim());
    } else {
      url.searchParams.delete('q');
    }
    window.history.replaceState({}, '', url);

    if (searchQuery.trim() === '') {
      if (filesList.length > 0) {
        const activeId = selectedFileId && filesList.some(f => f.id === selectedFileId) 
          ? selectedFileId 
          : filesList[0].id;
        fetchFilePreview(activeId, filterSheet, filterUnit);
      } else {
        setSearchResults([]);
        setSelectedFileId(null);
        setSearchUIState('welcome');
      }
      return;
    }

    setSearchUIState('loading');

    const timer = setTimeout(() => {
      handleSearch();
    }, 600);

    return () => clearTimeout(timer);
  }, [searchQuery, filterSheet, filterUnit]);

  // Check backend & DB status
  const checkHealth = async () => {
    try {
      const res = await fetch('/api/health');
      const data = await res.json();
      setDbConnected(data.status === 'ok' && data.database === 'connected');
    } catch (err) {
      setDbConnected(false);
    }
  };

  // Toast notification system
  const showToast = useCallback((title, message, type = 'info') => {
    const id = Date.now();
    setToasts(prev => [...prev, { id, title, message, type }]);
    setTimeout(() => {
      setToasts(prev => prev.map(t => t.id === id ? { ...t, exiting: true } : t));
      setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), 350);
    }, 5000);
  }, []);

  const dismissToast = (id) => {
    setToasts(prev => prev.map(t => t.id === id ? { ...t, exiting: true } : t));
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), 350);
  };

  // Share link
  const handleShareLink = () => {
    const url = new URL(window.location);
    if (searchQuery.trim()) url.searchParams.set('q', searchQuery.trim());
    navigator.clipboard.writeText(url.toString()).then(() => {
      showToast('🔗 Link Disalin!', `Link pencarian "${searchQuery}" berhasil disalin ke clipboard.`, 'info');
    }).catch(() => {
      showToast('⚠️ Gagal', 'Tidak bisa menyalin link.', 'error');
    });
  };

  // Export to Excel
  const handleExportExcel = (orderedDataHeaders, showSheet, showBaris) => {
    if (!selectedFileData) return;
    const headers = [
      ...(showSheet ? ['Sheet'] : []),
      ...(showBaris ? ['Baris'] : []),
      ...orderedDataHeaders.map(h => getColumnLabel(h))
    ];
    const rows = selectedFileData.matches.map(match => [
      ...(showSheet ? [match.sheetName] : []),
      ...(showBaris ? [match.rowNumber] : []),
      ...orderedDataHeaders.map(h => match.rowData[h] ?? '')
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Hasil Pencarian');
    const filename = `Hasil-${searchQuery.trim() || 'Pratinjau'}-${new Date().toLocaleDateString('id-ID').replace(/\//g, '-')}.xlsx`;
    XLSX.writeFile(wb, filename);
    showToast('📥 Berhasil Diekspor!', `Data disimpan ke "${filename}"`, 'success');
  };

  // Fetch dashboard statistics
  const fetchDashboard = async () => {
    setLoadingDashboard(true);
    try {
      const res = await fetch('/api/stats');
      if (res.ok) {
        const data = await res.json();
        setDashboardStats(data);
      }
    } catch (err) {
      console.error('Gagal memuat dashboard:', err);
    } finally {
      setLoadingDashboard(false);
    }
  };

  // Get list of uploaded files
  const fetchFiles = async () => {
    setLoadingFiles(true);
    try {
      const res = await fetch('/api/files');
      if (res.ok) {
        const data = await res.json();
        setFilesList(data.files);
        
        // Calculate statistics
        const totalRows = data.files.reduce((sum, f) => sum + f.row_count, 0);
        setStats({ totalFiles: data.files.length, totalRows });

        if (data.files.length === 0) {
          setSearchUIState('welcome');
        }
      }
    } catch (err) {
      console.error('Error fetching files:', err);
    } finally {
      setLoadingFiles(false);
    }
  };

  // Get paginated + searchable activity logs
  const fetchLogs = async (q = logSearch, page = logPage, limit = logLimit) => {
    setLoadingLogs(true);
    setLogsError(null);
    try {
      const params = new URLSearchParams({ page, limit });
      if (q) params.append('q', q);
      const res = await fetch(`/api/logs?${params}`);
      if (res.ok) {
        const data = await res.json();
        setLogs(data.logs || []);
        setLogTotal(data.total || 0);
        setLogTotalPages(data.totalPages || 1);
      } else {
        throw new Error('Gagal mengambil log');
      }
    } catch (err) {
      console.error(err);
      setLogsError(err.message);
    } finally {
      setLoadingLogs(false);
    }
  };

  // Search History helpers
  const addToHistory = (query) => {
    if (!query || query.trim() === '') return;
    const cleanQuery = query.trim();
    setSearchHistory(prev => {
      const filtered = prev.filter(q => q !== cleanQuery);
      const next = [cleanQuery, ...filtered].slice(0, 8);
      localStorage.setItem('spreadsheet_search_history', JSON.stringify(next));
      return next;
    });
  };

  const removeFromHistory = (e, queryToRemove) => {
    e.stopPropagation();
    setSearchHistory(prev => {
      const next = prev.filter(q => q !== queryToRemove);
      localStorage.setItem('spreadsheet_search_history', JSON.stringify(next));
      return next;
    });
  };

  const clearHistory = () => {
    setSearchHistory([]);
    localStorage.removeItem('spreadsheet_search_history');
  };

  // Fetch and display a preview of a file (first 200 rows)
  const fetchFilePreview = async (fileId, sheet = filterSheet, unit = filterUnit) => {
    setSearchUIState('loading');
    setSearchResults([]);
    try {
      const params = new URLSearchParams({ fileId });
      if (sheet) params.append('sheet', sheet);
      if (unit) params.append('unit', unit);

      const res = await fetch(`/api/search?${params}`);
      if (res.ok) {
        const data = await res.json();
        setSearchResults(data.results);
        if (data.results.length > 0) {
          setSelectedFileId(fileId);
          setSearchUIState('preview');
        } else {
          setSearchUIState('empty');
        }
      } else {
        setSearchUIState('empty');
      }
    } catch (err) {
      console.error('Error fetching file preview:', err);
      setSearchUIState('empty');
    }
  };

  // Execute Search
  const handleSearch = async () => {
    if (searchQuery.trim() === '') return;

    // Abort previous search request if any
    if (searchAbortControllerRef.current) {
      searchAbortControllerRef.current.abort();
    }

    const controller = new AbortController();
    searchAbortControllerRef.current = controller;

    setSearchUIState('loading');
    setSearchResults([]);
    try {
      const params = new URLSearchParams({ q: searchQuery });
      if (filterSheet) params.append('sheet', filterSheet);
      if (filterUnit) params.append('unit', filterUnit);

      const res = await fetch(`/api/search?${params}`, {
        signal: controller.signal
      });
      if (res.ok) {
        const data = await res.json();
        setSearchResults(data.results);
        
        // Save to search history
        addToHistory(searchQuery);

        // Auto select the first file in the search results to display it instantly
        if (data.results.length > 0) {
          setSelectedFileId(data.results[0].fileId);
          setSearchUIState('results');
        } else {
          setSelectedFileId(null);
          setSearchUIState('empty');
        }
      } else {
        setSearchUIState('empty');
      }
    } catch (err) {
      if (err.name === 'AbortError') {
        // Ignore aborted requests
        return;
      }
      console.error('Error searching:', err);
      setSearchUIState('empty');
    }
  };

  // Handle file tab click
  const handleFileTabClick = (fileId) => {
    if (searchQuery.trim() === '') {
      fetchFilePreview(fileId, filterSheet, filterUnit);
    } else {
      setSelectedFileId(fileId);
    }
  };

  // Helper to check if a column is visible
  const isColumnVisible = (colName) => {
    if (visibleColumnsOrder.length === 0) return true;        // default: all visible
    if (visibleColumnsOrder.includes('__all_hidden__')) return false; // sentinel: all hidden
    return visibleColumnsOrder.includes(colName);
  };

  // Toggle column visibility with chronological sorting
  const toggleColumnVisibility = (colName) => {
    setVisibleColumnsOrder((prev) => {
      // Coming from "all hidden" sentinel: start fresh with just the clicked column
      if (prev.includes('__all_hidden__')) {
        return [colName];
      }

      // Coming from default "all visible" state: initialize with all columns
      let currentOrder = prev.length === 0
        ? ['Sheet', 'Baris', ...activeHeaders]
        : [...prev];
      
      if (currentOrder.includes(colName)) {
        // Uncheck: remove column from list
        return currentOrder.filter(c => c !== colName);
      } else {
        // Check: append column to the end of list (chronological order)
        return [...currentOrder, colName];
      }
    });
  };

  // Select all columns to be visible in default activeHeaders order
  const selectAllColumns = (headers) => {
    setVisibleColumnsOrder(['Sheet', 'Baris', ...headers]);
  };

  // Clear all columns (including Sheet & Baris)
  const clearAllColumns = () => {
    setVisibleColumnsOrder(['__all_hidden__']);
  };

  // Helper to check if a bookmark column is visible
  const isBookmarkColumnVisible = (colName) => {
    if (bookmarkVisibleColumns.length === 0) return true;
    if (bookmarkVisibleColumns.includes('__all_hidden__')) return false;
    return bookmarkVisibleColumns.includes(colName);
  };

  // Toggle bookmark column visibility
  const toggleBookmarkColumnVisibility = (colName, headers) => {
    setBookmarkVisibleColumns((prev) => {
      if (prev.includes('__all_hidden__')) {
        return [colName];
      }
      let currentOrder = prev.length === 0
        ? ['File', 'Sheet', 'Baris', ...headers]
        : [...prev];
      if (currentOrder.includes(colName)) {
        return currentOrder.filter(c => c !== colName);
      } else {
        return [...currentOrder, colName];
      }
    });
  };

  // Select all columns for bookmarks
  const selectBookmarkAllColumns = (headers) => {
    setBookmarkVisibleColumns(['File', 'Sheet', 'Baris', ...headers]);
  };

  // Clear all columns for bookmarks
  const clearBookmarkAllColumns = () => {
    setBookmarkVisibleColumns(['__all_hidden__']);
  };

  // Helper to get headers in their selected order
  const getOrderedColumns = () => {
    if (visibleColumnsOrder.length === 0) {
      return ['Sheet', 'Baris', ...activeHeaders];
    }
    return visibleColumnsOrder.filter(h => 
      h === 'Sheet' || h === 'Baris' || activeHeaders.includes(h)
    );
  };

  // Fetch bookmarked rows from server
  const fetchBookmarks = async () => {
    setLoadingBookmarks(true);
    try {
      const res = await fetch('/api/bookmarks');
      if (res.ok) {
        const data = await res.json();
        setBookmarksList(data.bookmarks || []);
      }
    } catch (err) {
      console.error('Gagal memuat bookmark:', err);
    } finally {
      setLoadingBookmarks(false);
    }
  };

  // Toggle row bookmark state
  const handleToggleBookmark = async (rowId, currentStatus) => {
    try {
      const method = currentStatus ? 'DELETE' : 'POST';
      const url = currentStatus ? `/api/bookmarks/${rowId}` : '/api/bookmarks';
      const body = currentStatus ? null : JSON.stringify({ rowId });

      const res = await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body
      });

      if (res.ok) {
        // Toggle locally in searchResults
        setSearchResults(prev => prev.map(file => ({
          ...file,
          matches: file.matches.map(m => m.id === rowId ? { ...m, isBookmarked: !currentStatus } : m)
        })));

        // If bookmarks list is active or displayed, refresh it
        fetchBookmarks();
        showToast(
          currentStatus ? '⭐ Bookmark Dihapus' : '⭐ Tersimpan ke Bookmark',
          currentStatus ? 'Baris dokumen berhasil dihapus dari bookmark.' : 'Baris dokumen berhasil disimpan.',
          'success'
        );
      }
    } catch (err) {
      console.error('Gagal toggle bookmark:', err);
      showToast('⚠️ Gagal', 'Gagal memproses bookmark.', 'error');
    }
  };

  // Bulk delete selected files
  const handleBulkDelete = async () => {
    if (selectedFileIds.length === 0) return;

    const filenames = filesList
      .filter(f => selectedFileIds.includes(f.id))
      .map(f => f.filename)
      .join(', ');

    if (!window.confirm(`Apakah Anda yakin ingin menghapus ${selectedFileIds.length} berkas berikut beserta seluruh datanya?\n\n${filenames}`)) {
      return;
    }

    try {
      const res = await fetch('/api/files/bulk-delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids: selectedFileIds })
      });

      const data = await res.json();
      if (res.ok) {
        showToast('🗑️ Penghapusan Massal Sukses', `${selectedFileIds.length} file berhasil dihapus secara permanen.`, 'success');
        setSelectedFileIds([]);
        setSelectedFileId(null);
        fetchFiles();
      } else {
        showToast('⚠️ Gagal Hapus Massal', data.error || 'Terjadi kesalahan.', 'error');
      }
    } catch (err) {
      console.error('Error bulk delete:', err);
      showToast('⚠️ Error Koneksi', 'Gagal menghubungi server.', 'error');
    }
  };

  // Delete file
  const handleDeleteFile = async (id, filename) => {
    if (!window.confirm(`Apakah Anda yakin ingin menghapus file "${filename}"? Semua data di dalamnya akan terhapus permanen dari aplikasi.`)) {
      return;
    }

    try {
      const res = await fetch(`/api/files/${id}`, { method: 'DELETE' });
      if (res.ok) {
        showToast('🗑️ File Dihapus', `Berkas "${filename}" berhasil dihapus.`, 'success');
        setSelectedFileId(null);
        fetchFiles();
      } else {
        const data = await res.json();
        alert('Gagal menghapus file: ' + (data.error || 'Unknown error'));
      }
    } catch (err) {
      alert('Error saat menghapus file: ' + err.message);
    }
  };

  // Handle Drag-and-Drop events
  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFileUpload(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      processFileUpload(e.target.files[0]);
    }
  };

  // Upload file logic with background job status polling
  const processFileUpload = async (file) => {
    const fileExt = file.name.split('.').pop().toLowerCase();
    if (fileExt !== 'xlsx' && fileExt !== 'xls') {
      setUploadError('Tipe file salah! Harap unggah file dengan format .xlsx atau .xls');
      setUploadSuccess(null);
      return;
    }

    setUploading(true);
    setUploadProgress(10);
    setUploadStatusMessage('Mengunggah berkas ke server...');
    setUploadError(null);
    setUploadSuccess(null);

    const formData = new FormData();
    formData.append('file', file);

    try {
      // 1. Initial POST request
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      });

      const uploadData = await res.json();
      if (!res.ok) {
        throw new Error(uploadData.error || 'Gagal mengunggah file.');
      }

      const jobId = uploadData.jobId;
      setUploadProgress(20);
      setUploadStatusMessage('Berkas terunggah. Memulai impor data...');

      // 2. Poll status endpoint until completed or failed
      const pollInterval = setInterval(async () => {
        try {
          const statusRes = await fetch(`/api/upload/status/${jobId}`);
          if (!statusRes.ok) {
            clearInterval(pollInterval);
            setUploading(false);
            setUploadError('Gagal memantau status impor.');
            return;
          }

          const job = await statusRes.json();
          setUploadProgress(job.progress);
          setUploadStatusMessage(job.message);

          if (job.status === 'completed') {
            clearInterval(pollInterval);
            setUploadSuccess(job.result);
            setUploading(false);
            
            // Toast notification on upload complete
            showToast(
              '✅ Upload Selesai!',
              `File "${job.result.filename}" berhasil diimpor. ${job.result.rowsInserted?.toLocaleString('id-ID') || 0} baris data siap dicari.`,
              'upload'
            );
            
            // Set newly uploaded file as active preview file
            setSelectedFileId(job.result.fileId);
            fetchFiles(); 
          } else if (job.status === 'error') {
            clearInterval(pollInterval);
            setUploadError(job.error || 'Terjadi kesalahan saat memproses file.');
            setUploading(false);
            showToast('⚠️ Upload Gagal', job.error || 'Terjadi kesalahan saat memproses file.', 'error');
          }
        } catch (pollErr) {
          clearInterval(pollInterval);
          setUploading(false);
          setUploadError('Terjadi kesalahan koneksi saat memantau status.');
        }
      }, 1000);

    } catch (err) {
      setUploading(false);
      setUploadProgress(0);
      setUploadError(err.message || 'Gagal menghubungi server.');
    }
  };

  // Helper function to highlight query matches (token-based / Google-style)
  const highlightText = (text, highlight) => {
    if (text === null || text === undefined) return '';
    const textStr = text.toString();
    if (!highlight || highlight.trim() === '') return textStr;
    
    // Split highlight query into individual tokens (split by spaces or dots)
    const baseTokens = highlight.trim().split(/[\s\.]+/).filter(Boolean);
    if (baseTokens.length === 0) return textStr;

    // Generate tokens, splitting mixed alphanumeric tokens (like A001 into A and 001) for partial highlights
    const tokens = [];
    baseTokens.forEach(t => {
      tokens.push(t);
      const match = t.match(/^([a-zA-Z]+)(0*[1-9]\d*)$/);
      if (match) {
        tokens.push(match[1]); // letters (e.g. A)
        tokens.push(match[2]); // digits (e.g. 001)
      }
    });
    if (tokens.length === 0) return textStr;
    
    // Create a regex that matches any of the tokens
    // Escape special characters for each token
    const escapedTokens = tokens.map(t => t.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'));
    const regexPattern = `(${escapedTokens.join('|')})`;
    const regex = new RegExp(regexPattern, 'gi');
    
    const parts = textStr.split(regex);
    
    return (
      <>
        {parts.map((part, i) => 
          regex.test(part) 
            ? <mark key={i} className="highlight">{part}</mark> 
            : part
        )}
      </>
    );
  };

  // Copy visible table data to clipboard as TSV (paste-ready for Excel/Sheets)
  const handleCopyData = (orderedDataHeaders, showSheet, showBaris) => {
    if (!selectedFileData) return;
    const headers = [
      ...(showSheet ? ['Sheet'] : []),
      ...(showBaris ? ['Baris'] : []),
      ...orderedDataHeaders.map(h => getColumnLabel(h))
    ];
    const rows = selectedFileData.matches.map(match => [
      ...(showSheet ? [match.sheetName] : []),
      ...(showBaris ? [match.rowNumber] : []),
      ...orderedDataHeaders.map(h => match.rowData[h] ?? '')
    ]);
    const tsv = [headers, ...rows]
      .map(row => row.map(cell => String(cell).replace(/\t/g, ' ')).join('\t'))
      .join('\n');
    navigator.clipboard.writeText(tsv).then(() => {
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2500);
    });
  };

  // Open print-friendly window for the visible table
  const handlePrint = (orderedDataHeaders, showSheet, showBaris) => {
    if (!selectedFileData) return;
    const headers = [
      ...(showSheet ? ['Sheet'] : []),
      ...(showBaris ? ['Baris'] : []),
      ...orderedDataHeaders.map(h => getColumnLabel(h))
    ];
    const rows = selectedFileData.matches.map(match => [
      ...(showSheet ? [match.sheetName] : []),
      ...(showBaris ? [match.rowNumber] : []),
      ...orderedDataHeaders.map(h => match.rowData[h] ?? '')
    ]);
    const tableHTML = `
      <table border="1" cellspacing="0" cellpadding="6"
        style="border-collapse:collapse;width:100%;font-size:11px;font-family:Arial,sans-serif">
        <thead style="background:#e8eaf0">
          <tr>${headers.map(h => `<th style="text-align:left;padding:6px 8px;border:1px solid #ccc">${h}</th>`).join('')}</tr>
        </thead>
        <tbody>
          ${rows.map((row, ri) =>
            `<tr style="background:${ri % 2 === 0 ? '#fff' : '#f5f7fb'}">
              ${row.map(cell => `<td style="padding:5px 8px;border:1px solid #ddd">${cell ?? ''}</td>`).join('')}
            </tr>`
          ).join('')}
        </tbody>
      </table>`;
    const win = window.open('', '_blank');
    win.document.write(`<!DOCTYPE html><html><head>
      <meta charset="UTF-8">
      <title>Cetak Data — ${selectedFileData.filename}</title>
      <style>
        body { margin: 20px; font-family: Arial, sans-serif; color: #222; }
        h2   { font-size: 14px; margin: 0 0 4px; }
        p    { font-size: 11px; color: #666; margin: 0 0 14px; }
        @media print { @page { margin: 1.5cm; size: landscape; } }
      </style>
      </head><body>
      <h2>📁 ${selectedFileData.filename}</h2>
      <p>${searchQuery.trim()
        ? `Kata kunci: "${searchQuery}" · ${selectedFileData.matches.length} baris cocok`
        : `Pratinjau · ${selectedFileData.matches.length} baris pertama`
      }</p>
      ${tableHTML}
      <script>window.onload = () => { window.print(); }<\/script>
      </body></html>`);
    win.document.close();
  };

  // Find headers (keys) from all rows in the active file search results
  const getActiveFileHeaders = () => {
    const activeFile = searchResults.find(f => f.fileId === selectedFileId);
    if (!activeFile || activeFile.matches.length === 0) return [];
    
    // Aggregate all keys from all matching rows
    const allKeys = new Set();
    activeFile.matches.forEach(m => {
      Object.keys(m.rowData).forEach(key => allKeys.add(key));
    });
    
    const headersArray = Array.from(allKeys);
    
    // Move pertelaan keys (Kode Kopel) to the very beginning (before Kolom_1)
    const pertelaanHeaders = headersArray.filter(h => h.toLowerCase().includes('pertelaan arsip'));
    const otherHeaders = headersArray.filter(h => !h.toLowerCase().includes('pertelaan arsip'));
    
    return [...pertelaanHeaders, ...otherHeaders];
  };

  const activeHeaders = getActiveFileHeaders();
  const selectedFileData = searchResults.find(f => f.fileId === selectedFileId);

  return (
    <div className="app-container">
      {/* Dynamic Header */}
      <header className="app-header">
        <div className="logo-section">
          <div className="logo-icon">📊</div>
          <div>
            <h1>SpreadSheet Finder</h1>
            <p className="subtitle">Cari data di seluruh dokumen Excel secara instan</p>
          </div>
        </div>
        
        {/* Status & Stats info */}
        <div className="header-meta">
          <div className="stats-badge">
            <span className="stats-item">📁 <strong>{stats.totalFiles}</strong> File</span>
            <span className="stats-divider">|</span>
            <span className="stats-item">⚡ <strong>{stats.totalRows.toLocaleString('id-ID')}</strong> Baris Data</span>
          </div>
          
          {/* Theme Toggle */}
          <button
            className="theme-toggle-btn"
            onClick={() => setIsDarkMode(m => !m)}
            title={isDarkMode ? 'Beralih ke Mode Terang' : 'Beralih ke Mode Gelap'}
          >
            {isDarkMode ? '☀️ Terang' : '🌙 Gelap'}
          </button>

          <div className={`db-status-pill ${dbConnected ? 'connected' : 'disconnected'}`}>
            <span className="dot"></span>
            <span>{dbConnected ? 'Database Connected' : 'Database Offline'}</span>
          </div>
        </div>
      </header>

      {/* Main Tab Navigation */}
      <nav className="tab-navigation">
        <button 
          className={`tab-btn ${activeTab === 'search' ? 'active' : ''}`}
          onClick={() => setActiveTab('search')}
        >
          🔍 Pencarian Data
        </button>
        <button 
          className={`tab-btn ${activeTab === 'bookmarks' ? 'active' : ''}`}
          onClick={() => setActiveTab('bookmarks')}
        >
          ⭐ Bookmark ({bookmarksList.length})
        </button>
        <button 
          className={`tab-btn ${activeTab === 'upload' ? 'active' : ''}`}
          onClick={() => setActiveTab('upload')}
        >
          📤 Unggah Excel
        </button>
        <button 
          className={`tab-btn ${activeTab === 'files' ? 'active' : ''}`}
          onClick={() => setActiveTab('files')}
        >
          ⚙️ Kelola File ({filesList.length})
        </button>
        <button 
          className={`tab-btn ${activeTab === 'logs' ? 'active' : ''}`}
          onClick={() => setActiveTab('logs')}
        >
          📜 Log Aktivitas
        </button>
        <button 
          className={`tab-btn ${activeTab === 'dashboard' ? 'active' : ''}`}
          onClick={() => { setActiveTab('dashboard'); fetchDashboard(); }}
        >
          📊 Dashboard
        </button>
      </nav>

      {/* Main Content Area */}
      <main className="main-content">
        
        {/* TAB 1: SEARCH */}
        {activeTab === 'search' && (
          <div className="search-tab-content">
            <div className="search-bar-row">
              <div className="search-bar-container">
                <span className="search-icon">🔍</span>
                <input
                  ref={searchInputRef}
                  type="text"
                  className="search-input"
                  placeholder="Cari nama, alamat, nomor telepon, kode barang... (Tekan '/' untuk fokus)"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                />
                {searchQuery && (
                  <button className="clear-btn" onClick={() => setSearchQuery('')}>✕</button>
                )}
              </div>
              <button
                className={`adv-filter-toggle-btn ${showAdvancedFilters ? 'active' : ''}`}
                onClick={() => setShowAdvancedFilters(!showAdvancedFilters)}
                title="Saring berdasarkan sheet atau unit tertentu"
              >
                ⚙️ Filter Lanjutan
              </button>
            </div>

            {/* Advanced Filters Panel */}
            {showAdvancedFilters && (
              <div className="advanced-filters-panel">
                <div className="filter-group">
                  <label>Saring Sheet:</label>
                  <input
                    type="text"
                    className="filter-input-field"
                    placeholder="Contoh: Sheet1, Central"
                    value={filterSheet}
                    onChange={(e) => setFilterSheet(e.target.value)}
                  />
                </div>
                <div className="filter-group">
                  <label>Saring Unit (Kode):</label>
                  <input
                    type="text"
                    className="filter-input-field"
                    placeholder="Contoh: 011, 012"
                    value={filterUnit}
                    onChange={(e) => setFilterUnit(e.target.value)}
                  />
                </div>
                {(filterSheet || filterUnit) && (
                  <button
                    className="clear-filters-btn"
                    onClick={() => { setFilterSheet(''); setFilterUnit(''); }}
                  >
                    Reset Filter
                  </button>
                )}
              </div>
            )}

            {/* Search History Chips */}
            {searchQuery.trim() === '' && searchHistory.length > 0 && (
              <div className="search-history-container">
                <span className="history-title">Riwayat Pencarian:</span>
                <div className="history-chips">
                  {searchHistory.map((q, idx) => (
                    <div
                      key={idx}
                      className="history-chip"
                      onClick={() => setSearchQuery(q)}
                    >
                      <span>{q}</span>
                      <button
                        className="delete-history-btn"
                        onClick={(e) => removeFromHistory(e, q)}
                        title="Hapus riwayat"
                      >
                        ✕
                      </button>
                    </div>
                  ))}
                  <button className="clear-all-history-btn" onClick={clearHistory}>
                    Hapus Semua
                  </button>
                </div>
              </div>
            )}

            {searchUIState === 'loading' && (
              <div className="search-loader">
                <div className="spinner"></div>
                <p>Mengambil data dari database...</p>
              </div>
            )}

            {/* Results Rendering */}
            {(searchUIState === 'preview' || searchUIState === 'results') && searchResults.length > 0 && (
              <div className="results-container">
                
                {/* Left Side: Sidebar of files (either containing matches or all uploaded files if browse mode) */}
                <div className="results-sidebar">
                  <h3 className="section-title">
                    {searchQuery.trim() === '' ? '📁 Daftar Dokumen' : '🎯 Dokumen Terkait'}
                  </h3>
                  <div className="file-tabs">
                    {searchQuery.trim() === '' ? (
                      filesList.map(file => (
                        <button
                          key={file.id}
                          className={`file-tab-btn ${selectedFileId === file.id ? 'active' : ''}`}
                          onClick={() => handleFileTabClick(file.id)}
                        >
                          <div className="file-tab-name">📄 {file.filename}</div>
                          <div className="file-tab-badge">⚡ {file.row_count} baris</div>
                        </button>
                      ))
                    ) : (
                      searchResults.map(file => (
                        <button
                          key={file.fileId}
                          className={`file-tab-btn ${selectedFileId === file.fileId ? 'active' : ''}`}
                          onClick={() => handleFileTabClick(file.fileId)}
                        >
                          <div className="file-tab-name">📄 {file.filename}</div>
                          <div className="file-tab-badge">{file.matches.length} baris cocok</div>
                        </button>
                      ))
                    )}
                  </div>
                </div>

                {/* Right Side: Data table of the selected file */}
                <div className="results-view">
                  {selectedFileData && (
                    <div className="results-card">
                      <div className="results-card-header">
                        <div>
                          <h2>📁 {selectedFileData.filename}</h2>
                          <p className="file-meta-date">
                            {searchQuery.trim() === '' 
                              ? `Mode Pratinjau (Menampilkan 200 baris pertama) • Diunggah: ${new Date(selectedFileData.uploadedAt).toLocaleString('id-ID')}`
                              : `Hasil Pencarian • Diunggah: ${new Date(selectedFileData.uploadedAt).toLocaleString('id-ID')}`
                            }
                          </p>
                        </div>

                        {/* Column Filter Toggle Button */}
                        <div className="column-filter-container">
                          <button 
                            className={`column-filter-trigger-btn ${showColumnFilter ? 'active' : ''}`}
                            onClick={() => {
                              setShowColumnFilter(!showColumnFilter);
                              if (showColumnFilter) setColumnSearchQuery('');
                            }}
                          >
                            {showColumnFilter ? '✕ Tutup Pilihan' : '⚙️ Pilih Kolom'}
                          </button>
                          
                          {showColumnFilter && (
                            <div className="column-filter-dropdown">
                              <div className="column-search-wrapper">
                                <input
                                  type="text"
                                  className="column-search-input"
                                  placeholder="Cari nama kolom..."
                                  value={columnSearchQuery}
                                  onChange={(e) => setColumnSearchQuery(e.target.value)}
                                  onClick={(e) => e.stopPropagation()}
                                />
                                {columnSearchQuery && (
                                  <button 
                                    className="column-search-clear-btn"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      setColumnSearchQuery('');
                                    }}
                                  >
                                    ✕
                                  </button>
                                )}
                              </div>
                              <div className="dropdown-actions">
                                <button className="dropdown-action-btn" onClick={() => { selectAllColumns(activeHeaders); setShowColumnFilter(false); }}>Tampilkan Semua</button>
                                <button className="dropdown-action-btn" onClick={() => { clearAllColumns(activeHeaders); setShowColumnFilter(false); }}>Sembunyikan Semua</button>
                                <button className="dropdown-action-btn dropdown-action-btn--close" onClick={() => setShowColumnFilter(false)}>✕ Tutup</button>
                              </div>
                              <div className="dropdown-items">
                                {['Sheet', 'Baris']
                                  .filter(h => h.toLowerCase().includes(columnSearchQuery.toLowerCase()))
                                  .map(h => (
                                    <label key={h} className="column-checkbox-label meta-checkbox">
                                      <input
                                        type="checkbox"
                                        checked={isColumnVisible(h)}
                                        onChange={() => toggleColumnVisibility(h)}
                                      />
                                      <strong>{h}</strong>
                                    </label>
                                  ))}
                                {['Sheet', 'Baris'].filter(h => h.toLowerCase().includes(columnSearchQuery.toLowerCase())).length > 0 &&
                                 activeHeaders.filter(h => getColumnLabel(h).toLowerCase().includes(columnSearchQuery.toLowerCase())).length > 0 && (
                                  <div className="dropdown-divider"></div>
                                )}
                                {activeHeaders
                                  .filter(h => getColumnLabel(h).toLowerCase().includes(columnSearchQuery.toLowerCase()))
                                  .map(h => (
                                    <label key={h} className="column-checkbox-label">
                                      <input
                                        type="checkbox"
                                        checked={isColumnVisible(h)}
                                        onChange={() => toggleColumnVisibility(h)}
                                      />
                                      <span>{getColumnLabel(h)}</span>
                                    </label>
                                  ))}
                                {['Sheet', 'Baris'].filter(h => h.toLowerCase().includes(columnSearchQuery.toLowerCase())).length === 0 &&
                                 activeHeaders.filter(h => getColumnLabel(h).toLowerCase().includes(columnSearchQuery.toLowerCase())).length === 0 && (
                                  <div className="dropdown-no-results">Kolom tidak ditemukan</div>
                                )}
                              </div>
                            </div>
                          )}
                        </div>
                      </div>

                      {(() => {
                        const orderedDataHeaders = visibleColumnsOrder.length === 0
                          ? activeHeaders
                          : visibleColumnsOrder.filter(h => h !== 'Sheet' && h !== 'Baris' && activeHeaders.includes(h));

                        const showSheet = isColumnVisible('Sheet');
                        const showBaris = isColumnVisible('Baris');
                        const noColumnsVisible = !showSheet && !showBaris && orderedDataHeaders.length === 0;

                        if (noColumnsVisible) {
                          return (
                            <div className="no-columns-state">
                              <div className="no-columns-icon">👁️</div>
                              <p>Semua kolom disembunyikan.<br/>Klik <strong>⚙️ Pilih Kolom</strong> lalu <strong>Tampilkan Semua</strong> untuk menampilkan data.</p>
                            </div>
                          );
                        }

                        return (
                          <>
                            {/* Column Filter Active Info Bar */}
                            {visibleColumnsOrder.length > 0 && !visibleColumnsOrder.includes('__all_hidden__') && (
                              <div className="active-filter-bar">
                                <span className="active-filter-text">
                                  ℹ️ <strong>Filter Kolom Aktif:</strong> Menampilkan {showSheet ? 1 : 0 + (showBaris ? 1 : 0) + orderedDataHeaders.length} dari {activeHeaders.length + 2} kolom.
                                </span>
                                <button className="reset-filter-link-btn" onClick={() => setVisibleColumnsOrder([])}>
                                  ✕ Reset Filter Kolom
                                </button>
                              </div>
                            )}

                            {/* Action buttons: Copy + Print + Export + Share */}
                            <div className="table-action-bar">
                              <button
                                className={`table-action-btn ${copySuccess ? 'success' : ''}`}
                                onClick={() => handleCopyData(orderedDataHeaders, showSheet, showBaris)}
                                title="Salin data ke clipboard (format TSV, bisa ditempel ke Excel)"
                              >
                                {copySuccess ? '✅ Tersalin!' : '📋 Salin Data'}
                              </button>
                              <button
                                className="table-action-btn table-action-btn--export"
                                onClick={() => handleExportExcel(orderedDataHeaders, showSheet, showBaris)}
                                title="Unduh data sebagai file Excel (.xlsx)"
                              >
                                📥 Ekspor Excel
                              </button>
                              <button
                                className="table-action-btn"
                                onClick={() => handlePrint(orderedDataHeaders, showSheet, showBaris)}
                                title="Buka jendela cetak"
                              >
                                🖨️ Cetak
                              </button>
                              {searchQuery.trim() && (
                                <button
                                  className="table-action-btn table-action-btn--share"
                                  onClick={handleShareLink}
                                  title="Salin link pencarian ini ke clipboard"
                                >
                                  🔗 Bagikan
                                </button>
                              )}
                              <span className="table-row-count">
                                {selectedFileData.matches.length} baris
                              </span>
                            </div>
                            <div className="table-responsive">
                              <table className="excel-table">
                                <thead>
                                  <tr>
                                    <th style={{ width: '40px', textAlign: 'center' }}>⭐</th>
                                    {showSheet && <th>Sheet</th>}
                                    {showBaris && <th>Baris</th>}
                                    {orderedDataHeaders.map(h => (
                                      <th key={h}>{getColumnLabel(h)}</th>
                                    ))}
                                  </tr>
                                </thead>
                                <tbody>
                                  {selectedFileData.matches.map(match => (
                                    <tr key={match.id}>
                                      <td className="meta-cell text-center" style={{ width: '40px' }}>
                                        <button
                                          className={`bookmark-star-btn ${match.isBookmarked ? 'active' : ''}`}
                                          onClick={() => handleToggleBookmark(match.id, match.isBookmarked)}
                                          title={match.isBookmarked ? 'Hapus dari Bookmark' : 'Simpan ke Bookmark'}
                                        >
                                          {match.isBookmarked ? '★' : '☆'}
                                        </button>
                                      </td>
                                      {showSheet && (
                                        <td className="meta-cell font-accent">{match.sheetName}</td>
                                      )}
                                      {showBaris && (
                                        <td className="meta-cell text-center">{match.rowNumber}</td>
                                      )}
                                      {orderedDataHeaders.map(h => {
                                        const isDesc = h === 'Kolom_17' || h === 'Kolom_18' || h.toLowerCase().includes('perihal') || h.toLowerCase().includes('uraian');
                                        return (
                                          <td key={h} className={isDesc ? 'description-cell' : ''}>
                                            {highlightText(match.rowData[h], searchQuery)}
                                          </td>
                                        );
                                      })}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </>
                        );
                      })()}
                    </div>
                  )}
                </div>

              </div>
            )}

            {/* Empty and Neutral States */}
            {searchUIState === 'empty' && (
              <div className="empty-state">
                <div className="empty-icon">🤷‍♂️</div>
                <h3>Tidak ada data yang cocok</h3>
                <p>Coba gunakan kata kunci lain atau unggah dokumen Excel baru yang relevan.</p>
              </div>
            )}

            {searchUIState === 'welcome' && (
              <div className="empty-state welcome-state">
                <div className="welcome-glow-icon">🔍</div>
                <h3>Pencarian Data Global</h3>
                <p>Masukkan kata kunci untuk mencari data atau unggah dokumen Excel terlebih dahulu di tab <strong>Unggah Excel</strong>.</p>
                <div className="welcome-shortcuts">
                  <div className="shortcut-card">
                    <span className="shortcut-key">/</span>
                    <span className="shortcut-desc">Fokus ke Input Pencarian</span>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* TAB 1b: BOOKMARKS */}
        {activeTab === 'bookmarks' && (
          <div className="bookmarks-tab-content">
            <div className="dashboard-header">
              <h2>⭐ Data Bookmark</h2>
              <p>Daftar baris data spreadsheet penting yang Anda simpan untuk referensi cepat.</p>
            </div>

            {loadingBookmarks ? (
              <div className="search-loader">
                <div className="spinner"></div>
                <p>Mengambil data bookmark...</p>
              </div>
            ) : bookmarksList.length === 0 ? (
              <div className="empty-state">
                <div className="empty-icon">⭐</div>
                <h3>Belum ada bookmark</h3>
                <p>Klik tombol bintang (☆) pada tabel hasil pencarian untuk menyimpan data penting di sini.</p>
              </div>
            ) : (
              <div className="results-card">
                {(() => {
                  // Find all headers across bookmarked rows
                  const allKeys = new Set();
                  bookmarksList.forEach(m => {
                    Object.keys(m.row_data).forEach(key => allKeys.add(key));
                  });
                  const bookmarkedHeaders = Array.from(allKeys);
                  
                  const orderedBookmarkHeaders = bookmarkVisibleColumns.length === 0
                    ? bookmarkedHeaders
                    : bookmarkVisibleColumns.filter(h => h !== 'File' && h !== 'Sheet' && h !== 'Baris' && bookmarkedHeaders.includes(h));

                  const showFile = isBookmarkColumnVisible('File');
                  const showSheet = isBookmarkColumnVisible('Sheet');
                  const showBaris = isBookmarkColumnVisible('Baris');
                  const noColumnsVisible = !showFile && !showSheet && !showBaris && orderedBookmarkHeaders.length === 0;

                  return (
                    <>
                      <div className="results-card-header" style={{ marginBottom: '1.25rem' }}>
                        <div>
                          <h3 style={{ margin: 0, fontSize: '1.1rem' }}>📋 Filter & Tampilan Bookmark</h3>
                          <p style={{ margin: '4px 0 0', fontSize: '0.82rem', color: 'var(--text-muted)' }}>
                            Sesuaikan kolom apa saja yang ingin ditampilkan di tabel bookmark Anda.
                          </p>
                        </div>
                        
                        {/* Bookmark Column Filter Toggle Button */}
                        <div className="column-filter-container">
                          <button 
                            className={`column-filter-trigger-btn ${showBookmarkColumnFilter ? 'active' : ''}`}
                            onClick={() => {
                              setShowBookmarkColumnFilter(!showBookmarkColumnFilter);
                              if (showBookmarkColumnFilter) setBookmarkColumnSearchQuery('');
                            }}
                          >
                            {showBookmarkColumnFilter ? '✕ Tutup Pilihan' : '⚙️ Pilih Kolom'}
                          </button>
                          
                          {showBookmarkColumnFilter && (
                            <div className="column-filter-dropdown">
                              <div className="column-search-wrapper">
                                <input
                                  type="text"
                                  className="column-search-input"
                                  placeholder="Cari nama kolom..."
                                  value={bookmarkColumnSearchQuery}
                                  onChange={(e) => setBookmarkColumnSearchQuery(e.target.value)}
                                  onClick={(e) => e.stopPropagation()}
                                />
                                {bookmarkColumnSearchQuery && (
                                  <button 
                                    className="column-search-clear-btn"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      setBookmarkColumnSearchQuery('');
                                    }}
                                  >
                                    ✕
                                  </button>
                                )}
                              </div>
                              <div className="dropdown-actions">
                                <button className="dropdown-action-btn" onClick={() => { selectBookmarkAllColumns(bookmarkedHeaders); setShowBookmarkColumnFilter(false); }}>Tampilkan Semua</button>
                                <button className="dropdown-action-btn" onClick={() => { clearBookmarkAllColumns(); setShowBookmarkColumnFilter(false); }}>Sembunyikan Semua</button>
                                <button className="dropdown-action-btn dropdown-action-btn--close" onClick={() => setShowBookmarkColumnFilter(false)}>✕ Tutup</button>
                              </div>
                              <div className="dropdown-items">
                                {['File', 'Sheet', 'Baris']
                                  .filter(h => h.toLowerCase().includes(bookmarkColumnSearchQuery.toLowerCase()))
                                  .map(h => (
                                    <label key={h} className="column-checkbox-label meta-checkbox">
                                      <input
                                        type="checkbox"
                                        checked={isBookmarkColumnVisible(h)}
                                        onChange={() => toggleBookmarkColumnVisibility(h, bookmarkedHeaders)}
                                      />
                                      <strong>{h}</strong>
                                    </label>
                                  ))}
                                {['File', 'Sheet', 'Baris'].filter(h => h.toLowerCase().includes(bookmarkColumnSearchQuery.toLowerCase())).length > 0 &&
                                 bookmarkedHeaders.filter(h => getColumnLabel(h).toLowerCase().includes(bookmarkColumnSearchQuery.toLowerCase())).length > 0 && (
                                  <div className="dropdown-divider"></div>
                                )}
                                {bookmarkedHeaders
                                  .filter(h => getColumnLabel(h).toLowerCase().includes(bookmarkColumnSearchQuery.toLowerCase()))
                                  .map(h => (
                                    <label key={h} className="column-checkbox-label">
                                      <input
                                        type="checkbox"
                                        checked={isBookmarkColumnVisible(h)}
                                        onChange={() => toggleBookmarkColumnVisibility(h, bookmarkedHeaders)}
                                      />
                                      <span>{getColumnLabel(h)}</span>
                                    </label>
                                  ))}
                                {['File', 'Sheet', 'Baris'].filter(h => h.toLowerCase().includes(bookmarkColumnSearchQuery.toLowerCase())).length === 0 &&
                                 bookmarkedHeaders.filter(h => getColumnLabel(h).toLowerCase().includes(bookmarkColumnSearchQuery.toLowerCase())).length === 0 && (
                                  <div className="dropdown-no-results">Kolom tidak ditemukan</div>
                                )}
                              </div>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Bookmark Column Filter Active Info Bar */}
                      {bookmarkVisibleColumns.length > 0 && !bookmarkVisibleColumns.includes('__all_hidden__') && (
                        <div className="active-filter-bar" style={{ marginBottom: '1rem' }}>
                          <span className="active-filter-text">
                            ℹ️ <strong>Filter Kolom Aktif:</strong> Menampilkan { (showFile ? 1 : 0) + (showSheet ? 1 : 0) + (showBaris ? 1 : 0) + orderedBookmarkHeaders.length } dari { bookmarkedHeaders.length + 3 } kolom.
                          </span>
                          <button className="reset-filter-link-btn" onClick={() => setBookmarkVisibleColumns([])}>
                            ✕ Reset Filter Kolom
                          </button>
                        </div>
                      )}

                      {noColumnsVisible ? (
                        <div className="no-columns-state">
                          <div className="no-columns-icon">👁️</div>
                          <p>Semua kolom disembunyikan.<br/>Klik <strong>⚙️ Pilih Kolom</strong> lalu <strong>Tampilkan Semua</strong> untuk menampilkan data bookmark.</p>
                        </div>
                      ) : (
                        <>
                          {/* Action buttons */}
                          <div className="table-action-bar">
                            <button
                              className="table-action-btn table-action-btn--export"
                              onClick={() => {
                                const headers = [
                                  ...(showFile ? ['File'] : []),
                                  ...(showSheet ? ['Sheet'] : []),
                                  ...(showBaris ? ['Baris'] : []),
                                  ...orderedBookmarkHeaders.map(h => getColumnLabel(h))
                                ];
                                const rows = bookmarksList.map(m => [
                                  ...(showFile ? [m.filename] : []),
                                  ...(showSheet ? [m.sheet_name] : []),
                                  ...(showBaris ? [m.row_number] : []),
                                  ...orderedBookmarkHeaders.map(h => m.row_data[h] ?? '')
                                ]);
                                const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                                const wb = XLSX.utils.book_new();
                                XLSX.utils.book_append_sheet(wb, ws, 'Bookmarks');
                                XLSX.writeFile(wb, `Bookmarks-${new Date().toLocaleDateString('id-ID').replace(/\//g, '-')}.xlsx`);
                                showToast('📥 Berhasil Diekspor!', 'Bookmarks disimpan ke Excel.', 'success');
                              }}
                              title="Unduh seluruh bookmark ke Excel"
                            >
                              📥 Ekspor Excel ({bookmarksList.length})
                            </button>
                            <button
                              className="table-action-btn"
                              onClick={() => {
                                const headers = [
                                  ...(showFile ? ['File'] : []),
                                  ...(showSheet ? ['Sheet'] : []),
                                  ...(showBaris ? ['Baris'] : []),
                                  ...orderedBookmarkHeaders.map(h => getColumnLabel(h))
                                ];
                                const rows = bookmarksList.map(m => [
                                  ...(showFile ? [m.filename] : []),
                                  ...(showSheet ? [m.sheet_name] : []),
                                  ...(showBaris ? [m.row_number] : []),
                                  ...orderedBookmarkHeaders.map(h => m.row_data[h] ?? '')
                                ]);
                                const tableHTML = `
                                  <table border="1" cellspacing="0" cellpadding="6"
                                    style="border-collapse:collapse;width:100%;font-size:10px;font-family:Arial,sans-serif">
                                    <thead style="background:#e8eaf0">
                                      <tr>${headers.map(h => `<th style="text-align:left;padding:6px;border:1px solid #ccc">${h}</th>`).join('')}</tr>
                                    </thead>
                                    <tbody>
                                      ${rows.map((row, ri) =>
                                        `<tr style="background:${ri % 2 === 0 ? '#fff' : '#f5f7fb'}">
                                          ${row.map(cell => `<td style="padding:5px;border:1px solid #ddd">${cell ?? ''}</td>`).join('')}
                                        </tr>`
                                      ).join('')}
                                    </tbody>
                                  </table>`;
                                const win = window.open('', '_blank');
                                win.document.write(`<!DOCTYPE html><html><head><title>Bookmarks Cetak</title></head><body onload="window.print()">${tableHTML}</body></html>`);
                                win.document.close();
                              }}
                            >
                              🖨️ Cetak
                            </button>
                            <span className="table-row-count">{bookmarksList.length} baris tersimpan</span>
                          </div>

                          <div className="table-responsive">
                            <table className="excel-table">
                              <thead>
                                <tr>
                                  <th style={{ width: '40px', textAlign: 'center' }}>⭐</th>
                                  {showFile && <th>Nama Berkas</th>}
                                  {showSheet && <th>Sheet</th>}
                                  {showBaris && <th>Baris</th>}
                                  {orderedBookmarkHeaders.map(h => (
                                    <th key={h}>{getColumnLabel(h)}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {bookmarksList.map(match => (
                                  <tr key={match.id}>
                                    <td className="meta-cell text-center" style={{ width: '40px' }}>
                                      <button
                                        className="bookmark-star-btn active"
                                        onClick={() => handleToggleBookmark(match.id, true)}
                                        title="Hapus dari Bookmark"
                                      >
                                        ★
                                      </button>
                                    </td>
                                    {showFile && <td className="meta-cell font-accent">{match.filename}</td>}
                                    {showSheet && <td className="meta-cell">{match.sheet_name}</td>}
                                    {showBaris && <td className="meta-cell text-center">{match.row_number}</td>}
                                    {orderedBookmarkHeaders.map(h => {
                                      const isDesc = h === 'Kolom_17' || h === 'Kolom_18' || h.toLowerCase().includes('perihal') || h.toLowerCase().includes('uraian');
                                      return (
                                        <td key={h} className={isDesc ? 'description-cell' : ''}>
                                          {match.row_data[h] ?? ''}
                                        </td>
                                      );
                                    })}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </>
                      )}
                    </>
                  );
                })()}
              </div>
            )}
          </div>
        )}

        {/* TAB 2: UPLOAD */}
        {activeTab === 'upload' && (
          <div className="upload-tab-content">
            <div className="upload-container">
              <h2>Unggah Dokumen Excel Baru</h2>
              <p className="upload-description">
                Unggah spreadsheet Excel Anda (`.xlsx` atau `.xls`). Sistem akan membaca dan mengindeks setiap baris secara dinamis ke dalam database agar siap dicari.
              </p>

              {/* Drag and Drop Zone */}
              <div
                className={`dropzone ${dragActive ? 'active' : ''} ${uploading ? 'disabled' : ''}`}
                onDragEnter={handleDrag}
                onDragOver={handleDrag}
                onDragLeave={handleDrag}
                onDrop={handleDrop}
                onClick={() => !uploading && fileInputRef.current.click()}
              >
                <input
                  ref={fileInputRef}
                  type="file"
                  className="file-input"
                  accept=".xlsx, .xls"
                  onChange={handleFileChange}
                  disabled={uploading}
                />
                <div className="dropzone-content">
                  <div className="dropzone-icon">📥</div>
                  {uploading ? (
                    <div className="uploading-state">
                      <p className="upload-status-text">{uploadStatusMessage}</p>
                      <div className="progress-bar-container">
                        <div className="progress-bar" style={{ width: `${uploadProgress}%` }}></div>
                      </div>
                      <span className="progress-text">{uploadProgress}%</span>
                    </div>
                  ) : (
                    <>
                      <p className="main-drop-text">Seret & taruh file Excel di sini, atau <strong>klik untuk memilih</strong></p>
                      <p className="sub-drop-text">Mendukung format file .xlsx dan .xls (Maksimal 200MB)</p>
                    </>
                  )}
                </div>
              </div>

              {/* Success Notification */}
              {uploadSuccess && (
                <div className="notification success">
                  <div className="notification-icon">✅</div>
                  <div className="notification-body">
                    <h4>Impor Berhasil!</h4>
                    <p>File "{uploadSuccess.filename}" telah berhasil diproses di database.</p>
                    <div className="success-details">
                      <span>Sheet diproses: <strong>{uploadSuccess.sheetsProcessed}</strong></span>
                      <span>Baris diimpor: <strong>{uploadSuccess.rowsInserted.toLocaleString('id-ID')}</strong></span>
                    </div>
                  </div>
                </div>
              )}

              {/* Error Notification */}
              {uploadError && (
                <div className="notification error">
                  <div className="notification-icon">⚠️</div>
                  <div className="notification-body">
                    <h4>Impor Gagal</h4>
                    <p>{uploadError}</p>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* TAB 3: MANAGE FILES */}
        {activeTab === 'files' && (
          <div className="files-tab-content">
            <div className="files-header">
              <h2>Dokumen Terdaftar</h2>
              <p>Kelola file spreadsheet yang datanya saat ini aktif di dalam mesin pencari.</p>
            </div>

            {loadingFiles ? (
              <div className="search-loader">
                <div className="spinner"></div>
                <p>Mengambil daftar file...</p>
              </div>
            ) : filesList.length === 0 ? (
              <div className="empty-state">
                <div className="empty-icon">📁</div>
                <h3>Belum ada file terdaftar</h3>
                <p>Silakan buka tab <strong>Unggah Excel</strong> untuk mengimpor file pertama Anda.</p>
              </div>
            ) : (
              <>
                {/* Bulk Actions Header */}
                <div className="bulk-actions-header">
                  <label className="select-all-label">
                    <input
                      type="checkbox"
                      checked={selectedFileIds.length === filesList.length && filesList.length > 0}
                      onChange={(e) => {
                        if (e.target.checked) {
                          setSelectedFileIds(filesList.map(f => f.id));
                        } else {
                          setSelectedFileIds([]);
                        }
                      }}
                    />
                    <span>Pilih Semua ({filesList.length})</span>
                  </label>
                  
                  {selectedFileIds.length > 0 && (
                    <button
                      className="bulk-delete-btn"
                      onClick={handleBulkDelete}
                    >
                      🗑️ Hapus Terpilih ({selectedFileIds.length})
                    </button>
                  )}
                </div>

                <div className="files-grid">
                  {filesList.map(file => (
                    <div key={file.id} className={`file-card ${selectedFileIds.includes(file.id) ? 'selected' : ''}`}>
                      <div className="file-card-checkbox-wrapper">
                        <input
                          type="checkbox"
                          className="file-select-checkbox"
                          checked={selectedFileIds.includes(file.id)}
                          onChange={(e) => {
                            if (e.target.checked) {
                              setSelectedFileIds(prev => [...prev, file.id]);
                            } else {
                              setSelectedFileIds(prev => prev.filter(id => id !== file.id));
                            }
                          }}
                        />
                      </div>
                      <div className="file-card-details">
                        <h3 className="file-card-title" title={file.filename}>📄 {file.filename}</h3>
                        <div className="file-card-meta">
                          <span className="meta-tag blue">⚡ {file.row_count} baris</span>
                          <span className="meta-tag purple">📁 {file.sheet_names.length} sheet</span>
                        </div>
                        <p className="file-card-sheets">
                          <strong>Sheets:</strong> {file.sheet_names.join(', ')}
                        </p>
                        <p className="file-card-date">
                          Diunggah: {new Date(file.uploaded_at).toLocaleString('id-ID')}
                        </p>
                      </div>
                      <button 
                        className="delete-file-btn" 
                        onClick={() => handleDeleteFile(file.id, file.filename)}
                        title="Hapus file dan seluruh datanya"
                      >
                        🗑️ Hapus
                      </button>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        )}

        {/* TAB 4: ACTIVITY LOGS */}
        {activeTab === 'logs' && (
          <div className="logs-tab-content">
            {/* Header row */}
            <div className="section-header-row">
              <div>
                <h2>📜 Log Aktivitas Sistem</h2>
                <p className="subtitle">Riwayat audit pengunggahan, pencarian, dan penghapusan berkas.</p>
              </div>
              <button className="btn btn-secondary" onClick={() => fetchLogs(logSearch, logPage, logLimit)} disabled={loadingLogs}>
                {loadingLogs ? 'Memuat...' : '🔄 Segarkan'}
              </button>
            </div>

            {/* Search + Per-page controls */}
            <div className="logs-controls">
              <div className="logs-search-wrapper">
                <span className="logs-search-icon">🔍</span>
                <input
                  type="text"
                  className="logs-search-input"
                  placeholder="Cari aktivitas, kata kunci, tipe..."
                  value={logSearch}
                  onChange={e => setLogSearch(e.target.value)}
                />
                {logSearch && (
                  <button className="logs-search-clear" onClick={() => setLogSearch('')}>✕</button>
                )}
              </div>
              <div className="logs-per-page">
                <label>Tampilkan</label>
                <select
                  value={logLimit}
                  onChange={e => { setLogLimit(Number(e.target.value)); setLogPage(1); }}
                >
                  <option value={10}>10</option>
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value={200}>200</option>
                  <option value={500}>500</option>
                </select>
                <label>data per halaman</label>
              </div>
            </div>

            {/* Info bar */}
            {!loadingLogs && !logsError && logs.length > 0 && (
              <div className="logs-info-bar">
                Menampilkan <strong>{((logPage - 1) * logLimit) + 1}–{Math.min(logPage * logLimit, logTotal)}</strong> dari <strong>{logTotal.toLocaleString('id-ID')}</strong> aktivitas
                {logSearch && <span> · Filter: "<em>{logSearch}</em>"</span>}
              </div>
            )}

            {logsError && (
              <div className="error-status">⚠️ Error: {logsError}</div>
            )}

            {loadingLogs ? (
              <div className="search-loader">
                <div className="spinner"></div>
                <p>Memuat riwayat log...</p>
              </div>
            ) : logs.length === 0 ? (
              <div className="empty-state">
                <div className="empty-icon">📜</div>
                <h3>{logSearch ? 'Tidak ada log yang cocok' : 'Belum ada log aktivitas'}</h3>
                <p>{logSearch ? `Tidak ditemukan aktivitas dengan kata kunci "${logSearch}".` : 'Aktivitas sistem Anda akan tercatat di sini.'}</p>
              </div>
            ) : (
              <>
                <div className="table-responsive">
                  <table className="excel-table logs-table">
                    <thead>
                      <tr>
                        <th>Waktu Kejadian</th>
                        <th>Tipe</th>
                        <th>Detail Aktivitas</th>
                        <th>IP Address</th>
                        <th>Perangkat & OS</th>
                        <th>Browser</th>
                      </tr>
                    </thead>
                    <tbody>
                      {logs.map(log => {
                        let tagClass = 'tag-secondary';
                        if (log.activity_type === 'upload') tagClass = 'tag-success';
                        if (log.activity_type === 'delete') tagClass = 'tag-danger';
                        if (log.activity_type === 'search') tagClass = 'tag-info';
                        if (log.activity_type === 'access') tagClass = 'tag-access';

                        // Device type icon
                        const deviceType = (log.device_type || '').toLowerCase();
                        let deviceIcon = '🖥️';
                        if (deviceType === 'mobile') deviceIcon = '📱';
                        else if (deviceType === 'tablet') deviceIcon = '📟';
                        else if (deviceType === 'smarttv') deviceIcon = '📺';

                        // OS icon
                        const osName = (log.os || '').toLowerCase();
                        let osIcon = '💻';
                        if (osName.includes('windows')) osIcon = '🪟';
                        else if (osName.includes('mac') || osName.includes('ios')) osIcon = '🍎';
                        else if (osName.includes('android')) osIcon = '🤖';
                        else if (osName.includes('linux')) osIcon = '🐧';

                        // Browser icon
                        const browserName = (log.browser || '').toLowerCase();
                        let browserIcon = '🌐';
                        if (browserName.includes('chrome')) browserIcon = '🟡';
                        else if (browserName.includes('firefox')) browserIcon = '🦊';
                        else if (browserName.includes('safari')) browserIcon = '🧭';
                        else if (browserName.includes('edge')) browserIcon = '🔷';
                        else if (browserName.includes('opera')) browserIcon = '🔴';

                        return (
                          <tr key={log.id}>
                            <td className="meta-cell text-center" style={{ width: '160px', fontSize: '12px' }}>
                              {new Date(log.created_at).toLocaleString('id-ID')}
                            </td>
                            <td className="text-center" style={{ width: '100px' }}>
                              <span className={`activity-tag ${tagClass}`}>
                                {log.activity_type.toUpperCase()}
                              </span>
                            </td>
                            <td className="description-cell" style={{ fontSize: '13px' }}>
                              {log.activity_details}
                            </td>
                            <td className="meta-cell text-center" style={{ width: '140px' }}>
                              {log.ip_address ? (
                                <span className="device-badge ip-badge" title={log.ip_address}>
                                  🌐 {log.ip_address}
                                </span>
                              ) : <span className="no-data">—</span>}
                            </td>
                            <td className="meta-cell" style={{ width: '190px' }}>
                              {log.os ? (
                                <div className="device-info-cell">
                                  <span className="device-badge os-badge" title={log.os}>
                                    {osIcon} {log.os}
                                  </span>
                                  {log.device_label && log.device_label !== log.device_type && (
                                    <span className="device-badge device-type-badge" title={log.device_label}>
                                      {deviceIcon} {log.device_label}
                                    </span>
                                  )}
                                  {(!log.device_label || log.device_label === log.device_type) && log.device_type && (
                                    <span className="device-badge device-type-badge">
                                      {deviceIcon} {log.device_type}
                                    </span>
                                  )}
                                </div>
                              ) : <span className="no-data">—</span>}
                            </td>
                            <td className="meta-cell" style={{ width: '180px' }}>
                              {log.browser ? (
                                <span className="device-badge browser-badge" title={log.user_agent || log.browser}>
                                  {browserIcon} {log.browser}
                                </span>
                              ) : <span className="no-data">—</span>}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                {/* Pagination controls */}
                {logTotalPages > 1 && (
                  <div className="logs-pagination">
                    <button
                      className="page-btn"
                      onClick={() => setLogPage(1)}
                      disabled={logPage === 1}
                    >«</button>
                    <button
                      className="page-btn"
                      onClick={() => setLogPage(p => Math.max(1, p - 1))}
                      disabled={logPage === 1}
                    >‹ Sebelumnya</button>

                    {/* Page number pills */}
                    {Array.from({ length: logTotalPages }, (_, i) => i + 1)
                      .filter(p => p === 1 || p === logTotalPages || Math.abs(p - logPage) <= 2)
                      .reduce((acc, p, idx, arr) => {
                        if (idx > 0 && p - arr[idx - 1] > 1) acc.push('...');
                        acc.push(p);
                        return acc;
                      }, [])
                      .map((p, idx) =>
                        p === '...' ? (
                          <span key={`dots-${idx}`} className="page-dots">…</span>
                        ) : (
                          <button
                            key={p}
                            className={`page-btn ${logPage === p ? 'active' : ''}`}
                            onClick={() => setLogPage(p)}
                          >{p}</button>
                        )
                      )
                    }

                    <button
                      className="page-btn"
                      onClick={() => setLogPage(p => Math.min(logTotalPages, p + 1))}
                      disabled={logPage === logTotalPages}
                    >Berikutnya ›</button>
                    <button
                      className="page-btn"
                      onClick={() => setLogPage(logTotalPages)}
                      disabled={logPage === logTotalPages}
                    >»</button>
                  </div>
                )}
              </>
            )}
          </div>
        )}

      </main>

      {/* TAB 5: DASHBOARD */}
      {activeTab === 'dashboard' && (
        <main className="main-content">
          <div className="dashboard-header">
            <h2>📊 Dashboard Statistik</h2>
            <p>Ringkasan aktivitas dan kondisi data aplikasi secara real-time.</p>
            <button className="btn-refresh" onClick={fetchDashboard} disabled={loadingDashboard}>
              {loadingDashboard ? '⏳ Memuat...' : '🔄 Refresh'}
            </button>
          </div>

          {loadingDashboard ? (
            <div className="search-loader"><div className="spinner"></div><p>Memuat statistik...</p></div>
          ) : !dashboardStats ? (
            <div className="empty-state">
              <div className="empty-icon">📊</div>
              <h3>Belum ada data</h3>
              <p>Klik tombol Refresh untuk memuat statistik dashboard.</p>
            </div>
          ) : (
            <div className="dashboard-grid">

              {/* Summary Cards */}
              <div className="stat-card-grid">
                {[{
                  icon: '📁', label: 'Total File', value: parseInt(dashboardStats.summary?.total_files || 0).toLocaleString('id-ID'), color: 'blue'
                }, {
                  icon: '⚡', label: 'Total Baris Data', value: parseInt(dashboardStats.summary?.total_rows || 0).toLocaleString('id-ID'), color: 'purple'
                }, {
                  icon: '🔍', label: 'Pencarian Hari Ini', value: (dashboardStats.todayActivity?.find(a => a.activity_type === 'search')?.count || 0), color: 'teal'
                }, {
                  icon: '💻', label: 'Perangkat Aktif Hari Ini', value: dashboardStats.devicesOnline?.length || 0, color: 'orange'
                }].map((card, i) => (
                  <div key={i} className={`stat-card stat-card--${card.color}`}>
                    <div className="stat-card-icon">{card.icon}</div>
                    <div className="stat-card-value">{card.value}</div>
                    <div className="stat-card-label">{card.label}</div>
                  </div>
                ))}
              </div>

              {/* Top Searches */}
              <div className="dashboard-panel">
                <h3>🔥 Kata Kunci Terpopuler (30 Hari)</h3>
                {dashboardStats.topSearches?.length === 0 ? (
                  <p className="no-data-text">Belum ada aktivitas pencarian.</p>
                ) : (
                  <div className="top-searches-list">
                    {dashboardStats.topSearches?.map((s, i) => {
                      const detail = s.activity_details || '';
                      const match = detail.match(/"([^"]+)"/);
                      const keyword = match ? match[1] : detail;
                      const maxCount = dashboardStats.topSearches[0]?.count || 1;
                      const pct = Math.round((s.count / maxCount) * 100);
                      return (
                        <div key={i} className="search-keyword-row">
                          <span className="keyword-rank">#{i + 1}</span>
                          <div className="keyword-bar-wrap">
                            <div className="keyword-label" title={keyword}>{keyword}</div>
                            <div className="keyword-bar-bg">
                              <div className="keyword-bar-fill" style={{ width: `${pct}%` }}></div>
                            </div>
                          </div>
                          <span className="keyword-count">{s.count}x</span>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Docs per Unit */}
              <div className="dashboard-panel">
                <h3>🏦 Dokumen per Unit (Top 15)</h3>
                {dashboardStats.docsPerUnit?.length === 0 ? (
                  <p className="no-data-text">Tidak ada data unit.</p>
                ) : (
                  <div className="top-searches-list">
                    {dashboardStats.docsPerUnit?.map((u, i) => {
                      const maxCount = dashboardStats.docsPerUnit[0]?.count || 1;
                      const pct = Math.round((u.count / maxCount) * 100);
                      return (
                        <div key={i} className="search-keyword-row">
                          <span className="keyword-rank">#{i + 1}</span>
                          <div className="keyword-bar-wrap">
                            <div className="keyword-label">{u.unit || '(Kosong)'}</div>
                            <div className="keyword-bar-bg">
                              <div className="keyword-bar-fill keyword-bar-fill--green" style={{ width: `${pct}%` }}></div>
                            </div>
                          </div>
                          <span className="keyword-count">{parseInt(u.count).toLocaleString('id-ID')}</span>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Today's Activity Breakdown */}
              <div className="dashboard-panel">
                <h3>📅 Aktivitas Hari Ini</h3>
                <div className="today-activity-grid">
                  {['access', 'search', 'upload', 'delete'].map(type => {
                    const found = dashboardStats.todayActivity?.find(a => a.activity_type === type);
                    const icons = { access: '📲', search: '🔍', upload: '📤', delete: '🗑️' };
                    const labels = { access: 'Akses', search: 'Pencarian', upload: 'Upload', delete: 'Hapus' };
                    return (
                      <div key={type} className="today-act-card">
                        <div className="today-act-icon">{icons[type]}</div>
                        <div className="today-act-count">{found?.count || 0}</div>
                        <div className="today-act-label">{labels[type]}</div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Devices Online Today */}
              <div className="dashboard-panel dashboard-panel--wide">
                <h3>💻 Perangkat Aktif Hari Ini</h3>
                {dashboardStats.devicesOnline?.length === 0 ? (
                  <p className="no-data-text">Belum ada perangkat yang mengakses hari ini.</p>
                ) : (
                  <div className="devices-grid">
                    {dashboardStats.devicesOnline?.map((d, i) => {
                      const osName = (d.os || '').toLowerCase();
                      let osIcon = '💻';
                      if (osName.includes('windows')) osIcon = '🪟';
                      else if (osName.includes('mac') || osName.includes('ios')) osIcon = '🍎';
                      else if (osName.includes('android')) osIcon = '🤖';
                      else if (osName.includes('linux')) osIcon = '🐧';
                      return (
                        <div key={i} className="device-card">
                          <div className="device-card-icon">{osIcon}</div>
                          <div className="device-card-info">
                            <div className="device-card-ip">{d.ip_address}</div>
                            <div className="device-card-os">{d.os || 'Unknown OS'}</div>
                            <div className="device-card-browser">{d.browser || 'Unknown Browser'}</div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

            </div>
          )}
        </main>
      )}
      
      <footer className="app-footer-bar">
        <p>© 2026 SpreadSheet Finder — Didukung oleh React, Express, dan PostgreSQL</p>
      </footer>

      {/* Toast Notification Container */}
      <div className="toast-container">
        {toasts.map(toast => {
          const icons = { success: '✅', error: '⚠️', info: '🔔', upload: '📦' };
          return (
            <div
              key={toast.id}
              className={`toast toast-${toast.type} ${toast.exiting ? 'toast-exit' : ''}`}
              onClick={() => dismissToast(toast.id)}
            >
              <div className="toast-icon">{icons[toast.type] || '🔔'}</div>
              <div className="toast-body">
                <div className="toast-title">{toast.title}</div>
                <div className="toast-message">{toast.message}</div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

export default App;
