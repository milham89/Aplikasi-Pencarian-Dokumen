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
  const [token, setToken] = useState(() => localStorage.getItem('app_token'));
  const [user, setUser] = useState(() => {
    const saved = localStorage.getItem('app_user');
    return saved ? JSON.parse(saved) : null;
  });

  // Navigation Tabs: 'search' | 'upload' | 'files' | 'logs' | 'dashboard' | 'bookmarks' | 'users'
  const [activeTab, setActiveTab] = useState('search');
  
  // Dark/Light Theme
  const [isDarkMode, setIsDarkMode] = useState(() => {
    const saved = localStorage.getItem('app_theme');
    return saved ? saved === 'dark' : true;
  });

  // Toast notifications
  const [toasts, setToasts] = useState([]);

  // Live Clock & Time State
  const [currentTime, setCurrentTime] = useState(new Date());

  useEffect(() => {
    const timer = setInterval(() => {
      setCurrentTime(new Date());
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  const getGreeting = () => {
    const hour = currentTime.getHours();
    if (hour >= 5 && hour < 11) return '☀️ Selamat Pagi';
    if (hour >= 11 && hour < 15) return '🌤️ Selamat Siang';
    if (hour >= 15 && hour < 19) return '⛅ Selamat Sore';
    return '🌙 Selamat Malam';
  };

  // Dashboard stats
  const [dashboardStats, setDashboardStats] = useState(null);
  const [loadingDashboard, setLoadingDashboard] = useState(false);
  
  // Database status
  const [dbConnected, setDbConnected] = useState(false);
  const [stats, setStats] = useState({ totalFiles: 0, totalRows: 0 });
  
  // Search state
  const [searchQuery, setSearchQuery] = useState('');
  const [searchResults, setSearchResults] = useState([]);
  const [searchStats, setSearchStats] = useState({ totalMatches: 0, timeMs: 0 });
  const [searchUIState, setSearchUIState] = useState('loading'); // 'welcome' | 'preview' | 'loading' | 'results' | 'empty'
  const [searchProgress, setSearchProgress] = useState(0); // 1 - 100% progress
  const [selectedFileId, setSelectedFileId] = useState(null); // Active file tab in results

  // Manage smooth 1 to 100% progress animation during loading state
  useEffect(() => {
    let timer;
    if (searchUIState === 'loading') {
      setSearchProgress(10);
      timer = setInterval(() => {
        setSearchProgress(prev => {
          if (prev >= 92) {
            clearInterval(timer);
            return prev;
          }
          const step = prev < 40 ? 15 : prev < 75 ? 8 : prev < 88 ? 3 : 1;
          return Math.min(92, prev + step);
        });
      }, 90);
    } else {
      setSearchProgress(100);
    }
    return () => clearInterval(timer);
  }, [searchUIState]);
  
  // Advanced filters state
  const [filterSheet, setFilterSheet] = useState('');
  const [filterUnit, setFilterUnit] = useState('');
  const [searchUploaderFilter, setSearchUploaderFilter] = useState('all');
  const [showAdvancedFilters, setShowAdvancedFilters] = useState(false);
  const [enableAISearch, setEnableAISearch] = useState(true);
  const [aiSearchInfo, setAiSearchInfo] = useState({ used: false, method: 'none' });
  const [editingCell, setEditingCell] = useState(null); // { rowId, headerKey }
  const [editCellValue, setEditCellValue] = useState('');
  const [isSavingCell, setIsSavingCell] = useState(false);

  // AI Chatbot State
  const [isChatOpen, setIsChatOpen] = useState(false);
  const [chatMessages, setChatMessages] = useState([
    {
      id: 'welcome',
      sender: 'ai',
      text: 'Halo! Saya **SpreadSheet AI Assistant** 🤖. Saya siap membantu Anda menganalisis, merangkum, dan menjawab pertanyaan seputar berkas dokumen arsip di database.',
      timestamp: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }),
      sources: []
    }
  ]);
  const [chatInput, setChatInput] = useState('');
  const [isChatLoading, setIsChatLoading] = useState(false);
  const [chatSelectedFileId, setChatSelectedFileId] = useState('all');
  const [chatModelMode, setChatModelMode] = useState('auto'); // 'auto' | 'gemini' | 'local'
  const [highlightedRowNumber, setHighlightedRowNumber] = useState(null);

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
  const skipDebouncedSearchRef = useRef(false);
  const isManualTypingRef = useRef(false);
  const [copiedLogId, setCopiedLogId] = useState(null);

  const handleCopyLog = (log) => {
    const textToCopy = `[${new Date(log.created_at).toLocaleString('id-ID')}] Account: ${log.username || '-'} | Type: ${log.activity_type} | Detail: ${log.activity_details} | IP: ${log.ip_address || '-'} | OS: ${log.os || '-'} | Browser: ${log.browser || '-'}`;
    navigator.clipboard.writeText(textToCopy);
    setCopiedLogId(log.id);
    setTimeout(() => setCopiedLogId(null), 2000);
  };
  
  // Auto-scroll to AI cited row highlight when clicking reference chips
  useEffect(() => {
    if (highlightedRowNumber && (searchUIState === 'preview' || searchUIState === 'results')) {
      const timer = setTimeout(() => {
        const el = document.getElementById(`row-cited-${highlightedRowNumber}`);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }, 350);
      return () => clearTimeout(timer);
    }
  }, [highlightedRowNumber, searchUIState, searchResults]);
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

  // Deleting state
  const [deleting, setDeleting] = useState(false);
  const [deleteProgress, setDeleteProgress] = useState(0);
  const [deleteStatusMessage, setDeleteStatusMessage] = useState('');

  // Download progress state
  const [downloading, setDownloading] = useState(false);
  const downloadingRef = useRef(false);
  const [downloadProgress, setDownloadProgress] = useState(0);
  const [downloadStatusMessage, setDownloadStatusMessage] = useState('');
  const [downloadingFileId, setDownloadingFileId] = useState(null);

  // Check backend & DB status
  const healthFailCountRef = useRef(0);
  const checkHealth = async () => {
    try {
      const res = await apiFetch('/api/health');
      const data = await res.json();
      if (data.status === 'ok' && data.database === 'connected') {
        healthFailCountRef.current = 0;
        setDbConnected(true);
      } else {
        if (!downloadingRef.current) {
          healthFailCountRef.current += 1;
          if (healthFailCountRef.current >= 3) {
            setDbConnected(false);
          }
        }
      }
    } catch (err) {
      if (!downloadingRef.current) {
        healthFailCountRef.current += 1;
        if (healthFailCountRef.current >= 3) {
          setDbConnected(false);
        }
      }
    }
  };

  // User Management state
  const [usersList, setUsersList] = useState([]);
  const [loadingUsers, setLoadingUsers] = useState(false);
  const [newUsername, setNewUsername] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [newFullName, setNewFullName] = useState('');
  const [newUserRole, setNewUserRole] = useState('viewer');
  const [editingUser, setEditingUser] = useState(null); // holds user object when editing
  const [selectedUserForBookmarks, setSelectedUserForBookmarks] = useState(null);
  const [userBookmarksList, setUserBookmarksList] = useState([]);
  const [loadingUserBookmarks, setLoadingUserBookmarks] = useState(false);
  const [editingBookmark, setEditingBookmark] = useState(null);
  const [editingBookmarkNotes, setEditingBookmarkNotes] = useState('');
  const [editingBookmarkGroup, setEditingBookmarkGroup] = useState('Umum');
  const [editingBookmarkRowData, setEditingBookmarkRowData] = useState({});
  const [notificationsList, setNotificationsList] = useState([]);
  const [showNotificationsDropdown, setShowNotificationsDropdown] = useState(false);
  const [broadcastMessage, setBroadcastMessage] = useState('');
  const [broadcastRecipient, setBroadcastRecipient] = useState('all');
  const [selectedUserFilterForBookmarks, setSelectedUserFilterForBookmarks] = useState('mine');
  const [selectedBookmarkUserIds, setSelectedBookmarkUserIds] = useState([]);
  const [showUserFilterDropdown, setShowUserFilterDropdown] = useState(false);
  const [activeBookmarkGroupFilter, setActiveBookmarkGroupFilter] = useState('Semua');
  const [showBookmarkGroupModal, setShowBookmarkGroupModal] = useState(false);
  const [bookmarkGroupModalRowId, setBookmarkGroupModalRowId] = useState(null);
  const [bookmarkGroupModalCurrentStatus, setBookmarkGroupModalCurrentStatus] = useState(false);
  const [bookmarkGroupModalTargetUserId, setBookmarkGroupModalTargetUserId] = useState(null);
  const [bookmarkGroupModalSelectedOption, setBookmarkGroupModalSelectedOption] = useState('existing'); // 'existing' or 'new'
  const [bookmarkGroupModalSelectedExisting, setBookmarkGroupModalSelectedExisting] = useState('Umum');
  const [bookmarkGroupModalNewInput, setBookmarkGroupModalNewInput] = useState('');
  const fileInputRef = useRef(null);
  const searchInputRef = useRef(null);
  const searchAbortControllerRef = useRef(null);



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

  // Set default selected bookmark user ID to current user when logged in or usersList is fetched
  useEffect(() => {
    if (user && selectedBookmarkUserIds.length === 0) {
      const currentUserFromList = usersList.find(u => u.username === user.username);
      const currentUserId = currentUserFromList ? currentUserFromList.id : user.id;
      if (currentUserId) {
        setSelectedBookmarkUserIds([currentUserId]);
      }
    }
  }, [user, usersList]);

  // Set initial preview state when files list is first loaded
  useEffect(() => {
    if (searchQuery.trim() === '' && filesList.length > 0 && !selectedFileId) {
      setSearchUIState('preview');
    }
  }, [filesList]);

  // Fetch bookmarks when bookmarks tab is active or selected user filter/IDs changes
  useEffect(() => {
    if (activeTab === 'bookmarks') {
      fetchBookmarks();
    }
  }, [activeTab, selectedUserFilterForBookmarks, selectedBookmarkUserIds]);

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

    // If search query is empty, restore document preview mode without forcing filesList[0]
    if (searchQuery.trim() === '') {
      skipDebouncedSearchRef.current = false;
      isManualTypingRef.current = false;
      if (filesList.length > 0) {
        if (selectedFileId && filesList.some(f => String(f.id) === String(selectedFileId))) {
          fetchFilePreview(selectedFileId, filterSheet, filterUnit, null, '');
        } else {
          setSearchResults([]);
          setSelectedFileId(null);
          setSearchUIState('preview');
        }
      } else {
        setSearchResults([]);
        setSelectedFileId(null);
        setSearchUIState('welcome');
      }
      return;
    }

    if (skipDebouncedSearchRef.current || !isManualTypingRef.current) {
      skipDebouncedSearchRef.current = false;
      return;
    }

    setSearchUIState('loading');

    const timer = setTimeout(() => {
      isManualTypingRef.current = false;
      handleSearch();
    }, 600);

    return () => clearTimeout(timer);
  }, [searchQuery, filterSheet, filterUnit, filesList]);

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

  const handleLogout = useCallback(async () => {
    // Notify backend for audit log (fire-and-forget, don't wait)
    try {
      const currentToken = localStorage.getItem('app_token');
      if (currentToken) {
        fetch('/api/auth/logout', {
          method: 'POST',
          headers: { 'Authorization': `Bearer ${currentToken}` }
        }).catch(() => {});
      }
    } catch (_) {}

    localStorage.removeItem('app_token');
    localStorage.removeItem('app_user');
    setToken(null);
    setUser(null);
    setActiveTab('search');
    showToast('🔑 Logout Sukses', 'Sesi Anda telah diakhiri.', 'info');
  }, [showToast]);

  const apiFetch = useCallback(async (url, options = {}) => {
    const headers = {
      ...options.headers,
    };
    const currentToken = localStorage.getItem('app_token');
    if (currentToken) {
      headers['Authorization'] = `Bearer ${currentToken}`;
    }
    
    try {
      const res = await fetch(url, {
        ...options,
        headers,
      });

      if (res.status === 401 || res.status === 403) {
        if (!url.includes('/api/health') && !url.includes('/api/auth/login')) {
          handleLogout();
        }
      }
      return res;
    } catch (err) {
      console.error(`API Fetch Error (${url}):`, err);
      throw err;
    }
  }, [handleLogout]);

  const fetchUsers = useCallback(async () => {
    setLoadingUsers(true);
    try {
      const res = await apiFetch('/api/users');
      if (res.ok) {
        const data = await res.json();
        setUsersList(data);
      }
    } catch (err) {
      console.error('Gagal mengambil daftar pengguna:', err);
    } finally {
      setLoadingUsers(false);
    }
  }, [apiFetch]);

  const handleCreateUser = async (e) => {
    e.preventDefault();
    if (!newUsername.trim() || !newPassword) {
      showToast('⚠️ Peringatan', 'Username dan Password wajib diisi.', 'warning');
      return;
    }
    try {
      const res = await apiFetch('/api/users', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: newUsername.trim(), password: newPassword, role: newUserRole, full_name: newFullName.trim() || undefined })
      });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Gagal membuat pengguna.');
      }
      showToast('👤 Akun Dibuat', `Akun "${data.username}"${data.full_name ? ` (${data.full_name})` : ''} berhasil didaftarkan.`, 'success');
      setNewUsername('');
      setNewPassword('');
      setNewFullName('');
      setNewUserRole('viewer');
      fetchUsers();
    } catch (err) {
      showToast('⚠️ Gagal', err.message, 'error');
    }
  };

  const handleUpdateUser = async (e) => {
    e.preventDefault();
    if (!editingUser) return;
    if (!newUsername.trim()) {
      showToast('⚠️ Peringatan', 'Username wajib diisi.', 'warning');
      return;
    }
    try {
      const res = await apiFetch(`/api/users/${editingUser.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username: newUsername.trim(),
          password: newPassword ? newPassword : undefined,
          role: newUserRole,
          full_name: newFullName.trim() || null
        })
      });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Gagal memperbarui pengguna.');
      }
      showToast('👤 Akun Diperbarui', `Akun "${data.username}" berhasil diperbarui.`, 'success');
      setNewUsername('');
      setNewPassword('');
      setNewFullName('');
      setNewUserRole('viewer');
      setEditingUser(null);
      fetchUsers();
    } catch (err) {
      showToast('⚠️ Gagal', err.message, 'error');
    }
  };

  const startEditUser = (u) => {
    setEditingUser(u);
    setNewUsername(u.username);
    setNewFullName(u.full_name || '');
    setNewUserRole(u.role);
    setNewPassword('');
  };

  const cancelEditUser = () => {
    setEditingUser(null);
    setNewUsername('');
    setNewFullName('');
    setNewUserRole('viewer');
    setNewPassword('');
  };

  const handleDeleteUser = async (userId, username) => {
    if (!window.confirm(`Apakah Anda yakin ingin menghapus akun "${username}"?`)) {
      return;
    }
    try {
      const res = await apiFetch(`/api/users/${userId}`, {
        method: 'DELETE'
      });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Gagal menghapus pengguna.');
      }
      showToast('🗑️ Akun Dihapus', `Akun "${username}" berhasil dihapus.`, 'success');
      fetchUsers();
    } catch (err) {
      showToast('⚠️ Gagal', err.message, 'error');
    }
  };

  const fetchUserBookmarks = async (targetUser) => {
    setSelectedUserForBookmarks(targetUser);
    setLoadingUserBookmarks(true);
    try {
      const res = await apiFetch(`/api/bookmarks?userId=${targetUser.id}`);
      if (res.ok) {
        const data = await res.json();
        setUserBookmarksList(data.bookmarks || []);
      }
    } catch (err) {
      console.error('Gagal mengambil bookmark pengguna:', err);
    } finally {
      setLoadingUserBookmarks(false);
    }
  };

  const handleDeleteUserBookmark = async (rowId) => {
    if (!selectedUserForBookmarks) return;
    if (!window.confirm('Apakah Anda yakin ingin menghapus bookmark ini dari akun pengguna?')) return;
    try {
      const res = await apiFetch(`/api/bookmarks/${rowId}?userId=${selectedUserForBookmarks.id}`, {
        method: 'DELETE'
      });
      if (res.ok) {
        showToast('🗑️ Bookmark Dihapus', `Bookmark berhasil dihapus dari akun "${selectedUserForBookmarks.username}".`, 'success');
        // Refresh list
        const res2 = await apiFetch(`/api/bookmarks?userId=${selectedUserForBookmarks.id}`);
        if (res2.ok) {
          const data2 = await res2.json();
          setUserBookmarksList(data2.bookmarks || []);
        }
      }
    } catch (err) {
      showToast('⚠️ Gagal', 'Gagal menghapus bookmark.', 'error');
    }
  };

  const handleClearUserBookmarks = async () => {
    if (!selectedUserForBookmarks) return;
    if (!window.confirm(`Apakah Anda yakin ingin menghapus seluruh bookmark (${userBookmarksList.length} item) dari akun "${selectedUserForBookmarks.username}"?`)) return;
    try {
      const res = await apiFetch(`/api/bookmarks?userId=${selectedUserForBookmarks.id}`, {
        method: 'DELETE'
      });
      if (res.ok) {
        showToast('🗑️ Bookmark Dibersihkan', `Seluruh bookmark milik "${selectedUserForBookmarks.username}" berhasil dihapus.`, 'success');
        setUserBookmarksList([]);
        // Refresh local bookmarks list as well if it's the current user
        if (selectedUserForBookmarks.id === user.id) {
          fetchBookmarks();
        }
      }
    } catch (err) {
      showToast('⚠️ Gagal', 'Gagal membersihkan bookmark.', 'error');
    }
  };

  const handleUpdateBookmark = async (e) => {
    e.preventDefault();
    if (!editingBookmark) return;

    try {
      let targetUserId = user.id;
      if (selectedUserForBookmarks) {
        targetUserId = selectedUserForBookmarks.id;
      } else if (editingBookmark.owner_id) {
        targetUserId = editingBookmark.owner_id;
      }

      // 1. Update Notes of bookmark
      const bookmarkRes = await apiFetch(`/api/bookmarks/${editingBookmark.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ notes: editingBookmarkNotes, group_name: editingBookmarkGroup, userId: targetUserId })
      });

      if (!bookmarkRes.ok) {
        const errorData = await bookmarkRes.json();
        throw new Error(errorData.error || 'Gagal memperbarui catatan bookmark.');
      }

      // 2. Update Row Data of document
      const rowRes = await apiFetch(`/api/document-rows/${editingBookmark.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ row_data: editingBookmarkRowData })
      });

      if (!rowRes.ok) {
        const errorData = await rowRes.json();
        throw new Error(errorData.error || 'Gagal memperbarui data baris dokumen.');
      }

      showToast('✅ Berhasil Diperbarui', 'Bookmark dan data baris dokumen berhasil diperbarui.', 'success');
      setEditingBookmark(null);
      
      // Refresh lists
      if (selectedUserForBookmarks) {
        const res = await apiFetch(`/api/bookmarks?userId=${selectedUserForBookmarks.id}`);
        if (res.ok) {
          const data = await res.json();
          setUserBookmarksList(data.bookmarks || []);
        }
      }
      fetchBookmarks();
    } catch (err) {
      showToast('⚠️ Gagal', err.message, 'error');
    }
  };

  const fetchNotifications = useCallback(async () => {
    try {
      const res = await apiFetch('/api/notifications');
      if (res.ok) {
        const data = await res.json();
        setNotificationsList(data.notifications || []);
      }
    } catch (err) {
      console.error('Gagal memuat notifikasi:', err);
    }
  }, [apiFetch]);

  const handleBroadcastNotification = async (e) => {
    e.preventDefault();
    if (!broadcastMessage.trim()) {
      showToast('⚠️ Peringatan', 'Pesan notifikasi wajib diisi.', 'warning');
      return;
    }
    try {
      const res = await apiFetch('/api/notifications', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          message: broadcastMessage.trim(),
          recipientId: broadcastRecipient
        })
      });
      if (!res.ok) {
        const data = await res.json();
        throw new Error(data.error || 'Gagal mengirim notifikasi.');
      }
      
      const successMsg = broadcastRecipient === 'all' 
        ? 'Notifikasi berhasil dikirim ke semua akun.'
        : 'Notifikasi berhasil dikirim ke akun penerima yang dituju.';
        
      showToast('📢 Notifikasi Dikirim', successMsg, 'success');
      setBroadcastMessage('');
      setBroadcastRecipient('all');
      fetchNotifications();
    } catch (err) {
      showToast('⚠️ Gagal', err.message, 'error');
    }
  };

  const handleMarkAsRead = async (notifId = null) => {
    try {
      const res = await apiFetch('/api/notifications/read', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ notificationId: notifId })
      });
      if (res.ok) {
        fetchNotifications();
      }
    } catch (err) {
      console.error('Gagal menandai notifikasi dibaca:', err);
    }
  };

  useEffect(() => {
    if (token) {
      fetchNotifications();
      const interval = setInterval(() => {
        fetchNotifications();
      }, 30000);
      return () => clearInterval(interval);
    }
  }, [token]);

  useEffect(() => {
    if (token && user?.role === 'admin') {
      fetchUsers();
    }
  }, [token, user?.role]);

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
      const res = await apiFetch('/api/stats');
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

  // Reset dashboard statistics (Admin only)
  const handleResetStats = async () => {
    if (!window.confirm("Apakah Anda yakin ingin mereset seluruh data statistik dan riwayat log aktivitas? Tindakan ini tidak dapat dibatalkan.")) {
      return;
    }
    
    try {
      const res = await apiFetch('/api/stats/reset', {
        method: 'POST'
      });
      const data = await res.json();
      if (res.ok) {
        showToast('🗑️ Sukses Reset', data.message || 'Statistik berhasil direset.', 'success');
        fetchDashboard();
      } else {
        showToast('⚠️ Gagal', data.error || 'Gagal mereset statistik.', 'error');
      }
    } catch (err) {
      showToast('⚠️ Error', err.message, 'error');
    }
  };

  // Reset / clear system activity logs (Admin only)
  const handleClearLogs = async () => {
    if (!window.confirm("Apakah Anda yakin ingin menghapus/reset seluruh riwayat log aktivitas sistem? Tindakan ini tidak dapat dibatalkan.")) {
      return;
    }
    
    try {
      const res = await apiFetch('/api/logs', {
        method: 'DELETE'
      });
      const data = await res.json();
      if (res.ok) {
        showToast('🗑️ Sukses Reset', data.message || 'Log aktivitas berhasil direset.', 'success');
        setLogPage(1);
        fetchLogs(logSearch, 1, logLimit);
      } else {
        showToast('⚠️ Gagal', data.error || 'Gagal mereset log aktivitas.', 'error');
      }
    } catch (err) {
      showToast('⚠️ Error', err.message, 'error');
    }
  };

  // Get user profile details
  const fetchUserProfile = useCallback(async () => {
    try {
      const res = await apiFetch('/api/auth/me');
      if (res.ok) {
        const data = await res.json();
        setUser(data);
        localStorage.setItem('app_user', JSON.stringify(data));
      }
    } catch (err) {
      console.error('Gagal mengambil data profil pengguna:', err);
    }
  }, [apiFetch]);

  // Initial checks and loads - placed below helpers to avoid TDZ
  useEffect(() => {
    checkHealth();
    if (token) {
      fetchUserProfile();
      fetchFiles();
    }
  }, [token]);

  // Get list of uploaded files
  const fetchFiles = async () => {
    setLoadingFiles(true);
    try {
      const res = await apiFetch('/api/files');
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
      const res = await apiFetch(`/api/logs?${params}`);
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

  // Fetch File Preview with optional targetRowNumber highlight and query text filter
  const fetchFilePreview = async (fileId, sheet = null, unit = null, targetRowNumber = null, queryOverride = null, sheetName = null) => {
    setSearchUIState('loading');
    setSearchResults([]);
    setAiSearchInfo({ used: false, method: 'none' });
    try {
      const params = new URLSearchParams({ fileId: fileId.toString() });
      const activeQ = queryOverride !== null ? queryOverride : searchQuery;
      // Strip conversational words so backend gets only the keyword
      if (activeQ && activeQ.trim() !== '') {
        const stopWords = new Set([
          'tampilkan', 'data', 'pada', 'yang', 'dan', 'di', 'untuk', 'ke', 'dari', 
          'adalah', 'itu', 'dengan', 'ini', 'oleh', 'seperti', 'adapun', 'atau', 
          'sebagai', 'tentang', 'yaitu', 'ia', 'kami', 'mereka', 'saya', 'anda', 
          'dia', 'kita', 'tahun', 'bulan', 'hari', 'file', 'berkas', 'dokumen', 
          'tabel', 'sheet', 'cari', 'temukan', 'info', 'informasi', 'lihat', 'berapa',
          'apakah', 'siapa', 'bagaimana', 'apa', 'tolong', 'jelaskan', 'rangkum', 'semua',
          'daftar', 'total', 'ada', 'berurutan', 'entry', 'entri'
        ]);
        const rawTerms = activeQ.trim().split(/[\s\.]+/).filter(Boolean);
        const filtered = rawTerms.filter(t => !stopWords.has(t.toLowerCase()) && t.length >= 2);
        const cleanQ = (filtered.length > 0 ? filtered : rawTerms).join(' ');
        if (cleanQ.trim() !== '') params.append('q', cleanQ.trim());
      }
      if (sheet) params.append('sheet', sheet);
      if (sheetName) params.append('sheetName', sheetName);
      if (unit) params.append('unit', unit);
      if (targetRowNumber) params.append('rowNumber', targetRowNumber.toString());
      if (user?.role === 'admin' && searchUploaderFilter !== 'all') {
        params.append('uploaderId', searchUploaderFilter);
      }

      const res = await apiFetch(`/api/search?${params}`);
      if (res.ok) {
        const data = await res.json();
        setSearchResults(data.results);
        if (data.results.length > 0) {
          setSelectedFileId(fileId);
          setSearchUIState('preview');
          if (targetRowNumber) {
            setHighlightedRowNumber(parseInt(targetRowNumber, 10));
          } else {
            setHighlightedRowNumber(null);
          }
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
      if (enableAISearch) params.append('ai', 'true');
      if (user?.role === 'admin' && searchUploaderFilter !== 'all') {
        params.append('uploaderId', searchUploaderFilter);
      }

      const res = await apiFetch(`/api/search?${params}`, {
        signal: controller.signal
      });
      if (res.ok) {
        const data = await res.json();
        setSearchResults(data.results);
        setSearchStats({
          totalMatches: data.totalMatchesCount || 0,
          timeMs: data.executionTimeMs || 0
        });
        setAiSearchInfo({ used: !!data.aiSearchUsed, method: data.aiMethod || 'none' });
        
        // Save to search history
        addToHistory(searchQuery);

        // Pertahankan pilihan berkas jika pengguna sudah memilihnya sebelumnya dan masih ada di hasil.
        // Jika pengguna belum memilih berkas, biarkan selectedFileId = null agar pengguna memilih dokumen sendiri.
        if (data.results.length > 0) {
          const currentIdStillExists = selectedFileId && data.results.some(
            r => String(r.fileId) === String(selectedFileId)
          );

          if (currentIdStillExists) {
            // Refresh tabel untuk berkas yang sedang dipilih oleh pengguna
            fetchFilePreview(selectedFileId, filterSheet, filterUnit, null, searchQuery);
          } else {
            // Pengguna belum memilih / berkas sebelumnya tidak ada di hasil → minta pengguna memilih berkas
            setSelectedFileId(null);
          }
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

  // Handle file tab click - explicitly fetch preview rows for clicked fileId
  const handleFileTabClick = (fileId) => {
    skipDebouncedSearchRef.current = true;
    isManualTypingRef.current = false;
    setSelectedFileId(fileId);
    fetchFilePreview(fileId, filterSheet, filterUnit, null, searchQuery);
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
  const fetchBookmarks = async (targetUserId = null) => {
    setLoadingBookmarks(true);
    try {
      let userIdParam = '';
      if (targetUserId !== null) {
        userIdParam = `?userId=${targetUserId}`;
      } else if (user?.role === 'admin') {
        if (selectedBookmarkUserIds.length > 0) {
          userIdParam = `?userId=${selectedBookmarkUserIds.join(',')}`;
        } else {
          userIdParam = '?userId=mine';
        }
      }
      
      const res = await apiFetch(`/api/bookmarks${userIdParam}`);
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

  // Save inline cell edit to PostgreSQL database
  const handleSaveCell = async (rowId, headerKey, originalValue, currentMatches) => {
    // If the value hasn't changed, just close editing mode
    if (editCellValue.trim() === (originalValue || '').trim()) {
      setEditingCell(null);
      return;
    }

    setIsSavingCell(true);
    try {
      // Find the row matches object to copy other column values
      const matchingRow = currentMatches.find(m => m.id === rowId);
      if (!matchingRow) {
        throw new Error('Baris dokumen tidak ditemukan dalam state.');
      }

      // Prepare updated row_data payload
      const updatedRowData = {
        ...matchingRow.rowData,
        [headerKey]: editCellValue
      };

      const res = await apiFetch(`/api/document-rows/${rowId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ row_data: updatedRowData })
      });

      if (res.ok) {
        // Update local search results state immediately so the table updates visually
        setSearchResults(prevResults => {
          return prevResults.map(file => {
            const updatedMatches = file.matches.map(m => {
              if (m.id === rowId) {
                return { ...m, rowData: updatedRowData };
              }
              return m;
            });
            return { ...file, matches: updatedMatches };
          });
        });

        toast.success('Data sel berhasil diperbarui!');
      } else {
        const errData = await res.json();
        throw new Error(errData.error || 'Gagal menyimpan perubahan.');
      }
    } catch (err) {
      console.error('Error saving document row cell:', err);
      toast.error(`Gagal memperbarui data: ${err.message}`);
    } finally {
      setIsSavingCell(false);
      setEditingCell(null);
    }
  };

  const handleCancelEdit = () => {
    setEditingCell(null);
    setEditCellValue('');
  };

  // Handle AI Chatbot message sending
  const handleSendChatMessage = async (presetText = null) => {
    const textToSend = presetText || chatInput;
    if (!textToSend || !textToSend.trim() || isChatLoading) return;

    const userMessage = {
      id: crypto.randomUUID(),
      sender: 'user',
      text: textToSend.trim(),
      timestamp: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }),
      sources: []
    };

    setChatMessages(prev => [...prev, userMessage]);
    if (!presetText) setChatInput('');
    setIsChatLoading(true);

    const historyPayload = chatMessages
      .filter(m => m.id !== 'welcome')
      .slice(-4)
      .map(m => ({
        role: m.sender === 'user' ? 'user' : 'model',
        parts: [{ text: m.text }]
      }));

    try {
      const res = await apiFetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: userMessage.text,
          history: historyPayload,
          fileId: chatSelectedFileId,
          modelMode: chatModelMode
        })
      });

      if (res.ok) {
        const data = await res.json();
        const aiMessage = {
          id: crypto.randomUUID(),
          sender: 'ai',
          text: data.answer,
          timestamp: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }),
          sources: data.sources || [],
          methodUsed: data.methodUsed,
          userQuery: textToSend
        };
        setChatMessages(prev => [...prev, aiMessage]);
      } else {
        let errorMsg = `Server mengembalikan status ${res.status}`;
        try {
          const contentType = res.headers.get('content-type') || '';
          if (contentType.includes('application/json')) {
            const err = await res.json();
            errorMsg = err.error || err.message || errorMsg;
          } else {
            const text = await res.text();
            if (res.status === 504 || text.includes('504 Gateway') || text.includes('Timeout')) {
              errorMsg = 'Koneksi ke server mengalami batas waktu (Timeout 504). Silakan coba lagi.';
            } else if (res.status === 502 || text.includes('502 Bad Gateway')) {
              errorMsg = 'Server backend tidak dapat dijangkau (Bad Gateway 502).';
            } else {
              errorMsg = `Server mengembalikan error (Status ${res.status})`;
            }
          }
        } catch (_) {}
        throw new Error(errorMsg);
      }
    } catch (err) {
      console.error('Chat error:', err);
      const errorMessage = {
        id: crypto.randomUUID(),
        sender: 'ai',
        text: `Maaf, terjadi kesalahan saat menghubungi AI Chatbot: ${err.message}`,
        timestamp: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }),
        sources: []
      };
      setChatMessages(prev => [...prev, errorMessage]);
    } finally {
      setIsChatLoading(false);
    }
  };

  const handleSourceClick = (fileId, filename, rowNumber = null, userQuery = '', sheetName = null) => {
    skipDebouncedSearchRef.current = true;
    isManualTypingRef.current = false;
    // Show only the keyword part (strip conversational words) in search box
    if (userQuery && userQuery.trim() !== '') {
      const stopWords = new Set([
        'tampilkan', 'data', 'pada', 'yang', 'dan', 'di', 'untuk', 'ke', 'dari', 
        'adalah', 'itu', 'dengan', 'ini', 'oleh', 'seperti', 'adapun', 'atau', 
        'sebagai', 'tentang', 'yaitu', 'ia', 'kami', 'mereka', 'saya', 'anda', 
        'dia', 'kita', 'tahun', 'bulan', 'hari', 'file', 'berkas', 'dokumen', 
        'tabel', 'sheet', 'cari', 'temukan', 'info', 'informasi', 'lihat', 'berapa',
        'apakah', 'siapa', 'bagaimana', 'apa', 'tolong', 'jelaskan', 'rangkum', 'semua',
        'daftar', 'total', 'ada', 'berurutan', 'entry', 'entri'
      ]);
      const rawTerms = userQuery.trim().split(/[\s\.]+/).filter(Boolean);
      const filtered = rawTerms.filter(t => !stopWords.has(t.toLowerCase()) && t.length >= 2);
      const cleanQ = (filtered.length > 0 ? filtered : rawTerms).join(' ');
      setSearchQuery(cleanQ);
    } else {
      setSearchQuery('');
    }
    if (fileId) {
      fetchFilePreview(fileId, null, null, rowNumber, userQuery, sheetName);
    } else if (filename) {
      setSearchQuery(filename);
    }
    setActiveTab('search');
    setIsChatOpen(false);
  };

  // Toggle row bookmark state
  const handleToggleBookmark = async (rowId, currentStatus, targetUserId = null) => {
    try {
      if (currentStatus) {
        // Deleting bookmark
        let userIdParam = '';
        let activeUserId = targetUserId;
        if (!activeUserId && user?.role === 'admin') {
          if (selectedBookmarkUserIds.length === 1) {
            activeUserId = selectedBookmarkUserIds[0];
          }
        }
        if (user?.role === 'admin' && activeUserId) {
          userIdParam = `?userId=${activeUserId}`;
        }
        const url = `/api/bookmarks/${rowId}${userIdParam}`;
        const res = await apiFetch(url, {
          method: 'DELETE',
          headers: { 'Content-Type': 'application/json' }
        });
        if (res.ok) {
          setSearchResults(prev => prev.map(file => ({
            ...file,
            matches: file.matches.map(m => m.id === rowId ? { ...m, isBookmarked: false } : m)
          })));
          fetchBookmarks();
          showToast('⭐ Bookmark Dihapus', 'Baris dokumen berhasil dihapus dari bookmark.', 'success');
        }
        return;
      }

      // Adding bookmark - show beautiful React modal
      setBookmarkGroupModalRowId(rowId);
      setBookmarkGroupModalCurrentStatus(currentStatus);
      setBookmarkGroupModalTargetUserId(targetUserId);

      // Extract existing categories to show in dropdown
      const existingGroups = new Set();
      bookmarksList.forEach(b => {
        if (b.group_name) {
          existingGroups.add(b.group_name);
        }
      });
      const groupsList = Array.from(existingGroups).filter(g => g !== 'Semua');
      
      if (groupsList.length > 0) {
        setBookmarkGroupModalSelectedExisting(groupsList[0]);
        setBookmarkGroupModalSelectedOption('existing');
      } else {
        setBookmarkGroupModalSelectedExisting('Umum');
        setBookmarkGroupModalSelectedOption('new');
        setBookmarkGroupModalNewInput('Umum');
      }
      setBookmarkGroupModalNewInput('');
      setShowBookmarkGroupModal(true);
    } catch (err) {
      console.error('Gagal toggle bookmark:', err);
      showToast('⚠️ Gagal', 'Gagal memproses bookmark.', 'error');
    }
  };

  const handleConfirmBookmarkToggle = async () => {
    const rowId = bookmarkGroupModalRowId;
    const targetUserId = bookmarkGroupModalTargetUserId;

    let chosenGroup = 'Umum';
    if (bookmarkGroupModalSelectedOption === 'existing') {
      chosenGroup = bookmarkGroupModalSelectedExisting;
    } else {
      chosenGroup = bookmarkGroupModalNewInput.trim() !== '' ? bookmarkGroupModalNewInput.trim() : 'Umum';
    }

    try {
      let userIdParam = '';
      let activeUserId = targetUserId;
      if (!activeUserId && user?.role === 'admin') {
        if (selectedBookmarkUserIds.length === 1) {
          activeUserId = selectedBookmarkUserIds[0];
        }
      }
      if (user?.role === 'admin' && activeUserId) {
        userIdParam = `?userId=${activeUserId}`;
      }
      
      const url = `/api/bookmarks${userIdParam}`;
      const body = JSON.stringify({ 
        rowId, 
        userId: (user?.role === 'admin' && activeUserId) ? activeUserId : null,
        group_name: chosenGroup
      });

      const res = await apiFetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body
      });

      if (res.ok) {
        setSearchResults(prev => prev.map(file => ({
          ...file,
          matches: file.matches.map(m => m.id === rowId ? { ...m, isBookmarked: true } : m)
        })));
        fetchBookmarks();
        setShowBookmarkGroupModal(false);
        showToast('⭐ Tersimpan ke Bookmark', `Baris dokumen disimpan ke kategori "${chosenGroup}".`, 'success');
      } else {
        const errData = await res.json();
        throw new Error(errData.error || 'Gagal menyimpan bookmark.');
      }
    } catch (err) {
      showToast('⚠️ Gagal', err.message || 'Gagal memproses bookmark.', 'error');
    }
  };

  // Helper to start fake progress for deletion
  const startDeleteProgress = (message) => {
    setDeleting(true);
    setDeleteProgress(5);
    setDeleteStatusMessage(message);
    
    const progressMessages = [
      'Menghubungi server...',
      'Memverifikasi hak akses berkas...',
      'Membatalkan indeks pencarian terkait...',
      'Menghapus baris-baris dokumen dari database...',
      'Membersihkan metadata berkas...',
      'Menghapus entitas fisik...',
      'Menyinkronkan status database...',
      'Hampir selesai...'
    ];
    
    let currentPercent = 5;
    const interval = setInterval(() => {
      currentPercent += Math.floor(Math.random() * 8) + 5; // increment by 5-12%
      if (currentPercent > 92) {
        currentPercent = 92;
        clearInterval(interval);
      }
      setDeleteProgress(currentPercent);
      
      const msgIndex = Math.min(
        Math.floor((currentPercent / 100) * progressMessages.length),
        progressMessages.length - 1
      );
      setDeleteStatusMessage(progressMessages[msgIndex]);
    }, 250);
    
    return interval;
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

    const progressInterval = startDeleteProgress(`Memulai penghapusan massal ${selectedFileIds.length} berkas...`);

    try {
      const res = await apiFetch('/api/files/bulk-delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids: selectedFileIds })
      });

      clearInterval(progressInterval);
      const data = await res.json();
      
      if (res.ok) {
        setDeleteProgress(100);
        setDeleteStatusMessage('Penghapusan massal sukses!');
        setTimeout(() => {
          setDeleting(false);
          showToast('🗑️ Penghapusan Massal Sukses', `${selectedFileIds.length} file berhasil dihapus secara permanen.`, 'success');
          setSelectedFileIds([]);
          setSelectedFileId(null);
          fetchFiles();
        }, 500);
      } else {
        setDeleting(false);
        showToast('⚠️ Gagal Hapus Massal', data.error || 'Terjadi kesalahan.', 'error');
      }
    } catch (err) {
      clearInterval(progressInterval);
      setDeleting(false);
      console.error('Error bulk delete:', err);
      showToast('⚠️ Error Koneksi', 'Gagal menghubungi server.', 'error');
    }
  };

  // Delete file
  const handleDeleteFile = async (id, filename) => {
    if (!window.confirm(`Apakah Anda yakin ingin menghapus file "${filename}"? Semua data di dalamnya akan terhapus permanen dari aplikasi.`)) {
      return;
    }

    const progressInterval = startDeleteProgress(`Memulai penghapusan berkas "${filename}"...`);

    try {
      const res = await apiFetch(`/api/files/${id}`, { method: 'DELETE' });
      clearInterval(progressInterval);
      
      if (res.ok) {
        setDeleteProgress(100);
        setDeleteStatusMessage('Berkas berhasil dihapus!');
        setTimeout(() => {
          setDeleting(false);
          showToast('🗑️ File Dihapus', `Berkas "${filename}" berhasil dihapus.`, 'success');
          setSelectedFileId(null);
          fetchFiles();
        }, 500);
      } else {
        setDeleting(false);
        const data = await res.json();
        alert('Gagal menghapus file: ' + (data.error || 'Unknown error'));
      }
    } catch (err) {
      clearInterval(progressInterval);
      setDeleting(false);
      alert('Error saat menghapus file: ' + err.message);
    }
  };

  // Download file with active progress bar
  const handleDownloadFile = async (id, filename) => {
    if (downloading) return;

    downloadingRef.current = true;
    setDownloading(true);
    setDownloadingFileId(id);
    setDownloadProgress(5);
    setDownloadStatusMessage(`Menyiapkan data berkas "${filename}" dari database...`);

    // Simulated initial server processing progress (while server queries DB)
    const simInterval = setInterval(() => {
      setDownloadProgress(prev => {
        if (prev < 40) return prev + Math.floor(Math.random() * 6) + 3;
        return prev;
      });
    }, 180);

    try {
      const res = await apiFetch(`/api/files/${id}/download`);

      clearInterval(simInterval);

      if (!res.ok) {
        downloadingRef.current = false;
        setDownloading(false);
        setDownloadingFileId(null);
        let errMsg = 'Gagal mengunduh file';
        try {
          const errData = await res.json();
          errMsg = errData.error || errMsg;
        } catch (_) {}
        alert(errMsg);
        return;
      }

      setDownloadStatusMessage(`Mengekspor & mengunduh berkas "${filename}"...`);
      setDownloadProgress(45);

      const contentLength = res.headers.get('content-length');
      const totalBytes = contentLength ? parseInt(contentLength, 10) : 0;

      let blob;
      const reader = res.body ? res.body.getReader() : null;

      if (reader && totalBytes > 0) {
        let receivedBytes = 0;
        const chunks = [];
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          chunks.push(value);
          receivedBytes += value.length;
          const pct = Math.min(99, 45 + Math.round((receivedBytes / totalBytes) * 54));
          setDownloadProgress(pct);
        }
        blob = new Blob(chunks, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      } else {
        setDownloadProgress(85);
        blob = await res.blob();
      }

      setDownloadProgress(100);
      setDownloadStatusMessage(`Berkas "${filename}" siap diunduh!`);

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const safeFilename = filename.endsWith('.xlsx') || filename.endsWith('.xls') ? filename : `${filename}.xlsx`;
      a.download = safeFilename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      setTimeout(() => {
        downloadingRef.current = false;
        setDownloading(false);
        setDownloadingFileId(null);
        showToast('✅ Berhasil', `File "${safeFilename}" telah berhasil diunduh.`, 'success');
      }, 500);
    } catch (err) {
      clearInterval(simInterval);
      downloadingRef.current = false;
      setDownloading(false);
      setDownloadingFileId(null);
      alert('Error saat mengunduh file: ' + err.message);
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

  // Upload file logic with Chunked Upload (bypasses Cloudflare 100MB POST payload limit)
  const processFileUpload = async (file) => {
    const fileExt = file.name.split('.').pop().toLowerCase();
    if (fileExt !== 'xlsx' && fileExt !== 'xls') {
      setUploadError('Tipe file salah! Harap unggah file dengan format .xlsx atau .xls');
      setUploadSuccess(null);
      return;
    }

    setUploading(true);
    setUploadProgress(5);
    setUploadStatusMessage('Menyiapkan berkas untuk diunggah...');
    setUploadError(null);
    setUploadSuccess(null);

    const CHUNK_SIZE = 8 * 1024 * 1024; // 8MB per chunk (well below Cloudflare 100MB limit)
    const totalChunks = Math.ceil(file.size / CHUNK_SIZE);
    const uploadId = Date.now() + '_' + Math.random().toString(36).substring(2, 9);

    try {
      let jobId = null;

      if (totalChunks <= 1) {
        // Single direct upload for small files (<= 8MB)
        setUploadProgress(10);
        setUploadStatusMessage('Mengunggah berkas ke server...');

        const formData = new FormData();
        formData.append('file', file);

        const res = await apiFetch('/api/upload', {
          method: 'POST',
          body: formData,
        });

        let uploadData = {};
        try {
          uploadData = await res.json();
        } catch (_) {
          if (res.status === 413) {
            throw new Error('Ukuran file terlalu besar untuk diproses secara langsung.');
          } else {
            throw new Error(`Gagal mengunggah file (${res.status} ${res.statusText}).`);
          }
        }

        if (!res.ok) {
          throw new Error(uploadData.error || `Gagal mengunggah file (${res.status}).`);
        }

        jobId = uploadData.jobId;
      } else {
        // Chunked upload for large files (> 8MB, e.g., 155MB)
        for (let i = 0; i < totalChunks; i++) {
          const start = i * CHUNK_SIZE;
          const end = Math.min(file.size, start + CHUNK_SIZE);
          const chunk = file.slice(start, end);

          const formData = new FormData();
          formData.append('uploadId', uploadId);
          formData.append('chunkIndex', i);
          formData.append('totalChunks', totalChunks);
          formData.append('chunk', chunk, file.name);

          const pct = Math.floor(((i + 1) / totalChunks) * 20);
          setUploadProgress(pct);
          const mbUploaded = Math.round((end / (1024 * 1024)) * 10) / 10;
          const mbTotal = Math.round((file.size / (1024 * 1024)) * 10) / 10;
          setUploadStatusMessage(`Mengunggah bagian ${i + 1}/${totalChunks} (${mbUploaded} MB / ${mbTotal} MB)...`);

          const chunkRes = await apiFetch('/api/upload/chunk', {
            method: 'POST',
            body: formData,
          });

          if (!chunkRes.ok) {
            let chunkErr = 'Gagal mengunggah potongan file.';
            try {
              const errData = await chunkRes.json();
              chunkErr = errData.error || chunkErr;
            } catch (_) {}
            throw new Error(chunkErr);
          }
        }

        // Trigger merge chunks endpoint
        setUploadProgress(20);
        setUploadStatusMessage('Menggabungkan potongan berkas di server...');

        const mergeRes = await apiFetch('/api/upload/merge', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            uploadId,
            filename: file.name,
            totalChunks
          })
        });

        let mergeData = {};
        try {
          mergeData = await mergeRes.json();
        } catch (_) {
          throw new Error(`Gagal menggabungkan berkas di server (${mergeRes.status} ${mergeRes.statusText}).`);
        }

        if (!mergeRes.ok) {
          throw new Error(mergeData.error || `Gagal menggabungkan berkas di server (${mergeRes.status}).`);
        }

        jobId = mergeData.jobId;
      }

      setUploadProgress(20);
      setUploadStatusMessage('Berkas terunggah. Memulai impor data ke database...');

      // 2. Poll status endpoint until completed or failed
      const pollInterval = setInterval(async () => {
        try {
          const statusRes = await apiFetch(`/api/upload/status/${jobId}`);
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
  // Helper function to highlight query matches (Google-style highlighting)
  const highlightText = (text, highlight) => {
    if (text === null || text === undefined) return '';
    const textStr = text.toString();
    if (!highlight || highlight.trim() === '') return textStr;
    
    const baseTokens = highlight.trim().split(/[\s\.]+/).filter(Boolean);
    if (baseTokens.length === 0) return textStr;

    // Filter out common Indonesian stopwords unless query is only stop words
    const stopWords = new Set(['data', 'dan', 'di', 'ke', 'dari', 'yang', 'pada', 'untuk']);
    const nonStopTokens = baseTokens.filter(t => !stopWords.has(t.toLowerCase()));
    const targetTokens = nonStopTokens.length > 0 ? nonStopTokens : baseTokens;

    const tokens = [];
    targetTokens.forEach(t => {
      tokens.push(t);
      const match = t.match(/^([a-zA-Z]+)(0*[1-9]\d*)$/);
      if (match) {
        tokens.push(match[1]); // letters (e.g. A)
        tokens.push(match[2]); // digits (e.g. 001)
      }
    });

    const escapedTokens = Array.from(new Set(tokens))
      .map(t => t.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'))
      .filter(Boolean);

    if (escapedTokens.length === 0) return textStr;
    
    const splitRegex = new RegExp(`(${escapedTokens.join('|')})`, 'gi');
    const parts = textStr.split(splitRegex);
    const testRegex = new RegExp(`^(${escapedTokens.join('|')})$`, 'i');

    return (
      <>
        {parts.map((part, i) => 
          testRegex.test(part) 
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
    const activeFile = searchResults.find(f => String(f.fileId) === String(selectedFileId));
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
  const selectedFileData = selectedFileId 
    ? (searchResults.find(f => String(f.fileId) === String(selectedFileId)) || null) 
    : null;

  if (!token || !user) {
    return (
      <LoginScreen 
        setToken={setToken} 
        setUser={setUser} 
        isDarkMode={isDarkMode} 
        setIsDarkMode={setIsDarkMode} 
        showToast={showToast}
      />
    );
  }

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
          <div className="stats-badge" style={{ display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
            <span className="stats-item">📁 <strong>{stats.totalFiles}</strong> File</span>
            <span className="stats-divider">|</span>
            <span className="stats-item">⚡ <strong>{stats.totalRows.toLocaleString('id-ID')}</strong> Baris</span>
            <span className="stats-divider">|</span>
            <span className="stats-item" style={{ color: 'var(--text-secondary)' }}>
              📅 {currentTime.toLocaleDateString('id-ID', { weekday: 'short', day: 'numeric', month: 'short' })} • ⏰ {currentTime.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}
            </span>
          </div>
          
          {/* Theme Toggle */}
          <button
            className="theme-toggle-btn"
            onClick={() => setIsDarkMode(m => !m)}
            title={isDarkMode ? 'Beralih ke Mode Terang' : 'Beralih ke Mode Gelap'}
          >
            {isDarkMode ? '☀️ Terang' : '🌙 Gelap'}
          </button>

          {/* Notification Bell */}
          <div className="notification-container" style={{ position: 'relative' }}>
            <button 
              className="notification-bell-btn"
              onClick={() => setShowNotificationsDropdown(!showNotificationsDropdown)}
              title="Notifikasi"
            >
              🔔
              {notificationsList.filter(n => !n.is_read).length > 0 && (
                <span className="notification-badge">
                  {notificationsList.filter(n => !n.is_read).length}
                </span>
              )}
            </button>

            {showNotificationsDropdown && (
              <div className="notification-dropdown">
                <div className="notification-dropdown-header">
                  <h4>Notifikasi Global</h4>
                  {notificationsList.filter(n => !n.is_read).length > 0 && (
                    <button 
                      className="mark-all-read-btn"
                      onClick={() => handleMarkAsRead()}
                    >
                      Tandai semua dibaca
                    </button>
                  )}
                </div>
                <div className="notification-dropdown-list">
                  {notificationsList.length === 0 ? (
                    <div className="no-notifications">Tidak ada notifikasi baru</div>
                  ) : (
                    notificationsList.map(n => (
                      <div 
                        key={n.id} 
                        className={`notification-item ${!n.is_read ? 'unread' : ''}`}
                        onClick={() => handleMarkAsRead(n.id)}
                      >
                        <div className="notification-status-dot"></div>
                        <div className="notification-item-content">
                          <p className="notification-message">{n.message}</p>
                          <span className="notification-time">
                            {new Date(n.created_at).toLocaleString('id-ID')}
                          </span>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Logged in User info */}
          <div className="user-profile-badge">
            <span className="user-avatar" style={{ fontSize: '1.2rem' }}>👤</span>
            <div className="user-info-text">
              <span style={{ fontSize: '0.68rem', color: 'var(--text-secondary)', fontWeight: 500 }}>
                {getGreeting()}
              </span>
              <span className="username-label" style={{ fontWeight: '700' }} title={user.username}>
                {user.full_name || user.username}
              </span>
              <span className={`role-label role-${user.role}`}>{user.role.toUpperCase()}</span>
            </div>
            <button className="logout-btn" onClick={handleLogout} title="Keluar dari Aplikasi">
              🚪 Keluar
            </button>
          </div>

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
        {(user.role === 'admin' || user.role === 'operator') && (
          <button 
            className={`tab-btn ${activeTab === 'upload' ? 'active' : ''}`}
            onClick={() => setActiveTab('upload')}
          >
            📤 Unggah Excel
          </button>
        )}
        {(user.role === 'admin' || user.role === 'operator') && (
          <button 
            className={`tab-btn ${activeTab === 'files' ? 'active' : ''}`}
            onClick={() => setActiveTab('files')}
          >
            ⚙️ Kelola File ({filesList.length})
          </button>
        )}
        {user.role === 'admin' && (
          <button 
            className={`tab-btn ${activeTab === 'logs' ? 'active' : ''}`}
            onClick={() => setActiveTab('logs')}
          >
            📜 Log Aktivitas
          </button>
        )}
        {user.role === 'admin' && (
          <button 
            className={`tab-btn ${activeTab === 'dashboard' ? 'active' : ''}`}
            onClick={() => { setActiveTab('dashboard'); fetchDashboard(); }}
          >
            📊 Dashboard
          </button>
        )}
        {user.role === 'admin' && (
          <button 
            className={`tab-btn ${activeTab === 'users' ? 'active' : ''}`}
            onClick={() => setActiveTab('users')}
          >
            👥 Kelola Pengguna
          </button>
        )}
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
                  onChange={(e) => {
                    isManualTypingRef.current = true;
                    setSearchQuery(e.target.value);
                  }}
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
                {user?.role === 'admin' && (
                  <div className="filter-group">
                    <label>Saring Pengunggah:</label>
                    <select
                      className="filter-input-field"
                      style={{ background: 'var(--bg-tertiary)', border: '1px solid var(--glass-border)', color: '#fff', padding: '0.55rem', borderRadius: '6px', fontSize: '0.85rem', cursor: 'pointer' }}
                      value={searchUploaderFilter}
                      onChange={(e) => {
                        const val = e.target.value;
                        setSearchUploaderFilter(val);
                        if (searchQuery.trim() !== '') {
                          setTimeout(() => {
                            handleSearch();
                          }, 50);
                        } else {
                          const filtered = filesList.filter(f => val === 'all' || String(f.uploaded_by) === String(val));
                          if (filtered.length > 0) {
                            const isStillValid = filtered.some(f => f.id === selectedFileId);
                            if (!isStillValid) {
                              setSelectedFileId(filtered[0].id);
                              fetchFilePreview(filtered[0].id);
                            }
                          } else {
                            setSelectedFileId(null);
                          }
                        }
                      }}
                    >
                      <option value="all">📁 Semua Pengunggah</option>
                      {usersList.map(u => (
                        <option key={u.id} value={u.id}>
                          👤 {u.full_name ? `${u.full_name} (${u.username})` : u.username}
                        </option>
                      ))}
                    </select>
                  </div>
                )}
                <div className="filter-group ai-search-toggle-group">
                  <label className="ai-checkbox-label">
                    <input
                      type="checkbox"
                      checked={enableAISearch}
                      onChange={(e) => {
                        const val = e.target.checked;
                        setEnableAISearch(val);
                        if (searchQuery.trim() !== '') {
                          setTimeout(() => {
                            handleSearch();
                          }, 50);
                        }
                      }}
                    />
                    <span className="ai-checkbox-custom"></span>
                    <span>Aktifkan Pencarian Konteks AI</span>
                  </label>
                </div>
                {(filterSheet || filterUnit || (user?.role === 'admin' && searchUploaderFilter !== 'all')) && (
                  <button
                    className="clear-filters-btn"
                    onClick={() => { 
                      setFilterSheet(''); 
                      setFilterUnit(''); 
                      setSearchUploaderFilter('all');
                      if (searchQuery.trim() !== '') {
                        setTimeout(() => { handleSearch(); }, 50);
                      }
                    }}
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
                <div className="progress-bar-header">
                  <span className="loader-status-text">⚡ Mengambil data dari database...</span>
                  <span className="loader-percent-text">{Math.min(100, Math.max(1, Math.round(searchProgress)))}%</span>
                </div>
                <div className="progress-bar-container">
                  <div 
                    className="progress-bar-fill" 
                    style={{ width: `${Math.min(100, Math.max(1, Math.round(searchProgress)))}%` }}
                  ></div>
                </div>
              </div>
            )}

            {/* Google-Style Search Statistics Banner */}
            {searchUIState === 'results' && searchQuery.trim() !== '' && (
              <div className="google-search-stats-banner">
                <div className="google-stats-info">
                  🔍 Menampilkan hasil pencarian untuk <strong>"{searchQuery.trim()}"</strong> — 
                  Sekitar <strong>{searchStats.totalMatches || searchResults.reduce((acc, f) => acc + (f.matches ? f.matches.length : 0), 0)}</strong> baris cocok di <strong>{searchResults.length}</strong> dokumen
                  {searchStats.timeMs > 0 && <span className="google-time-text"> ({(searchStats.timeMs / 1000).toFixed(2)} detik)</span>}
                </div>
                {aiSearchInfo.used && (
                  <div className="google-ai-tag">
                    ⚡ AI Smart Search ({aiSearchInfo.method === 'gemini' ? 'Gemini 2.0' : 'Vektor Lokal'})
                  </div>
                )}
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
                      (() => {
                        const displayedFiles = filesList.filter(f => {
                          if (searchUploaderFilter === 'all') return true;
                          return String(f.uploaded_by) === String(searchUploaderFilter);
                        });
                        if (displayedFiles.length === 0) {
                          return <div style={{ fontSize: '0.82rem', color: 'var(--text-muted)', fontStyle: 'italic', padding: '1rem', textAlign: 'center' }}>Tidak ada dokumen dari pengunggah ini.</div>;
                        }
                        return displayedFiles.map(file => (
                          <button
                            key={file.id}
                            className={`file-tab-btn ${String(selectedFileId) === String(file.id) ? 'active' : ''}`}
                            onClick={() => handleFileTabClick(file.id)}
                            style={{ display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: '4px' }}
                          >
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%', gap: '8px' }}>
                              <div className="file-tab-name">📄 {file.filename}</div>
                              <div className="file-tab-badge">⚡ {file.row_count} baris</div>
                            </div>
                            {user?.role === 'admin' && (
                              <div style={{ fontSize: '0.68rem', opacity: 0.75, textAlign: 'left', borderTop: '1px solid rgba(255,255,255,0.05)', paddingTop: '3px', color: 'var(--accent-secondary)' }}>
                                👤 Pengunggah: {file.uploader_fullname || file.uploader_username || 'Sistem'}
                              </div>
                            )}
                          </button>
                        ));
                      })()
                    ) : (
                      searchResults.map(file => {
                        // Find uploader info from filesList
                        const fileMeta = filesList.find(f => f.id === file.fileId);
                        const uploaderName = fileMeta ? (fileMeta.uploader_fullname || fileMeta.uploader_username || 'Sistem') : 'Sistem';
                        
                        return (
                          <button
                            key={file.fileId}
                            className={`file-tab-btn ${String(selectedFileId) === String(file.fileId) ? 'active' : ''}`}
                            onClick={() => handleFileTabClick(file.fileId)}
                            style={{ display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: '4px' }}
                          >
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%', gap: '8px' }}>
                              <div className="file-tab-name">📄 {file.filename}</div>
                              <div className="file-tab-badge">{file.matches.length} baris cocok</div>
                            </div>
                            {user?.role === 'admin' && (
                              <div style={{ fontSize: '0.68rem', opacity: 0.75, textAlign: 'left', borderTop: '1px solid rgba(255,255,255,0.05)', paddingTop: '3px', color: 'var(--accent-secondary)' }}>
                                👤 Pengunggah: {uploaderName}
                              </div>
                            )}
                          </button>
                        );
                      })
                    )}
                  </div>
                </div>

                {/* Right Side: Data table of the selected file */}
                <div className="results-view">
                  {selectedFileData ? (
                    <div className="results-card">
                      <div className="results-card-header">
                        <div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', marginBottom: '4px' }}>
                            <h2 style={{ margin: 0 }}>📁 {selectedFileData.filename}</h2>
                            {aiSearchInfo.used && (
                              <span className={`ai-badge ai-badge-${aiSearchInfo.method}`} title={`Metode Pencarian: ${aiSearchInfo.method === 'gemini' ? 'Google Gemini API (Free Tier)' : 'Model Vektor Lokal (TF-IDF)'}`}>
                                ✨ AI Search ({aiSearchInfo.method === 'gemini' ? 'Gemini API' : 'Model Lokal'})
                              </span>
                            )}
                          </div>
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
                              <div className="active-filter-bar" style={{ display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: '0.4rem', padding: '0.9rem 1.25rem' }}>
                                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%', flexWrap: 'wrap', gap: '0.5rem' }}>
                                  <span className="active-filter-text">
                                    ℹ️ <strong>Filter Kolom Aktif:</strong> Menampilkan { (showSheet ? 1 : 0) + (showBaris ? 1 : 0) + orderedDataHeaders.length } dari { activeHeaders.length + 2 } kolom.
                                  </span>
                                  <button className="reset-filter-link-btn" onClick={() => setVisibleColumnsOrder([])}>
                                    ✕ Reset Filter Kolom
                                  </button>
                                </div>
                                <div className="selected-columns-list-info" style={{ fontSize: '0.78rem', color: 'var(--text-secondary)', lineHeight: '1.4', borderTop: '1px solid var(--glass-border)', paddingTop: '0.4rem' }}>
                                  <strong>Kolom Terpilih:</strong> {
                                    [
                                      ...(showSheet ? ['Sheet'] : []),
                                      ...(showBaris ? ['Baris'] : []),
                                      ...orderedDataHeaders.map(h => getColumnLabel(h))
                                    ].join(', ')
                                  }
                                </div>
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
                            <DragScrollContainer className="table-responsive">
                              <table className="excel-table">
                                <thead>
                                  <tr>
                                    <th style={{ width: '40px', textAlign: 'center' }}>⭐</th>
                                    {aiSearchInfo.used && <th style={{ width: '80px', textAlign: 'center' }}>Relevansi</th>}
                                    {showSheet && <th>Sheet</th>}
                                    {showBaris && <th>Baris</th>}
                                    {orderedDataHeaders.map(h => (
                                      <th key={h}>{getColumnLabel(h)}</th>
                                    ))}
                                  </tr>
                                </thead>
                                <tbody>
                                  {selectedFileData.matches.map(match => (
                                    <tr
                                      key={match.id}
                                      id={`row-cited-${match.row_number}`}
                                      className={highlightedRowNumber && match.row_number === parseInt(highlightedRowNumber, 10) ? 'ai-cited-row-highlight' : ''}
                                    >
                                      <td className="meta-cell text-center" style={{ width: '40px' }}>
                                        <button
                                          className={`bookmark-star-btn ${match.isBookmarked ? 'active' : ''}`}
                                          onClick={() => handleToggleBookmark(match.id, match.isBookmarked)}
                                          title={match.isBookmarked ? 'Hapus dari Bookmark' : 'Simpan ke Bookmark'}
                                        >
                                          {match.isBookmarked ? '★' : '☆'}
                                        </button>
                                      </td>
                                      {aiSearchInfo.used && (
                                        <td className="meta-cell text-center font-accent" style={{ fontWeight: 'bold' }}>
                                          <div style={{
                                            display: 'inline-block',
                                            padding: '2px 8px',
                                            borderRadius: '12px',
                                            background: match.relevanceScore >= 80 ? 'rgba(16, 185, 129, 0.15)' : match.relevanceScore >= 50 ? 'rgba(245, 158, 11, 0.15)' : 'rgba(107, 114, 128, 0.15)',
                                            color: match.relevanceScore >= 80 ? '#10b981' : match.relevanceScore >= 50 ? '#f59e0b' : '#9ca3af',
                                            border: match.relevanceScore >= 80 ? '1px solid rgba(16, 185, 129, 0.25)' : match.relevanceScore >= 50 ? '1px solid rgba(245, 158, 11, 0.25)' : '1px solid rgba(107, 114, 128, 0.25)',
                                            fontSize: '0.78rem'
                                          }}>
                                            {match.relevanceScore}%
                                          </div>
                                        </td>
                                      )}
                                      {showSheet && (
                                        <td className="meta-cell font-accent">{match.sheetName}</td>
                                      )}
                                      {showBaris && (
                                        <td className="meta-cell text-center">{match.rowNumber}</td>
                                      )}
                                      {orderedDataHeaders.map(h => {
                                        const isDesc = h === 'Kolom_17' || h === 'Kolom_18' || h.toLowerCase().includes('perihal') || h.toLowerCase().includes('uraian');
                                        const isEditing = editingCell && editingCell.rowId === match.id && editingCell.headerKey === h;
                                        const canEdit = user?.role === 'admin' || user?.role === 'operator';

                                        return (
                                          <td
                                            key={h}
                                            className={`${isDesc ? 'description-cell' : ''} ${canEdit ? 'editable-cell' : ''} ${isEditing ? 'editing-active-cell' : ''}`}
                                            onDoubleClick={() => {
                                              if (canEdit && !isSavingCell) {
                                                setEditingCell({ rowId: match.id, headerKey: h });
                                                setEditCellValue(match.rowData[h] || '');
                                              }
                                            }}
                                            title={canEdit ? 'Klik 2x untuk mengedit' : ''}
                                          >
                                            {isEditing ? (
                                              <div className="inline-cell-editor-container">
                                                <input
                                                  type="text"
                                                  className="inline-cell-input"
                                                  value={editCellValue}
                                                  onChange={(e) => setEditCellValue(e.target.value)}
                                                  onBlur={() => handleSaveCell(match.id, h, match.rowData[h], selectedFileData.matches)}
                                                  onKeyDown={(e) => {
                                                    if (e.key === 'Enter') {
                                                      handleSaveCell(match.id, h, match.rowData[h], selectedFileData.matches);
                                                    } else if (e.key === 'Escape') {
                                                      handleCancelEdit();
                                                    }
                                                  }}
                                                  autoFocus
                                                  disabled={isSavingCell}
                                                />
                                                {isSavingCell && <span className="inline-cell-saving-spinner">⌛</span>}
                                              </div>
                                            ) : (
                                              highlightText(match.rowData[h], searchQuery)
                                            )}
                                          </td>
                                        );
                                      })}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </DragScrollContainer>
                          </>
                        );
                      })()}
                    </div>
                  ) : (
                    <div className="empty-state select-document-prompt" style={{ margin: '1.5rem auto', padding: '3.5rem 2rem', background: 'var(--bg-secondary)', border: '1px solid var(--glass-border)', borderRadius: '16px', textAlign: 'center' }}>
                      <div className="empty-icon" style={{ fontSize: '3rem', marginBottom: '1rem' }}>👈</div>
                      <h3 style={{ fontSize: '1.25rem', marginBottom: '0.5rem', color: 'var(--text-primary)' }}>Silakan Pilih Salah Satu Dokumen</h3>
                      <p style={{ color: 'var(--text-secondary)', fontSize: '0.9rem', maxWidth: '460px', margin: '0 auto', lineHeight: '1.5' }}>
                        Klik salah satu dokumen dari daftar di sebelah kiri untuk melihat isi data baris spreadsheet.
                      </p>
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
            <div className="dashboard-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px', borderBottom: '1px solid var(--glass-border)', paddingBottom: '1.25rem', marginBottom: '1.75rem' }}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                <h2 style={{ margin: 0, fontSize: '1.6rem', fontWeight: '700', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span>⭐</span>
                  <span style={{ background: 'linear-gradient(135deg, #fff 0%, #a78bfa 100%)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
                    Data Bookmark
                  </span>
                </h2>
                <p style={{ margin: 0, fontSize: '0.88rem', color: 'var(--text-secondary)', lineHeight: '1.4', maxWidth: '550px' }}>
                  Daftar baris data spreadsheet penting yang disimpan untuk referensi cepat.
                </p>
              </div>
              
              {user?.role === 'admin' && (
                <div className="bookmark-user-selector-container" style={{ position: 'relative' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <label style={{ fontSize: '0.85rem', fontWeight: '600', color: 'var(--text-secondary)' }}>Rekap Bookmark:</label>
                    <button
                      type="button"
                      className="column-filter-trigger-btn"
                      onClick={() => setShowUserFilterDropdown(!showUserFilterDropdown)}
                      style={{ padding: '0.5rem 1rem', background: 'var(--bg-secondary)', border: '1px solid var(--glass-border)', borderRadius: '8px', color: '#fff', fontSize: '0.88rem', cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                    >
                      👥 Pilih Akun ({selectedBookmarkUserIds.length === 0 ? 'Hanya Admin Anda' : `${selectedBookmarkUserIds.length} Akun Terpilih`}) ▾
                    </button>
                  </div>
                  
                  {showUserFilterDropdown && (
                    <div className="column-filter-dropdown" style={{ right: 0, width: '280px', top: '42px', zIndex: 110 }}>
                      <div className="dropdown-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 12px', borderBottom: '1px solid var(--glass-border)' }}>
                        <strong style={{ fontSize: '0.8rem' }}>Pilih Akun Pengguna</strong>
                        <div style={{ display: 'flex', gap: '6px' }}>
                          <button
                            type="button"
                            className="mark-all-read-btn"
                            onClick={() => {
                              setSelectedBookmarkUserIds(usersList.map(u => u.id));
                            }}
                            style={{ fontSize: '0.72rem' }}
                          >
                            Pilih Semua
                          </button>
                          <span style={{ color: 'var(--text-muted)', fontSize: '0.72rem' }}>|</span>
                          <button
                            type="button"
                            className="mark-all-read-btn"
                            onClick={() => {
                              setSelectedBookmarkUserIds([]);
                            }}
                            style={{ fontSize: '0.72rem', color: 'var(--text-muted)' }}
                          >
                            Bersihkan
                          </button>
                        </div>
                      </div>
                      
                      <div className="dropdown-items" style={{ maxHeight: '200px', overflowY: 'auto', padding: '6px' }}>
                        {usersList.map(u => {
                          const isChecked = selectedBookmarkUserIds.includes(u.id);
                          return (
                            <label key={u.id} className="column-checkbox-label" style={{ display: 'flex', alignItems: 'center', padding: '6px 8px', borderRadius: '4px', cursor: 'pointer', transition: 'background 0.2s', gap: '8px', fontSize: '0.85rem' }}>
                              <input
                                type="checkbox"
                                checked={isChecked}
                                onChange={() => {
                                  if (isChecked) {
                                    setSelectedBookmarkUserIds(prev => prev.filter(id => id !== u.id));
                                  } else {
                                    setSelectedBookmarkUserIds(prev => [...prev, u.id]);
                                  }
                                }}
                              />
                              <span style={{ color: u.username === user.username ? 'var(--accent-secondary)' : '#fff' }}>
                                {u.username === user.username ? `⭐ ${u.username} (Anda)` : `👤 ${u.username}`}
                              </span>
                              <span className={`role-badge role-${u.role}`} style={{ fontSize: '0.65rem', padding: '1px 4px', borderRadius: '3px', marginLeft: 'auto' }}>
                                {u.role.toUpperCase()}
                              </span>
                            </label>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {editingBookmark && !selectedUserForBookmarks && (
              <div className="bookmark-edit-form-wrapper" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--glass-border)', borderRadius: '12px', padding: '1.5rem', marginBottom: '2rem' }}>
                <h3 style={{ marginBottom: '1rem', fontSize: '1.1rem' }}>✏️ Edit Bookmark Baris #{editingBookmark.row_number} - {editingBookmark.filename}</h3>
                <form onSubmit={handleUpdateBookmark}>
                  <div className="form-group" style={{ marginBottom: '1rem' }}>
                    <label style={{ display: 'block', fontWeight: 'bold', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Nama Kategori / Grup Bookmark</label>
                    <input
                      type="text"
                      value={editingBookmarkGroup}
                      onChange={(e) => setEditingBookmarkGroup(e.target.value)}
                      placeholder="Contoh: Keuangan, Penting, Umum, dll."
                      style={{ width: '100%', padding: '0.65rem 0.85rem', background: 'var(--bg-tertiary)', border: '1px solid var(--glass-border)', borderRadius: '8px', color: '#fff', fontSize: '0.88rem', marginBottom: '1rem' }}
                    />
                  </div>
                  <div className="form-group" style={{ marginBottom: '1rem' }}>
                    <label style={{ display: 'block', fontWeight: 'bold', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Catatan / Keterangan Bookmark</label>
                    <textarea
                      value={editingBookmarkNotes}
                      onChange={(e) => setEditingBookmarkNotes(e.target.value)}
                      placeholder="Tambahkan catatan khusus untuk bookmark ini..."
                      style={{ width: '100%', padding: '0.65rem 0.85rem', background: 'var(--bg-tertiary)', border: '1px solid var(--glass-border)', borderRadius: '8px', color: '#fff', fontSize: '0.88rem', minHeight: '80px', fontFamily: 'inherit', resize: 'vertical' }}
                    />
                  </div>
                  
                  <div className="form-group" style={{ marginBottom: '1.5rem' }}>
                    <label style={{ display: 'block', fontWeight: 'bold', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Konten Baris Dokumen (Key-Values)</label>
                    <div style={{ maxHeight: '200px', overflowY: 'auto', border: '1px solid var(--glass-border)', borderRadius: '8px', padding: '0.85rem', background: 'var(--bg-tertiary)' }}>
                      {Object.entries(editingBookmarkRowData).map(([key, val]) => (
                        <div key={key} style={{ display: 'flex', alignItems: 'center', marginBottom: '0.6rem', gap: '10px' }}>
                          <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', width: '160px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }} title={key}>{key}:</span>
                          <input
                            type="text"
                            value={String(val)}
                            onChange={(e) => {
                              const updatedVal = e.target.value;
                              setEditingBookmarkRowData(prev => ({ ...prev, [key]: updatedVal }));
                            }}
                            style={{ flex: 1, padding: '0.4rem 0.75rem', background: 'var(--bg-secondary)', border: '1px solid var(--glass-border)', borderRadius: '6px', color: '#fff', fontSize: '0.85rem' }}
                          />
                        </div>
                      ))}
                    </div>
                  </div>

                  <div style={{ display: 'flex', gap: '12px' }}>
                    <button type="submit" className="btn-submit-user" style={{ padding: '0.55rem 1.25rem' }}>
                      Simpan Perubahan
                    </button>
                    <button 
                      type="button" 
                      className="btn-delete-user" 
                      onClick={() => setEditingBookmark(null)}
                      style={{ padding: '0.55rem 1.25rem', background: 'rgba(255, 255, 255, 0.05)', borderColor: 'var(--glass-border)', color: 'var(--text-secondary)' }}
                    >
                      Batal
                    </button>
                  </div>
                </form>
              </div>
            )}

            {loadingBookmarks ? (
              <div className="search-loader">
                <div className="spinner"></div>
                <p>Mengambil data bookmark...</p>
              </div>
            ) : bookmarksList.length === 0 ? (
              <div className="empty-state">
                <div className="empty-icon">⭐</div>
                <h3>Belum ada bookmark</h3>
                <p>
                  {selectedBookmarkUserIds.length === 1 && selectedBookmarkUserIds[0] === user.id
                    ? 'Klik tombol bintang (☆) pada tabel hasil pencarian untuk menyimpan data penting di sini.'
                    : 'Pengguna terpilih belum menyimpan berkas bookmark apa pun.'}
                </p>
              </div>
            ) : (
              <div className="results-card">
                {(() => {
                  // Filter bookmarks by active group
                  const filteredBookmarks = bookmarksList.filter(b => {
                    if (activeBookmarkGroupFilter === 'Semua') return true;
                    if (activeBookmarkGroupFilter === 'Umum') return !b.group_name || b.group_name === 'Umum';
                    return b.group_name === activeBookmarkGroupFilter;
                  });

                  // Build unique list of groups/categories
                  const uniqueGroups = new Set();
                  uniqueGroups.add('Semua');
                  bookmarksList.forEach(b => {
                    uniqueGroups.add(b.group_name || 'Umum');
                  });
                  const categoriesList = Array.from(uniqueGroups);

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
                      {/* Bookmark Category Filter Pills */}
                      <div className="bookmark-category-tabs" style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', padding: '1rem 1.25rem', marginBottom: '1rem', borderBottom: '1px solid var(--glass-border)', background: 'rgba(255, 255, 255, 0.02)', borderRadius: '8px' }}>
                        <span style={{ alignSelf: 'center', fontSize: '0.85rem', fontWeight: 600, color: 'var(--text-secondary)', marginRight: '8px' }}>📂 Kategori Bookmark:</span>
                        {categoriesList.map(cat => {
                          const count = bookmarksList.filter(b => {
                            if (cat === 'Semua') return true;
                            if (cat === 'Umum') return !b.group_name || b.group_name === 'Umum';
                            return b.group_name === cat;
                          }).length;
                          
                          return (
                            <button
                              key={cat}
                              type="button"
                              onClick={() => setActiveBookmarkGroupFilter(cat)}
                              className={`tab-btn ${activeBookmarkGroupFilter === cat ? 'active' : ''}`}
                              style={{ 
                                padding: '0.35rem 0.8rem', 
                                fontSize: '0.8rem', 
                                borderRadius: '20px', 
                                background: activeBookmarkGroupFilter === cat ? 'var(--accent-secondary)' : 'rgba(255,255,255,0.05)',
                                color: activeBookmarkGroupFilter === cat ? '#fff' : 'var(--text-secondary)',
                                border: '1px solid var(--glass-border)',
                                cursor: 'pointer',
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '6px',
                                transition: 'all 0.2s ease'
                              }}
                            >
                              <span>{cat}</span>
                              <span style={{ fontSize: '0.72rem', opacity: 0.8, background: 'rgba(0,0,0,0.2)', padding: '1px 6px', borderRadius: '10px' }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>

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
                        <div className="active-filter-bar" style={{ display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: '0.4rem', padding: '0.9rem 1.25rem', marginBottom: '1rem' }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%', flexWrap: 'wrap', gap: '0.5rem' }}>
                            <span className="active-filter-text">
                              ℹ️ <strong>Filter Kolom Aktif:</strong> Menampilkan { (showFile ? 1 : 0) + (showSheet ? 1 : 0) + (showBaris ? 1 : 0) + orderedBookmarkHeaders.length } dari { bookmarkedHeaders.length + 3 } kolom.
                            </span>
                            <button className="reset-filter-link-btn" onClick={() => setBookmarkVisibleColumns([])}>
                              ✕ Reset Filter Kolom
                            </button>
                          </div>
                          <div className="selected-columns-list-info" style={{ fontSize: '0.78rem', color: 'var(--text-secondary)', lineHeight: '1.4', borderTop: '1px solid var(--glass-border)', paddingTop: '0.4rem' }}>
                            <strong>Kolom Terpilih:</strong> {
                              [
                                ...(showFile ? ['File'] : []),
                                ...(showSheet ? ['Sheet'] : []),
                                ...(showBaris ? ['Baris'] : []),
                                ...orderedBookmarkHeaders.map(h => getColumnLabel(h))
                              ].join(', ')
                            }
                          </div>
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
                                  ...(selectedBookmarkUserIds.length > 1 ? ['Pemilik'] : []),
                                  'Kategori',
                                  ...(showFile ? ['File'] : []),
                                  ...(showSheet ? ['Sheet'] : []),
                                  ...(showBaris ? ['Baris'] : []),
                                  ...orderedBookmarkHeaders.map(h => getColumnLabel(h)),
                                  'Catatan'
                                ];
                                const rows = filteredBookmarks.map(m => [
                                  ...(selectedBookmarkUserIds.length > 1 ? [m.owner_username || 'Sistem'] : []),
                                  m.group_name || 'Umum',
                                  ...(showFile ? [m.filename] : []),
                                  ...(showSheet ? [m.sheet_name] : []),
                                  ...(showBaris ? [m.row_number] : []),
                                  ...orderedBookmarkHeaders.map(h => m.row_data[h] ?? ''),
                                  m.notes ?? ''
                                ]);
                                const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                                const wb = XLSX.utils.book_new();
                                XLSX.utils.book_append_sheet(wb, ws, 'Bookmarks');
                                XLSX.writeFile(wb, `Bookmarks-${new Date().toLocaleDateString('id-ID').replace(/\//g, '-')}.xlsx`);
                                showToast('📥 Berhasil Diekspor!', 'Bookmarks disimpan ke Excel.', 'success');
                              }}
                              title="Unduh seluruh bookmark ke Excel"
                            >
                              📥 Ekspor Excel ({filteredBookmarks.length})
                            </button>
                            <button
                              className="table-action-btn"
                              onClick={() => {
                                const headers = [
                                  ...(selectedBookmarkUserIds.length > 1 ? ['Pemilik'] : []),
                                  'Kategori',
                                  ...(showFile ? ['File'] : []),
                                  ...(showSheet ? ['Sheet'] : []),
                                  ...(showBaris ? ['Baris'] : []),
                                  ...orderedBookmarkHeaders.map(h => getColumnLabel(h)),
                                  'Catatan'
                                ];
                                const rows = filteredBookmarks.map(m => [
                                  ...(selectedBookmarkUserIds.length > 1 ? [m.owner_username || 'Sistem'] : []),
                                  m.group_name || 'Umum',
                                  ...(showFile ? [m.filename] : []),
                                  ...(showSheet ? [m.sheet_name] : []),
                                  ...(showBaris ? [m.row_number] : []),
                                  ...orderedBookmarkHeaders.map(h => m.row_data[h] ?? ''),
                                  m.notes ?? ''
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
                            <span className="table-row-count">{filteredBookmarks.length} baris tersimpan</span>
                          </div>

                          {filteredBookmarks.length === 0 ? (
                            <div className="empty-state" style={{ padding: '3rem 1.5rem', background: 'rgba(255, 255, 255, 0.01)', border: '1px solid var(--glass-border)', borderRadius: '10px' }}>
                              <p style={{ color: 'var(--text-muted)', fontStyle: 'italic', fontSize: '0.9rem' }}>Tidak ada data bookmark dalam kategori "{activeBookmarkGroupFilter}".</p>
                            </div>
                          ) : (
                            <DragScrollContainer className="table-responsive">
                              <table className="excel-table">
                                <thead>
                                  <tr>
                                    <th style={{ width: '40px', textAlign: 'center' }}>⭐</th>
                                    {selectedBookmarkUserIds.length > 1 && <th>Pemilik</th>}
                                    <th>Kategori</th>
                                    {showFile && <th>Nama Berkas</th>}
                                    {showSheet && <th>Sheet</th>}
                                    {showBaris && <th>Baris</th>}
                                    {orderedBookmarkHeaders.map(h => (
                                      <th key={h}>{getColumnLabel(h)}</th>
                                    ))}
                                    <th>Catatan</th>
                                    <th>Aksi</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {filteredBookmarks.map(match => (
                                    <tr key={match.id}>
                                      <td className="meta-cell text-center" style={{ width: '40px' }}>
                                        <button
                                          className="bookmark-star-btn active"
                                          onClick={() => handleToggleBookmark(match.id, true, match.owner_id)}
                                          title="Hapus dari Bookmark"
                                        >
                                          ★
                                        </button>
                                      </td>
                                      {selectedBookmarkUserIds.length > 1 && (
                                        <td className="meta-cell">
                                          <span className="role-badge role-admin" style={{ padding: '0.2rem 0.5rem', fontSize: '0.72rem', borderRadius: '4px', display: 'inline-flex', alignItems: 'center', gap: '4px', background: 'rgba(139, 92, 246, 0.12)', border: '1px solid rgba(139, 92, 246, 0.25)', color: '#c084fc' }}>
                                            👤 {match.owner_username || 'Sistem'}
                                          </span>
                                        </td>
                                      )}
                                      <td className="meta-cell">
                                        <span className="role-badge role-viewer" style={{ padding: '0.2rem 0.5rem', fontSize: '0.72rem', borderRadius: '4px', display: 'inline-flex', alignItems: 'center', gap: '4px', background: 'rgba(16, 185, 129, 0.12)', border: '1px solid rgba(16, 185, 129, 0.25)', color: '#34d399', fontWeight: 600 }}>
                                          📂 {match.group_name || 'Umum'}
                                        </span>
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
                                      <td>
                                        {match.notes ? (
                                          <span style={{ fontSize: '0.88rem' }}>{match.notes}</span>
                                        ) : (
                                          <span style={{ color: 'var(--text-muted)', fontStyle: 'italic', fontSize: '0.8rem' }}>-</span>
                                        )}
                                      </td>
                                      <td>
                                        <button
                                          type="button"
                                          className="btn-edit-bookmark"
                                          onClick={() => {
                                            setEditingBookmark(match);
                                            setEditingBookmarkNotes(match.notes || '');
                                            setEditingBookmarkGroup(match.group_name || 'Umum');
                                            setEditingBookmarkRowData({ ...match.row_data });
                                          }}
                                          title="Edit catatan, kategori & data baris bookmark"
                                          style={{ padding: '0.25rem 0.5rem', fontSize: '0.75rem' }}
                                        >
                                          ✏️ Edit
                                        </button>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </DragScrollContainer>
                          )}
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
                  
                  {user.role === 'admin' && selectedFileIds.length > 0 && (
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
                    <div key={file.id} className={`file-card ${user.role === 'admin' && selectedFileIds.includes(file.id) ? 'selected' : ''}`}>
                      {user.role === 'admin' && (
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
                      )}
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
                      <div className="file-card-actions" style={{ display: 'flex', gap: '8px', marginTop: '0.75rem', flexWrap: 'wrap', width: '100%' }}>
                        {downloadingFileId === file.id ? (
                          <div className="card-download-progress">
                            <div className="card-progress-bar-bg">
                              <div className="card-progress-bar-fill" style={{ width: `${downloadProgress}%` }}></div>
                            </div>
                            <span className="card-progress-text">📥 Mengunduh Excel... {downloadProgress}%</span>
                          </div>
                        ) : (
                          <>
                            <button 
                              className="download-file-btn" 
                              onClick={() => handleDownloadFile(file.id, file.filename)}
                              disabled={downloading}
                              title="Unduh berkas Excel ini"
                              style={{
                                padding: '0.45rem 0.85rem',
                                borderRadius: '8px',
                                border: '1px solid rgba(59, 130, 246, 0.4)',
                                background: 'rgba(59, 130, 246, 0.15)',
                                color: '#60a5fa',
                                cursor: downloading ? 'not-allowed' : 'pointer',
                                opacity: downloading ? 0.6 : 1,
                                fontSize: '0.82rem',
                                fontWeight: '600',
                                transition: 'all 0.2s ease',
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '4px'
                              }}
                            >
                              📥 Unduh Excel
                            </button>
                            {user.role === 'admin' && (
                              <button 
                                className="delete-file-btn" 
                                onClick={() => handleDeleteFile(file.id, file.filename)}
                                disabled={downloading}
                                title="Hapus file dan seluruh datanya"
                              >
                                🗑️ Hapus
                              </button>
                            )}
                          </>
                        )}
                      </div>
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
            <div className="section-header-row" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px', borderBottom: '1px solid var(--glass-border)', paddingBottom: '1.25rem', marginBottom: '1.75rem' }}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                <h2 style={{ margin: 0, fontSize: '1.6rem', fontWeight: '700', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span>📜</span>
                  <span style={{ background: 'linear-gradient(135deg, #fff 0%, #a78bfa 100%)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
                    Log Aktivitas Sistem
                  </span>
                </h2>
                <p style={{ margin: 0, fontSize: '0.88rem', color: 'var(--text-secondary)', lineHeight: '1.4', maxWidth: '550px' }}>
                  Riwayat audit pengunggahan, pencarian, dan penghapusan berkas.
                </p>
              </div>
              <div style={{ display: 'flex', gap: '10px' }}>
                <button className="btn-refresh" onClick={() => fetchLogs(logSearch, logPage, logLimit)} disabled={loadingLogs} style={{ height: '40px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                  {loadingLogs ? '⏳ Memuat...' : '🔄 Refresh'}
                </button>
                {user?.role === 'admin' && (
                  <button 
                    className="btn-delete-user" 
                    onClick={handleClearLogs}
                    style={{ height: '40px', display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '0.6rem 1.25rem', fontSize: '0.88rem' }}
                  >
                    🗑️ Reset Log
                  </button>
                )}
              </div>
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
                <div className="table-responsive logs-table-container">
                  <table className="excel-table logs-table">
                    <thead>
                      <tr>
                        <th>Waktu Kejadian</th>
                        <th>Akun</th>
                        <th>Tipe</th>
                        <th>Detail Aktivitas</th>
                        <th>IP Address</th>
                        <th>Perangkat & OS</th>
                        <th>Browser</th>
                        <th style={{ width: '80px', textAlign: 'center' }}>Aksi</th>
                      </tr>
                    </thead>
                    <tbody>
                      {logs.map(log => {
                        let tagClass = 'tag-secondary';
                        let typeLabel = log.activity_type;
                        if (log.activity_type === 'upload')            { tagClass = 'tag-success';   typeLabel = '📤 Unggah'; }
                        if (log.activity_type === 'delete')            { tagClass = 'tag-danger';    typeLabel = '🗑️ Hapus'; }
                        if (log.activity_type === 'search')            { tagClass = 'tag-info';      typeLabel = '🔍 Cari'; }
                        if (log.activity_type === 'access')            { tagClass = 'tag-access';    typeLabel = '🔑 Akses'; }
                        if (log.activity_type === 'reset_stats')       { tagClass = 'tag-danger';    typeLabel = '♻️ Reset'; }
                        if (log.activity_type === 'clear_logs')        { tagClass = 'tag-danger';    typeLabel = '🗑️ Reset Log'; }
                        if (log.activity_type === 'login')             { tagClass = 'tag-login';     typeLabel = '🔓 Login'; }
                        if (log.activity_type === 'logout')            { tagClass = 'tag-logout';    typeLabel = '🚪 Logout'; }
                        if (log.activity_type === 'login_failed')      { tagClass = 'tag-warning';   typeLabel = '⚠️ Gagal Login'; }
                        if (log.activity_type === 'create_user')       { tagClass = 'tag-success';   typeLabel = '👤 Buat User'; }
                        if (log.activity_type === 'delete_user')       { tagClass = 'tag-danger';    typeLabel = '🗑️ Hapus User'; }
                        if (log.activity_type === 'add_bookmark')      { tagClass = 'tag-success';   typeLabel = '⭐ +Bookmark'; }
                        if (log.activity_type === 'delete_bookmark')   { tagClass = 'tag-danger';    typeLabel = '⭐ -Bookmark'; }
                        if (log.activity_type === 'edit_bookmark')     { tagClass = 'tag-info';      typeLabel = '✏️ Catatan Book'; }
                        if (log.activity_type === 'view_file')         { tagClass = 'tag-info';      typeLabel = '📂 Lihat Berkas'; }
                        if (log.activity_type === 'edit_document')     { tagClass = 'tag-warning';   typeLabel = '✏️ Edit Baris'; }
                        if (log.activity_type === 'send_notification') { tagClass = 'tag-access';    typeLabel = '📢 Kirim Notif'; }

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
                            <td className="meta-cell text-center" style={{ width: '155px', fontSize: '12px' }}>
                              {new Date(log.created_at).toLocaleString('id-ID')}
                            </td>
                            <td className="text-center" style={{ width: '120px' }}>
                              {log.username ? (
                                <span style={{
                                  display: 'inline-flex', alignItems: 'center', gap: '4px',
                                  background: 'rgba(139,92,246,0.15)', border: '1px solid rgba(139,92,246,0.35)',
                                  color: '#c4b5fd', borderRadius: '6px', padding: '2px 8px',
                                  fontSize: '11px', fontWeight: '600', whiteSpace: 'nowrap'
                                }}>
                                  👤 {log.username}
                                </span>
                              ) : <span className="no-data">—</span>}
                            </td>
                            <td className="text-center" style={{ width: '110px' }}>
                              <span className={`activity-tag ${tagClass}`}>
                                {typeLabel}
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
                            <td className="text-center" style={{ width: '80px' }}>
                              <button
                                onClick={() => handleCopyLog(log)}
                                className={`btn-copy-log ${copiedLogId === log.id ? 'copied' : ''}`}
                                title="Salin rincian log ke clipboard"
                              >
                                {copiedLogId === log.id ? '✓ Tersalin' : '📋 Salin'}
                              </button>
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
          <div className="dashboard-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px', borderBottom: '1px solid var(--glass-border)', paddingBottom: '1.25rem', marginBottom: '1.75rem' }}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
              <h2 style={{ margin: 0, fontSize: '1.6rem', fontWeight: '700', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span>📊</span>
                <span style={{ background: 'linear-gradient(135deg, #fff 0%, #a78bfa 100%)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
                  Dashboard Statistik
                </span>
              </h2>
              <p style={{ margin: 0, fontSize: '0.88rem', color: 'var(--text-secondary)', lineHeight: '1.4', maxWidth: '550px' }}>
                Ringkasan aktivitas dan kondisi data aplikasi secara real-time.
              </p>
            </div>
            <div style={{ display: 'flex', gap: '10px' }}>
              <button className="btn-refresh" onClick={fetchDashboard} disabled={loadingDashboard} style={{ height: '40px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                {loadingDashboard ? '⏳ Memuat...' : '🔄 Refresh'}
              </button>
              {user?.role === 'admin' && (
                <button 
                  className="btn-delete-user" 
                  onClick={handleResetStats}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '0.6rem 1.25rem', fontSize: '0.88rem' }}
                >
                  🗑️ Reset Statistik
                </button>
              )}
            </div>
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

              {/* ANALITIK VISUAL (CHARTS & GRAPHS) */}
              <div className="dashboard-panel dashboard-panel--wide" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '2rem', padding: '1.75rem', marginBottom: '1.5rem' }}>
                
                {/* Chart 1: Tren Aktivitas 7 Hari Terakhir (SVG Line / Area Chart) */}
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <h3 style={{ margin: '0 0 1.25rem 0', fontSize: '1rem', color: '#fff', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    📈 Tren Aktivitas Mingguan (Total Log harian)
                  </h3>
                  {(() => {
                    // Process dailyActivity to get total counts per day
                    const dayMap = {};
                    dashboardStats.dailyActivity?.forEach(item => {
                      const dayStr = new Date(item.day).toLocaleDateString('id-ID', { weekday: 'short', day: 'numeric' });
                      dayMap[dayStr] = (dayMap[dayStr] || 0) + parseInt(item.count || 0);
                    });

                    const days = Object.keys(dayMap);
                    const values = Object.values(dayMap);
                    
                    if (days.length === 0) {
                      return <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', fontStyle: 'italic', margin: '2rem 0' }}>Belum ada aktivitas dalam 7 hari terakhir.</p>;
                    }

                    // Normalize to fit inside SVG viewbox (width: 400, height: 200)
                    const maxValue = Math.max(...values, 5); // default min height 5
                    const points = days.map((day, index) => {
                      const x = 40 + (index * 320) / (days.length - 1 || 1);
                      const y = 160 - (values[index] * 120) / maxValue;
                      return { x, y, day, value: values[index] };
                    });

                    const linePath = points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');
                    const areaPath = points.length > 0 
                      ? `${linePath} L ${points[points.length - 1].x} 160 L ${points[0].x} 160 Z`
                      : '';

                    return (
                      <div style={{ position: 'relative', width: '100%', height: '220px' }}>
                        <svg viewBox="0 0 400 200" style={{ width: '100%', height: '100%', overflow: 'visible' }}>
                          <defs>
                            <linearGradient id="areaGradient" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="0%" stopColor="var(--accent-secondary)" stopOpacity="0.45"/>
                              <stop offset="100%" stopColor="var(--accent-secondary)" stopOpacity="0.0"/>
                            </linearGradient>
                          </defs>

                          {/* Grid Lines */}
                          {[0, 0.25, 0.5, 0.75, 1].map((r, idx) => {
                            const y = 40 + r * 120;
                            const val = Math.round(maxValue * (1 - r));
                            return (
                              <g key={idx}>
                                <line x1="40" y1={y} x2="360" y2={y} stroke="rgba(255,255,255,0.06)" strokeDasharray="4 4" />
                                <text x="12" y={y + 4} fill="var(--text-muted)" fontSize="8" textAnchor="start">{val}</text>
                              </g>
                            );
                          })}

                          {/* Area under the line */}
                          {areaPath && <path d={areaPath} fill="url(#areaGradient)" />}

                          {/* Line path */}
                          {linePath && <path d={linePath} fill="none" stroke="var(--accent-secondary)" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />}

                          {/* Grid bottom axis */}
                          <line x1="40" y1="160" x2="360" y2="160" stroke="rgba(255,255,255,0.15)" strokeWidth="1" />

                          {/* Points and Labels */}
                          {points.map((p, idx) => (
                            <g key={idx}>
                              <circle cx={p.x} cy={p.y} r="4" fill="#fff" stroke="var(--accent-secondary)" strokeWidth="2.5" />
                              <text x={p.x} y={p.y - 10} fill="#fff" fontSize="8" fontWeight="600" textAnchor="middle">
                                {p.value}
                              </text>
                              <text x={p.x} y="178" fill="var(--text-secondary)" fontSize="8" textAnchor="middle">{p.day}</text>
                            </g>
                          ))}
                        </svg>
                      </div>
                    );
                  })()}
                </div>

                {/* Chart 2: Distribusi Kategori Bookmark (Donut Chart) */}
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <h3 style={{ margin: '0 0 1.25rem 0', fontSize: '1rem', color: '#fff', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    🍩 Distribusi Kategori Bookmark
                  </h3>
                  {(() => {
                    const groups = dashboardStats.bookmarkGroups || [];
                    const totalCount = groups.reduce((acc, g) => acc + parseInt(g.count || 0), 0);
                    
                    if (totalCount === 0) {
                      return <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', fontStyle: 'italic', margin: '2rem 0' }}>Belum ada bookmark tersimpan.</p>;
                    }

                    const colors = ['#8b5cf6', '#10b981', '#f59e0b', '#3b82f6', '#ec4899', '#34d399'];
                    
                    let accumulatedPercent = 0;
                    const segments = groups.map((g, i) => {
                      const count = parseInt(g.count || 0);
                      const percent = (count / totalCount) * 100;
                      const segmentColor = colors[i % colors.length];
                      const startPercent = accumulatedPercent;
                      accumulatedPercent += percent;
                      return {
                        name: g.group_name || 'Umum',
                        count,
                        percent,
                        color: segmentColor,
                        startPercent
                      };
                    });

                    const radius = 60;
                    const circumference = 2 * Math.PI * radius;

                    return (
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-around', flexWrap: 'wrap', gap: '20px', height: '220px' }}>
                        <div style={{ position: 'relative', width: '150px', height: '150px' }}>
                          <svg viewBox="0 0 160 160" style={{ width: '100%', height: '100%', transform: 'rotate(-90deg)' }}>
                            <circle cx="80" cy="80" r={radius} fill="transparent" stroke="rgba(255,255,255,0.03)" strokeWidth="18" />
                            {segments.map((seg, idx) => {
                              const strokeDash = (seg.percent / 100) * circumference;
                              const strokeOffset = circumference - (seg.startPercent / 100) * circumference;
                              return (
                                <circle
                                  key={idx}
                                  cx="80"
                                  cy="80"
                                  r={radius}
                                  fill="transparent"
                                  stroke={seg.color}
                                  strokeWidth="18"
                                  strokeDasharray={`${strokeDash} ${circumference - strokeDash}`}
                                  strokeDashoffset={strokeOffset}
                                  strokeLinecap="round"
                                  style={{ transition: 'stroke-dashoffset 0.5s ease-out' }}
                                />
                              );
                            })}
                          </svg>
                          <div style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center' }}>
                            <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#fff' }}>{totalCount}</span>
                            <span style={{ fontSize: '0.68rem', color: 'var(--text-muted)' }}>Bookmark</span>
                          </div>
                        </div>

                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '150px', overflowY: 'auto', paddingRight: '10px' }}>
                          {segments.map((seg, idx) => (
                            <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.82rem' }}>
                              <span style={{ display: 'inline-block', width: '10px', height: '10px', borderRadius: '50%', background: seg.color }}></span>
                              <span style={{ color: 'var(--text-secondary)', fontWeight: 500 }}>{seg.name}:</span>
                              <span style={{ color: '#fff', fontWeight: 600 }}>{seg.count} ({Math.round(seg.percent)}%)</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })()}
                </div>

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

      {/* TAB 6: KELOLA PENGGUNA (ADMIN ONLY) */}
      {activeTab === 'users' && user.role === 'admin' && (
        <main className="main-content">
          <div className="section-header-row">
            <div>
              <h2>👥 Kelola Pengguna (User Management)</h2>
              <p>Daftarkan dan kelola hak akses akun pengguna aplikasi pencarian dokumen.</p>
            </div>
          </div>

          <div className="users-management-container">
            <div className="users-grid-layout">
              {/* Left Column Wrapper */}
              <div className="users-left-column">
                {/* Form Tambah User */}
                <div className="users-form-panel">
                  <h3>{editingUser ? `✏️ Edit Akun: ${editingUser.username}` : '➕ Daftarkan Akun Baru'}</h3>
                  <form onSubmit={editingUser ? handleUpdateUser : handleCreateUser} className="users-create-form">
                    <div className="form-group">
                      <label>Username</label>
                      <input
                        type="text"
                        value={newUsername}
                        onChange={(e) => setNewUsername(e.target.value)}
                        placeholder="Contoh: operator_budi"
                        required
                      />
                    </div>
                    <div className="form-group">
                      <label>Nama Lengkap <span style={{ color: 'var(--text-muted)', fontWeight: 400, fontSize: '0.8rem' }}>(opsional)</span></label>
                      <input
                        type="text"
                        value={newFullName}
                        onChange={(e) => setNewFullName(e.target.value)}
                        placeholder="Contoh: Budi Santoso"
                      />
                    </div>
                    <div className="form-group">
                      <label>Password {editingUser && <span style={{ color: 'var(--text-muted)', fontWeight: 400, fontSize: '0.8rem' }}>(kosongkan jika tidak diubah)</span>}</label>
                      <input
                        type="password"
                        value={newPassword}
                        onChange={(e) => setNewPassword(e.target.value)}
                        placeholder={editingUser ? "Biarkan kosong jika tidak diubah" : "Masukkan password akun baru"}
                        required={!editingUser}
                      />
                    </div>
                    <div className="form-group">
                      <label>Role / Hak Akses</label>
                      <select
                        value={newUserRole}
                        onChange={(e) => setNewUserRole(e.target.value)}
                      >
                        <option value="viewer">Viewer (Hanya Cari)</option>
                        <option value="operator">Operator (Unggah + Cari)</option>
                        <option value="admin">Admin (Akses Penuh)</option>
                      </select>
                    </div>
                    <div style={{ display: 'flex', gap: '10px' }}>
                      <button type="submit" className="btn-submit-user" style={{ flex: 1 }}>
                        {editingUser ? 'Simpan Perubahan 💾' : 'Daftarkan Akun 👥'}
                      </button>
                      {editingUser && (
                        <button type="button" onClick={cancelEditUser} className="btn-submit-user" style={{ background: '#475569', color: '#fff', flex: 0.4 }}>
                          Batal
                        </button>
                      )}
                    </div>
                  </form>
                </div>

                {/* Broadcast Global Notification Form */}
                <div className="users-form-panel" style={{ marginTop: '1.5rem' }}>
                  <h3>📢 Kirim Notifikasi</h3>
                  <form onSubmit={handleBroadcastNotification} className="users-create-form">
                    <div className="form-group">
                      <label>Penerima Notifikasi</label>
                      <select
                        value={broadcastRecipient}
                        onChange={(e) => setBroadcastRecipient(e.target.value)}
                      >
                        <option value="all">📢 Semua Akun (Global Broadcast)</option>
                        {usersList.map(u => (
                          <option key={u.id} value={u.id}>
                            👤 {u.full_name ? `${u.full_name} (${u.username})` : u.username} — {u.role.toUpperCase()}
                          </option>
                        ))}
                      </select>
                    </div>
                    
                    <div className="form-group">
                      <label>Pesan Notifikasi</label>
                      <textarea
                        value={broadcastMessage}
                        onChange={(e) => setBroadcastMessage(e.target.value)}
                        placeholder={broadcastRecipient === 'all' ? "Masukkan pengumuman penting untuk semua pengguna..." : "Masukkan pesan khusus untuk pengguna ini..."}
                        required
                      />
                    </div>
                    <button type="submit" className="btn-submit-user btn-broadcast">
                      {broadcastRecipient === 'all' ? 'Kirim ke Semua Akun 🚀' : 'Kirim ke Akun Dituju 🚀'}
                    </button>
                  </form>
                </div>
              </div>

              {/* Tabel Daftar User */}
              <div className="users-list-panel">
                <h3>📋 Daftar Pengguna Terdaftar</h3>
                {loadingUsers ? (
                  <div className="search-loader"><div className="spinner"></div><p>Memuat daftar pengguna...</p></div>
                ) : usersList.length === 0 ? (
                  <p>Tidak ada pengguna terdaftar.</p>
                ) : (
                  <div className="users-table-wrapper">
                    <table className="users-table">
                      <thead>
                        <tr>
                          <th>Username</th>
                          <th>Nama Lengkap</th>
                          <th>Role</th>
                          <th>Tanggal Dibuat</th>
                          <th>Aksi</th>
                        </tr>
                      </thead>
                      <tbody>
                        {usersList.map((u) => (
                          <tr key={u.id}>
                            <td className="user-td-username">
                              <strong>{u.username}</strong>
                              {u.username === user.username && <span className="current-user-badge">Anda</span>}
                            </td>
                            <td style={{ color: u.full_name ? 'var(--text-primary)' : 'var(--text-muted)', fontStyle: u.full_name ? 'normal' : 'italic' }}>
                              {u.full_name || '—'}
                            </td>
                            <td>
                              <span className={`role-badge role-${u.role}`}>
                                {u.role.toUpperCase()}
                              </span>
                            </td>
                            <td>
                              {new Date(u.created_at).toLocaleDateString('id-ID')}
                            </td>
                            <td>
                              <button
                                type="button"
                                className="btn-view-bookmarks"
                                onClick={() => fetchUserBookmarks(u)}
                                title={`Lihat & Kelola Bookmark milik ${u.username}`}
                                style={{ marginRight: '8px' }}
                              >
                                ⭐ Bookmark
                              </button>
                              <button
                                type="button"
                                className="btn-edit-user"
                                onClick={() => startEditUser(u)}
                                title={`Edit akun "${u.username}"`}
                                style={{ marginRight: '8px' }}
                              >
                                ✏️ Edit
                              </button>
                              <button
                                type="button"
                                className="btn-delete-user"
                                onClick={() => handleDeleteUser(u.id, u.username)}
                                disabled={u.username === user.username}
                                title={u.username === user.username ? "Anda tidak dapat menghapus akun sendiri" : "Hapus Akun"}
                              >
                                🗑️ Hapus
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>

            {/* Kelola Bookmark Per User Panel */}
            {selectedUserForBookmarks && (
              <div className="users-bookmarks-panel" style={{ marginTop: '2rem' }}>
                <div className="bookmarks-panel-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem', borderBottom: '1px solid var(--glass-border)', paddingBottom: '0.5rem' }}>
                  <h3>⭐ Kelola Bookmark Akun: <span style={{ color: 'var(--accent-secondary)' }}>{selectedUserForBookmarks.username}</span> ({userBookmarksList.length})</h3>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    {userBookmarksList.length > 0 && (
                      <button 
                        onClick={handleClearUserBookmarks}
                        className="btn-delete-user"
                        style={{ padding: '0.45rem 0.9rem', background: 'rgba(239, 68, 68, 0.15)', border: '1px solid rgba(239, 68, 68, 0.3)', color: '#f87171', borderRadius: '6px', fontSize: '0.82rem', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px' }}
                      >
                        🗑️ Hapus Semua Bookmark
                      </button>
                    )}
                    <button 
                      className="btn-close-bookmarks"
                      onClick={() => { setSelectedUserForBookmarks(null); setEditingBookmark(null); }}
                      style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', fontSize: '1.1rem', display: 'flex', alignItems: 'center', gap: '4px' }}
                    >
                      ❌ Tutup
                    </button>
                  </div>
                </div>

                {editingBookmark && (
                  <div className="bookmark-edit-form-wrapper" style={{ background: 'var(--bg-tertiary)', border: '1px solid var(--glass-border)', borderRadius: '10px', padding: '1.25rem', marginBottom: '1.5rem' }}>
                    <h4 style={{ marginBottom: '1rem' }}>✏️ Edit Bookmark Baris #{editingBookmark.row_number} - {editingBookmark.filename}</h4>
                    <form onSubmit={handleUpdateBookmark}>
                      <div className="form-group" style={{ marginBottom: '1rem' }}>
                        <label style={{ display: 'block', fontWeight: 'bold', fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Nama Kategori / Grup Bookmark</label>
                        <input
                          type="text"
                          value={editingBookmarkGroup}
                          onChange={(e) => setEditingBookmarkGroup(e.target.value)}
                          placeholder="Contoh: Keuangan, Penting, Umum, dll."
                          style={{ width: '100%', padding: '0.6rem 0.85rem', background: 'var(--bg-secondary)', border: '1px solid var(--glass-border)', borderRadius: '6px', color: '#fff', fontSize: '0.85rem', marginBottom: '1rem' }}
                        />
                      </div>
                      <div className="form-group" style={{ marginBottom: '1rem' }}>
                        <label style={{ display: 'block', fontWeight: 'bold', fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Catatan / Keterangan Bookmark</label>
                        <textarea
                          value={editingBookmarkNotes}
                          onChange={(e) => setEditingBookmarkNotes(e.target.value)}
                          placeholder="Tambahkan catatan khusus untuk bookmark ini..."
                          style={{ width: '100%', padding: '0.6rem', background: 'var(--bg-secondary)', border: '1px solid var(--glass-border)', borderRadius: '6px', color: '#fff', fontSize: '0.85rem', minHeight: '60px', fontFamily: 'inherit', resize: 'vertical' }}
                        />
                      </div>
                      
                      <div className="form-group" style={{ marginBottom: '1.25rem' }}>
                        <label style={{ display: 'block', fontWeight: 'bold', fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Konten Baris Dokumen (Key-Values)</label>
                        <div style={{ maxHeight: '200px', overflowY: 'auto', border: '1px solid var(--glass-border)', borderRadius: '6px', padding: '0.75rem', background: 'var(--bg-secondary)' }}>
                          {Object.entries(editingBookmarkRowData).map(([key, val]) => (
                            <div key={key} style={{ display: 'flex', alignItems: 'center', marginBottom: '0.5rem', gap: '8px' }}>
                              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)', width: '150px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }} title={key}>{key}:</span>
                              <input
                                type="text"
                                value={String(val)}
                                onChange={(e) => {
                                  const updatedVal = e.target.value;
                                  setEditingBookmarkRowData(prev => ({ ...prev, [key]: updatedVal }));
                                }}
                                style={{ flex: 1, padding: '0.35rem 0.6rem', background: 'var(--bg-tertiary)', border: '1px solid var(--glass-border)', borderRadius: '4px', color: '#fff', fontSize: '0.82rem' }}
                              />
                            </div>
                          ))}
                        </div>
                      </div>

                      <div style={{ display: 'flex', gap: '10px' }}>
                        <button type="submit" className="btn-submit-user" style={{ padding: '0.5rem 1rem' }}>
                          Simpan Perubahan
                        </button>
                        <button 
                          type="button" 
                          className="btn-delete-user" 
                          onClick={() => setEditingBookmark(null)}
                          style={{ padding: '0.5rem 1rem', background: 'rgba(255, 255, 255, 0.05)', borderColor: 'var(--glass-border)', color: 'var(--text-secondary)' }}
                        >
                          Batal
                        </button>
                      </div>
                    </form>
                  </div>
                )}

                {loadingUserBookmarks ? (
                  <div className="search-loader"><div className="spinner"></div><p>Memuat bookmark pengguna...</p></div>
                ) : userBookmarksList.length === 0 ? (
                  <div className="empty-state" style={{ padding: '2rem' }}>
                    <p style={{ color: 'var(--text-muted)', fontStyle: 'italic' }}>Pengguna ini belum mem-bookmark dokumen apa pun.</p>
                  </div>
                ) : (
                  <div className="users-table-wrapper">
                    <table className="users-table">
                      <thead>
                        <tr>
                          <th>Berkas / File</th>
                          <th>Sheet</th>
                          <th>No. Baris</th>
                          <th>Data / Konten</th>
                          <th>Catatan</th>
                          <th>Aksi</th>
                        </tr>
                      </thead>
                      <tbody>
                        {userBookmarksList.map((b) => (
                          <tr key={b.id}>
                            <td>📄 {b.filename}</td>
                            <td>📁 {b.sheet_name}</td>
                            <td>#{b.row_number}</td>
                            <td>
                              <div className="row-data-preview-container" style={{ maxHeight: '100px', overflowY: 'auto', fontSize: '0.8rem' }}>
                                {Object.entries(b.row_data || {}).map(([k, v]) => (
                                  <div key={k} style={{ marginBottom: '2px' }}>
                                    <strong style={{ color: 'var(--text-muted)' }}>{k}:</strong> {String(v)}
                                  </div>
                                ))}
                              </div>
                            </td>
                            <td>
                              {b.notes ? (
                                <span style={{ fontSize: '0.85rem' }}>{b.notes}</span>
                              ) : (
                                <span style={{ color: 'var(--text-muted)', fontStyle: 'italic', fontSize: '0.8rem' }}>(Kosong)</span>
                              )}
                            </td>
                            <td>
                              <button
                                type="button"
                                className="btn-edit-bookmark"
                                onClick={() => {
                                  setEditingBookmark(b);
                                  setEditingBookmarkNotes(b.notes || '');
                                  setEditingBookmarkGroup(b.group_name || 'Umum');
                                  setEditingBookmarkRowData({ ...b.row_data });
                                }}
                                title="Edit catatan & data baris bookmark"
                                style={{ marginRight: '8px' }}
                              >
                                ✏️ Edit
                              </button>
                              <button
                                type="button"
                                className="btn-delete-user"
                                onClick={() => handleDeleteUserBookmark(b.id)}
                                title="Hapus bookmark dari pengguna ini"
                              >
                                🗑️ Hapus
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}
          </div>
        </main>
      )}
      
      <footer className="app-footer-bar">
        <p>© 2026 SpreadSheet Finder — Didukung oleh React, Express, dan PostgreSQL</p>
      </footer>

      {/* Modal Dialog Pemilihan Kategori / Grup Bookmark */}
      {showBookmarkGroupModal && (
        <div className="modal-overlay" style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0, 0, 0, 0.65)',
          backdropFilter: 'blur(8px)',
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          zIndex: 9999,
          padding: '20px'
        }}>
          <div className="modal-box" style={{
            background: 'var(--bg-secondary)',
            border: '1px solid var(--glass-border)',
            borderRadius: '16px',
            padding: '2rem',
            width: '100%',
            maxWidth: '480px',
            boxShadow: '0 20px 40px rgba(0, 0, 0, 0.4)',
            color: '#fff',
            animation: 'fadeIn 0.3s ease-out'
          }}>
            <h3 style={{ margin: '0 0 1rem 0', fontSize: '1.25rem', display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--accent-secondary)' }}>
              ⭐ Simpan ke Kategori Bookmark
            </h3>
            <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '1.5rem', lineHeight: '1.4' }}>
              Pilih kategori yang sudah ada atau buat kategori baru agar bookmark data Anda terpisah dan terorganisir dengan rapi.
            </p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginBottom: '2rem' }}>
              {/* Option 1: Pilih kategori yang sudah ada */}
              {(() => {
                const existingGroups = new Set();
                bookmarksList.forEach(b => {
                  if (b.group_name) existingGroups.add(b.group_name);
                });
                const groupsList = Array.from(existingGroups).filter(g => g !== 'Semua');

                return (
                  <>
                    {groupsList.length > 0 && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: '10px', cursor: 'pointer', background: 'rgba(255,255,255,0.02)', padding: '0.75rem', borderRadius: '8px', border: bookmarkGroupModalSelectedOption === 'existing' ? '1px solid var(--accent-secondary)' : '1px solid rgba(255,255,255,0.05)', transition: 'all 0.2s ease' }}>
                        <input
                          type="radio"
                          name="bookmark-option"
                          checked={bookmarkGroupModalSelectedOption === 'existing'}
                          onChange={() => setBookmarkGroupModalSelectedOption('existing')}
                          style={{ marginTop: '3px', cursor: 'pointer' }}
                        />
                        <div style={{ flex: 1 }}>
                          <span style={{ fontSize: '0.88rem', fontWeight: 600, display: 'block', marginBottom: '6px' }}>Pilih Kategori Terdaftar:</span>
                          <select
                            value={bookmarkGroupModalSelectedExisting}
                            onChange={(e) => {
                              setBookmarkGroupModalSelectedExisting(e.target.value);
                              setBookmarkGroupModalSelectedOption('existing');
                            }}
                            disabled={bookmarkGroupModalSelectedOption !== 'existing'}
                            style={{
                              width: '100%',
                              padding: '0.5rem',
                              borderRadius: '6px',
                              background: 'var(--bg-tertiary)',
                              border: '1px solid var(--glass-border)',
                              color: '#fff',
                              fontSize: '0.85rem',
                              cursor: bookmarkGroupModalSelectedOption === 'existing' ? 'pointer' : 'default'
                            }}
                          >
                            {groupsList.map(g => (
                              <option key={g} value={g}>{g}</option>
                            ))}
                          </select>
                        </div>
                      </label>
                    )}

                    {/* Option 2: Buat kategori baru */}
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: '10px', cursor: 'pointer', background: 'rgba(255,255,255,0.02)', padding: '0.75rem', borderRadius: '8px', border: bookmarkGroupModalSelectedOption === 'new' ? '1px solid var(--accent-secondary)' : '1px solid rgba(255,255,255,0.05)', transition: 'all 0.2s ease' }}>
                      <input
                        type="radio"
                        name="bookmark-option"
                        checked={bookmarkGroupModalSelectedOption === 'new'}
                        onChange={() => setBookmarkGroupModalSelectedOption('new')}
                        style={{ marginTop: '3px', cursor: 'pointer' }}
                      />
                      <div style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.88rem', fontWeight: 600, display: 'block', marginBottom: '6px' }}>Buat Kategori Baru:</span>
                        <input
                          type="text"
                          placeholder="Masukkan nama kategori baru (contoh: Logistik)"
                          value={bookmarkGroupModalNewInput}
                          onChange={(e) => {
                            setBookmarkGroupModalNewInput(e.target.value);
                            setBookmarkGroupModalSelectedOption('new');
                          }}
                          disabled={bookmarkGroupModalSelectedOption !== 'new'}
                          style={{
                            width: '100%',
                            padding: '0.55rem 0.75rem',
                            borderRadius: '6px',
                            background: 'var(--bg-tertiary)',
                            border: '1px solid var(--glass-border)',
                            color: '#fff',
                            fontSize: '0.85rem'
                          }}
                        />
                      </div>
                    </label>
                  </>
                );
              })()}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px' }}>
              <button
                type="button"
                className="btn-delete-user"
                onClick={() => setShowBookmarkGroupModal(false)}
                style={{
                  padding: '0.55rem 1.25rem',
                  background: 'rgba(255, 255, 255, 0.05)',
                  border: '1px solid var(--glass-border)',
                  color: 'var(--text-secondary)',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontSize: '0.88rem'
                }}
              >
                Batal
              </button>
              <button
                type="button"
                className="btn-submit-user"
                onClick={handleConfirmBookmarkToggle}
                style={{
                  padding: '0.55rem 1.5rem',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontSize: '0.88rem',
                  fontWeight: 600
                }}
              >
                Simpan Bookmark 💾
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Deleting Progress Overlay */}
      {deleting && (
        <div className="delete-progress-overlay">
          <div className="delete-progress-card">
            <div className="delete-spinner-wrapper">
              <div className="delete-spinner"></div>
              <span className="delete-trash-icon">🗑️</span>
            </div>
            <h3>Menghapus Data</h3>
            <p className="delete-status-text">{deleteStatusMessage}</p>
            <div className="delete-progress-bar-container">
              <div className="delete-progress-bar" style={{ width: `${deleteProgress}%` }}></div>
            </div>
            <span className="delete-progress-percent">{deleteProgress}%</span>
          </div>
        </div>
      )}

      {/* Downloading Progress Overlay */}
      {downloading && (
        <div className="download-progress-overlay">
          <div className="download-progress-card">
            <div className="download-spinner-wrapper">
              <div className="download-spinner"></div>
              <span className="download-file-icon">📥</span>
            </div>
            <h3>Mengunduh Dokumen Excel</h3>
            <p className="download-status-text">{downloadStatusMessage}</p>
            <div className="download-progress-bar-container">
              <div className="download-progress-bar" style={{ width: `${downloadProgress}%` }}></div>
            </div>
            <span className="download-progress-percent">{downloadProgress}%</span>
          </div>
        </div>
      )}

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
      {/* AI Chatbot Floating Button & Drawer */}
      <button
        type="button"
        className="floating-chat-toggle"
        onClick={() => setIsChatOpen(!isChatOpen)}
        title="Buka AI Document Chatbot"
      >
        <span className="chat-toggle-icon">🤖</span>
        <span className="chat-toggle-text">Chat AI</span>
      </button>

      {isChatOpen && (
        <div className="chat-drawer-container glass-panel">
          <div className="chat-drawer-header">
            <div className="chat-header-title">
              <span className="chat-ai-avatar">🤖</span>
              <div>
                <h4>SpreadSheet AI Assistant</h4>
                <p className="chat-status-text">
                  <span className="online-indicator"></span> Siap Menjawab Pertanyaan
                </p>
              </div>
            </div>
            <button
              type="button"
              className="chat-close-btn"
              onClick={() => setIsChatOpen(false)}
              title="Tutup Chat"
            >
              ✕
            </button>
          </div>

          {/* Document Scope & AI Engine Selector Bar */}
          <div className="chat-scope-bar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
            <div style={{ flex: '1 1 180px', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span className="chat-scope-label" style={{ fontSize: '0.78rem', whiteSpace: 'nowrap' }}>🎯 Dokumen:</span>
              <select
                className="chat-scope-select"
                style={{ flex: 1, padding: '0.35rem 0.6rem', fontSize: '0.78rem' }}
                value={chatSelectedFileId}
                onChange={(e) => setChatSelectedFileId(e.target.value)}
                title="Pilih apakah AI mencari di semua dokumen atau 1 dokumen spesifik"
              >
                <option value="all">🌐 Semua Dokumen ({(filesList || []).length} File)</option>
                {(filesList || []).map(f => (
                  <option key={f.id} value={f.id}>
                    📄 {f.filename}
                  </option>
                ))}
              </select>
            </div>

            <div style={{ flex: '1 1 150px', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span className="chat-scope-label" style={{ fontSize: '0.78rem', whiteSpace: 'nowrap' }}>🤖 Mesin AI:</span>
              <select
                className="chat-scope-select"
                style={{ flex: 1, padding: '0.35rem 0.6rem', fontSize: '0.78rem' }}
                value={chatModelMode}
                onChange={(e) => setChatModelMode(e.target.value)}
                title="Pilih mesin AI yang digunakan untuk menjawab"
              >
                <option value="auto">🤖 Otomatis (Prioritas Gemini)</option>
                <option value="gemini">✨ Google Gemini AI (Cloud)</option>
                <option value="local">⚡ Local AI (Offline / Cepat)</option>
              </select>
            </div>
          </div>

          <div className="chat-messages-list">
            {chatMessages.map(msg => (
              <div key={msg.id} className={`chat-message-item ${msg.sender === 'user' ? 'user-msg' : 'ai-msg'}`}>
                <div className="chat-bubble">
                  {msg.sender === 'ai' && (
                    <div className="chat-bubble-badge">
                      {msg.methodUsed === 'gemini' ? '✨ Gemini AI' : '⚡ Local AI'}
                    </div>
                  )}
                  <div className="chat-bubble-text">
                    {msg.text.split('\n').map((line, idx) => (
                      <p key={idx}>{line}</p>
                    ))}
                  </div>

                  {msg.sources && msg.sources.length > 0 && (
                    <div className="chat-sources-container">
                      <span className="sources-label">📄 Referensi Berkas:</span>
                      <div className="sources-chips">
                        {msg.sources.map(src => (
                          <button
                            key={src.fileId}
                            type="button"
                            className="source-chip"
                            onClick={() => handleSourceClick(src.fileId, src.filename, src.rowNumber, msg.userQuery, src.sheetName)}
                            title={`Klik untuk melihat berkas di tabel${src.sheetName ? ' · Sheet: ' + src.sheetName : ''}`}
                          >
                            🔍 {src.filename} (#{src.rowNumber}{src.sheetName && src.sheetName !== 'Sheet1' ? ` · ${src.sheetName}` : ''})
                          </button>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="chat-bubble-footer" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '6px' }}>
                    <span className="chat-bubble-time">{msg.timestamp}</span>
                    {msg.sender === 'ai' && (
                      <button
                        type="button"
                        className="chat-copy-btn"
                        onClick={() => {
                          navigator.clipboard.writeText(msg.text);
                          showToast('Tersalin!', 'Jawaban AI berhasil disalin ke clipboard.', 'success');
                        }}
                        title="Salin jawaban AI"
                        style={{ background: 'transparent', border: 'none', color: 'var(--accent-primary)', fontSize: '0.72rem', cursor: 'pointer', opacity: 0.85 }}
                      >
                        📋 Salin Jawaban
                      </button>
                    )}
                  </div>
                </div>
              </div>
            ))}

            {isChatLoading && (
              <div className="chat-message-item ai-msg">
                <div className="chat-bubble thinking-bubble">
                  <span className="chat-ai-avatar">🤖</span>
                  <div className="typing-dots">
                    <span>.</span><span>.</span><span>.</span>
                  </div>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>Mencari data & menganalisis...</span>
                </div>
              </div>
            )}
          </div>

          {/* Quick Suggestion Chips */}
          <div className="chat-suggestions-bar">
            <span className="suggestions-title">💡 Rekomendasi Pertanyaan:</span>
            <div className="suggestions-scroll">
              <button type="button" className="suggestion-chip" onClick={() => handleSendChatMessage('Berapa total dokumen tahun 2025?')}>
                📊 Dokumen Tahun 2025
              </button>
              <button type="button" className="suggestion-chip" onClick={() => handleSendChatMessage('Tampilkan data divisi umum')}>
                🏛️ Divisi Umum
              </button>
              <button type="button" className="suggestion-chip" onClick={() => handleSendChatMessage('Rangkum berkas P011')}>
                📑 Berkas P011
              </button>
            </div>
          </div>

          <form className="chat-input-form" onSubmit={(e) => { e.preventDefault(); handleSendChatMessage(); }}>
            <input
              type="text"
              className="chat-input-field"
              placeholder="Ketik pertanyaan tentang dokumen Excel..."
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              disabled={isChatLoading}
            />
            <button
              type="submit"
              className="chat-send-btn"
              disabled={!chatInput.trim() || isChatLoading}
            >
              🚀
            </button>
          </form>
        </div>
      )}
    </div>
  );
}

function LoginScreen({ setToken, setUser, isDarkMode, setIsDarkMode, showToast }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [loginError, setLoginError] = useState('');

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoginError('');
    if (!username.trim() || !password) {
      setLoginError('Username dan Password wajib diisi.');
      return;
    }
    setLoading(true);
    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: username.trim(), password })
      });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Login gagal.');
      }
      localStorage.setItem('app_token', data.token);
      localStorage.setItem('app_user', JSON.stringify(data.user));
      setToken(data.token);
      setUser(data.user);
      showToast('🔑 Login Berhasil', `Selamat datang kembali, ${data.user.username}!`, 'success');
    } catch (err) {
      setLoginError(err.message || 'Username atau password salah.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={`login-container ${isDarkMode ? 'dark-theme' : 'light-theme'}`}>
      <div className="login-card-container">
        <div className="login-glass-card">
          <div className="login-logo-section">
            <div className="login-logo-icon">📊</div>
            <h2>SpreadSheet Finder</h2>
            <p>Silakan masuk untuk mengakses database dokumen</p>
          </div>
          
          <form className="login-form" onSubmit={handleSubmit}>
            <div className="form-group">
              <label>Username</label>
              <div className="input-with-icon">
                <span className="input-icon">👤</span>
                <input 
                  type="text" 
                  value={username}
                  onChange={(e) => { setUsername(e.target.value); setLoginError(''); }}
                  placeholder="Masukkan username Anda"
                  required
                  style={loginError ? { borderColor: '#f87171' } : {}}
                />
              </div>
            </div>
            
            <div className="form-group">
              <label>Password</label>
              <div className="input-with-icon">
                <span className="input-icon">🔒</span>
                <input 
                  type="password" 
                  value={password}
                  onChange={(e) => { setPassword(e.target.value); setLoginError(''); }}
                  placeholder="Masukkan password Anda"
                  required
                  style={loginError ? { borderColor: '#f87171' } : {}}
                />
              </div>
            </div>
            
            {loginError && (
              <div className="login-error-banner">
                <span className="login-error-icon">⚠️</span>
                <span>{loginError}</span>
              </div>
            )}

            <button type="submit" className="login-submit-btn" disabled={loading}>
              {loading ? 'Memverifikasi...' : 'Masuk ke Aplikasi 🚀'}
            </button>
          </form>

          {/* Login Footer Removed */}
        </div>
      </div>
    </div>
  );
}

// Reusable grab-to-scroll component to allow smooth click-and-drag scrolling on wide tables
function DragScrollContainer({ children, className }) {
  const containerRef = useRef(null);
  const [isDown, setIsDown] = useState(false);
  const [startX, setStartX] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  
  // States for scroll arrows visibility
  const [canScrollLeft, setCanScrollLeft] = useState(false);
  const [canScrollRight, setCanScrollRight] = useState(false);

  const checkScroll = () => {
    const el = containerRef.current;
    if (el) {
      setCanScrollLeft(el.scrollLeft > 5);
      setCanScrollRight(el.scrollLeft < el.scrollWidth - el.clientWidth - 5);
    }
  };

  useEffect(() => {
    const el = containerRef.current;
    if (el) {
      checkScroll();
      
      const resizeObserver = new ResizeObserver(() => {
        checkScroll();
      });
      resizeObserver.observe(el);
      
      const mutationObserver = new MutationObserver(checkScroll);
      mutationObserver.observe(el, { childList: true, subtree: true });

      el.addEventListener('scroll', checkScroll);
      window.addEventListener('resize', checkScroll);
      
      const interval = setInterval(checkScroll, 300);

      return () => {
        resizeObserver.disconnect();
        mutationObserver.disconnect();
        el.removeEventListener('scroll', checkScroll);
        window.removeEventListener('resize', checkScroll);
        clearInterval(interval);
      };
    }
  }, [children]);

  const onMouseDown = (e) => {
    if (e.button !== 0) return;
    const targetTag = e.target.tagName.toLowerCase();
    if (targetTag === 'button' || targetTag === 'input' || targetTag === 'select' || targetTag === 'a' || e.target.closest('.bookmark-star-btn') || e.target.closest('.btn-edit-bookmark')) {
      return;
    }
    
    setIsDown(true);
    setStartX(e.pageX - containerRef.current.offsetLeft);
    setScrollLeft(containerRef.current.scrollLeft);
  };

  const onMouseLeave = () => {
    setIsDown(false);
  };

  const onMouseUp = () => {
    setIsDown(false);
  };

  const onMouseMove = (e) => {
    if (!isDown) return;
    e.preventDefault();
    const x = e.pageX - containerRef.current.offsetLeft;
    const walk = (x - startX) * 1.5;
    containerRef.current.scrollLeft = scrollLeft - walk;
  };

  const scrollByAmount = (amount) => {
    const el = containerRef.current;
    if (el) {
      el.scrollTo({
        left: el.scrollLeft + amount,
        behavior: 'smooth'
      });
    }
  };

  return (
    <div className="drag-scroll-wrapper" style={{ position: 'relative', width: '100%' }}>

      <div
        ref={containerRef}
        className={className}
        onMouseDown={onMouseDown}
        onMouseLeave={onMouseLeave}
        onMouseUp={onMouseUp}
        onMouseMove={onMouseMove}
        style={{
          cursor: isDown ? 'grabbing' : 'grab',
          userSelect: isDown ? 'none' : 'auto'
        }}
      >
        {children}
      </div>
    </div>
  );
}

export default App;
