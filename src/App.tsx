import React, { useState, useEffect, useMemo } from 'react';
import {
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot,
  writeBatch
} from 'firebase/firestore';
import {
  GoogleAuthProvider,
  signInWithPopup,
  signOut,
  onAuthStateChanged,
  type User
} from 'firebase/auth';
import { db, auth, testConnection, handleFirestoreError, OperationType } from './firebase.ts';
import type { Application } from './types.ts';
import {
  SHEETS_SCOPES,
  setSheetsToken,
  getSheetsToken,
  createEventSpreadsheet,
  getSpreadsheetDetails,
  getExistingTicketCodes,
  appendApplicantRows,
  formatApplicationRow,
  type SheetsConfig
} from './googleSheets.ts';

const SLOTS = ['1회차 · 평일 오전', '2회차 · 평일 저녁', '3회차 · 주말 오후'] as const;
const ORGANIZER_EMAIL = 'jina.jung01@gmail.com';
const MY_APP_KEY = 'event_my_application_id_v2';

const MSG = {
  nameEmpty: '이름을 입력하세요.',
  nameChar: '이름은 한글 또는 영어로 입력하세요.',
  nameMax: '최대 10자입니다.',
  phone: '연락처를 000-0000-0000 형식으로 입력하세요. 예: 010-1234-5678',
  phoneDup: '이미 신청된 연락처입니다.',
  email: '올바르지 않은 이메일 형식입니다.',
  emailDup: '이미 신청된 이메일입니다.',
  slot: '희망 일정을 선택하세요.',
  memoMax: '최대 100자입니다.',
  agree: '개인정보 수집에 동의해야 신청할 수 있습니다.',
};

/* ───────── 개인정보 마스킹 ───────── */
function maskName(n: string) {
  const c = [...(n || '')];
  if (c.length <= 1) return '*';
  if (c.length === 2) return c[0] + '*';
  return c[0] + '*'.repeat(c.length - 2) + c[c.length - 1];
}

function maskPhone(p: string) {
  return (p || '').replace(/^(\d{3})-\d{4}-(\d{4})$/, '$1-****-$2');
}

function maskEmail(e: string) {
  const [id, dom] = (e || '').split('@');
  if (!dom) return '***';
  return (id.length <= 2 ? id[0] : id.slice(0, 2)) + '***@' + dom;
}

function fmtDate(iso: string) {
  try {
    return new Date(iso).toLocaleString('ko-KR', { dateStyle: 'short', timeStyle: 'short' });
  } catch {
    return '';
  }
}

function generateTicketCode(existingList: Application[]): string {
  let c = '';
  do {
    c = Math.random().toString(36).slice(2, 8).toUpperCase().padEnd(6, '0');
  } while (existingList.some(x => x.code === c));
  return c;
}

function newId(): string {
  return 'app_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

export default function App() {
  // Path Routing: '/' vs '/admin'
  const [route, setRoute] = useState<'apply' | 'admin'>(() => {
    const path = window.location.pathname;
    const hash = window.location.hash;
    return path === '/admin' || path.startsWith('/admin/') || hash === '#admin' ? 'admin' : 'apply';
  });

  const [applications, setApplications] = useState<Application[]>([]);
  const [isLiveConnected, setIsLiveConnected] = useState<boolean>(false);
  const [dbError, setDbError] = useState<string | null>(null);

  // Auth State
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [authLoading, setAuthLoading] = useState<boolean>(true);
  const [authError, setAuthError] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState<boolean>(false);

  // Google Sheets Config State (shared in Firestore and cached locally)
  const [sheetsConfig, setSheetsConfig] = useState<SheetsConfig | null>(() => {
    try {
      const saved = localStorage.getItem('event_sheets_config_v1');
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  });
  const [isSyncingSheets, setIsSyncingSheets] = useState<boolean>(false);
  const [sheetsSyncStatus, setSheetsSyncStatus] = useState<string | null>(null);
  const [isLinkSheetModalOpen, setIsLinkSheetModalOpen] = useState<boolean>(false);
  const [manualSheetInput, setManualSheetInput] = useState<string>('');

  // Participant State
  const [myAppId, setMyAppId] = useState<string | null>(() => {
    try {
      return localStorage.getItem(MY_APP_KEY) || null;
    } catch {
      return null;
    }
  });
  const [submittedApp, setSubmittedApp] = useState<Application | null>(null);
  const [isEditingMine, setIsEditingMine] = useState<boolean>(false);
  const [justSubmitted, setJustSubmitted] = useState<boolean>(false);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);

  // Form State (신청서)
  const [formName, setFormName] = useState('');
  const [formPhone, setFormPhone] = useState('');
  const [formEmail, setFormEmail] = useState('');
  const [formSlot, setFormSlot] = useState('');
  const [formMemo, setFormMemo] = useState('');
  const [formAgree, setFormAgree] = useState(false);

  // Admin Search State
  const [searchQuery, setSearchQuery] = useState('');

  // Dialog State
  const [popupMessage, setPopupMessage] = useState<string | null>(null);
  const [confirmDialog, setConfirmDialog] = useState<{
    message: string;
    onConfirm: () => void;
  } | null>(null);

  // Detail View Modal (Organizer)
  const [viewAppId, setViewAppId] = useState<string | null>(null);
  const [viewRevealed, setViewRevealed] = useState<boolean>(false);

  // Admin Add/Edit Modal
  const [adminEditApp, setAdminEditApp] = useState<Application | null>(null);
  const [isAdminEditOpen, setIsAdminEditOpen] = useState(false);
  const [editFormName, setEditFormName] = useState('');
  const [editFormPhone, setEditFormPhone] = useState('');
  const [editFormEmail, setEditFormEmail] = useState('');
  const [editFormSlot, setEditFormSlot] = useState('');
  const [editFormMemo, setEditFormMemo] = useState('');
  const [editFormAgree, setEditFormAgree] = useState(true);

  // Listen to browser forward/back buttons & hash changes
  useEffect(() => {
    const handlePopState = () => {
      const path = window.location.pathname;
      const hash = window.location.hash;
      setRoute(path === '/admin' || path.startsWith('/admin/') || hash === '#admin' ? 'admin' : 'apply');
    };
    window.addEventListener('popstate', handlePopState);
    window.addEventListener('hashchange', handlePopState);
    return () => {
      window.removeEventListener('popstate', handlePopState);
      window.removeEventListener('hashchange', handlePopState);
    };
  }, []);

  const navigateTo = (to: 'apply' | 'admin') => {
    const targetUrl = to === 'admin' ? '/admin' : '/';
    window.history.pushState({}, '', targetUrl);
    setRoute(to);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  // 1. Listen for Auth State Changes
  useEffect(() => {
    const unsubscribeAuth = onAuthStateChanged(auth, user => {
      setCurrentUser(user);
      setAuthLoading(false);
    });
    return () => unsubscribeAuth();
  }, []);

  // 2. Initial Connection Test & Real-time Snapshot for Applications
  useEffect(() => {
    testConnection().then(ok => {
      setIsLiveConnected(ok);
    });

    const unsubscribe = onSnapshot(
      collection(db, 'applications'),
      snapshot => {
        const list: Application[] = [];
        snapshot.forEach(docSnap => {
          const data = docSnap.data();
          list.push({
            id: docSnap.id,
            code: data.code || '',
            name: data.name || '',
            phone: data.phone || '',
            email: data.email || '',
            slot: data.slot || '',
            memo: data.memo || '',
            agree: data.agree ?? true,
            createdAt: data.createdAt || new Date().toISOString(),
            updatedAt: data.updatedAt || new Date().toISOString()
          });
        });
        // Sort chronologically
        list.sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)));
        setApplications(list);
        setIsLiveConnected(true);
        setDbError(null);
      },
      error => {
        setIsLiveConnected(false);
        setDbError('Firestore 연결에 실패했습니다: ' + error.message);
        handleFirestoreError(error, OperationType.GET, 'applications');
      }
    );

    return () => unsubscribe();
  }, []);

  // 3. Listen for Google Sheets settings document in Firestore
  useEffect(() => {
    const unsubSettings = onSnapshot(
      doc(db, 'settings', 'sheets'),
      snap => {
        if (snap.exists()) {
          const cfg = snap.data() as SheetsConfig;
          setSheetsConfig(cfg);
          try {
            localStorage.setItem('event_sheets_config_v1', JSON.stringify(cfg));
          } catch {}
        }
      },
      err => {
        console.warn('Sheets settings snapshot note:', err.message);
      }
    );
    return () => unsubSettings();
  }, []);

  // Check if current user is authorized organizer
  const isOrganizer = useMemo(() => {
    if (!currentUser?.email) return false;
    return currentUser.email.toLowerCase() === ORGANIZER_EMAIL.toLowerCase();
  }, [currentUser]);

  // 4. Auto-append new applicants to Google Sheets when organizer is active and sheet is configured
  useEffect(() => {
    const token = getSheetsToken();
    if (!token || !sheetsConfig?.spreadsheetId || !isOrganizer || applications.length === 0) {
      return;
    }

    let isCancelled = false;
    const syncNew = async () => {
      try {
        const existingCodes = await getExistingTicketCodes(
          token,
          sheetsConfig.spreadsheetId,
          sheetsConfig.sheetName
        );
        if (isCancelled) return;

        const unSynced = applications.filter(a => !existingCodes.has(a.code));
        if (unSynced.length > 0) {
          const rowsToAppend = unSynced.map((app, idx) =>
            formatApplicationRow(app, existingCodes.size + idx + 1)
          );
          await appendApplicantRows(
            token,
            sheetsConfig.spreadsheetId,
            sheetsConfig.sheetName,
            rowsToAppend
          );
          setSheetsSyncStatus(`새 참가자 ${unSynced.length}명이 구글 시트에 자동 누적되었습니다.`);
          setTimeout(() => setSheetsSyncStatus(null), 4000);
        }
      } catch (err) {
        console.warn('Auto-sync to sheets check failed:', err);
      }
    };

    syncNew();
    return () => {
      isCancelled = true;
    };
  }, [applications, sheetsConfig, isOrganizer]);

  // Find my application
  const myApplication = useMemo(() => {
    if (submittedApp && (!myAppId || submittedApp.id === myAppId)) {
      const liveMatch = applications.find(a => a.id === submittedApp.id);
      return liveMatch || submittedApp;
    }
    if (!myAppId) return null;
    return applications.find(a => a.id === myAppId) || null;
  }, [applications, myAppId, submittedApp]);

  // When my application changes or starts editing, sync form
  useEffect(() => {
    if (isEditingMine && myApplication) {
      setFormName(myApplication.name);
      setFormPhone(myApplication.phone);
      setFormEmail(myApplication.email);
      setFormSlot(myApplication.slot);
      setFormMemo(myApplication.memo || '');
      setFormAgree(myApplication.agree);
    }
  }, [isEditingMine, myApplication]);

  // Google Login Handler (requests Sheets and Drive.file scopes)
  const handleGoogleSignIn = async () => {
    setAuthError(null);
    setIsLoggingIn(true);
    try {
      const provider = new GoogleAuthProvider();
      SHEETS_SCOPES.forEach(s => provider.addScope(s));
      provider.setCustomParameters({ prompt: 'select_account' });
      const result = await signInWithPopup(auth, provider);
      const credential = GoogleAuthProvider.credentialFromResult(result);
      if (credential?.accessToken) {
        setSheetsToken(credential.accessToken);
      }
    } catch (err: any) {
      if (err?.code === 'auth/popup-closed-by-user') {
        setIsLoggingIn(false);
        return;
      }
      console.error('Google Sign-in failed:', err);
      setAuthError('Google 로그인 중 오류가 발생했습니다: ' + (err?.message || '잠시 후 다시 시도해 주세요.'));
    } finally {
      setIsLoggingIn(false);
    }
  };

  // Google Sign-Out Handler
  const handleSignOut = async () => {
    try {
      await signOut(auth);
      setSheetsToken(null);
      setSearchQuery('');
      setViewAppId(null);
      setIsAdminEditOpen(false);
    } catch (err) {
      console.error('Sign-out failed:', err);
    }
  };

  // Ensure OAuth token is valid in memory; request popup if missing
  const ensureSheetsToken = async (): Promise<string> => {
    const existing = getSheetsToken();
    if (existing) return existing;

    const provider = new GoogleAuthProvider();
    SHEETS_SCOPES.forEach(s => provider.addScope(s));
    provider.setCustomParameters({ prompt: 'consent' });
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Google 계정에서 스프레드시트 접근 토큰을 획득하지 못했습니다. 다시 로그인해 주세요.');
    }
    setSheetsToken(credential.accessToken);
    return credential.accessToken;
  };

  // Google Sheets: Create New Spreadsheet & Link
  const handleCreateNewSheet = () => {
    setConfirmDialog({
      message: "Google Drive에 새 스프레드시트 '[이벤트 참가 신청 명단] 실시간 접수 현황'을 생성하고 연동할까요?",
      onConfirm: async () => {
        setIsSyncingSheets(true);
        try {
          const token = await ensureSheetsToken();
          const result = await createEventSpreadsheet(token, '이벤트 참가 신청 명단 (실시간)');
          // Sync any currently existing applications immediately
          if (applications.length > 0) {
            const rows = applications.map((a, i) => formatApplicationRow(a, i + 1));
            await appendApplicantRows(token, result.spreadsheetId, result.sheetName, rows);
          }
          const newConfig: SheetsConfig = {
            spreadsheetId: result.spreadsheetId,
            spreadsheetUrl: result.spreadsheetUrl,
            sheetName: result.sheetName,
            autoSync: true,
            updatedAt: new Date().toISOString()
          };
          setSheetsConfig(newConfig);
          try {
            localStorage.setItem('event_sheets_config_v1', JSON.stringify(newConfig));
          } catch {}
          try {
            await setDoc(doc(db, 'settings', 'sheets'), newConfig);
          } catch (e) {
            console.warn('Firestore settings write note:', e);
          }
          setSheetsSyncStatus('새 구글 스프레드시트가 성공적으로 생성되고 연동되었습니다!');
          setTimeout(() => setSheetsSyncStatus(null), 5000);
        } catch (err: any) {
          console.error('Failed to create spreadsheet:', err);
          setPopupMessage(err?.message || '스프레드시트 생성 실패: 잠시 후 다시 시도해 주세요.');
        } finally {
          setIsSyncingSheets(false);
        }
      }
    });
  };

  // Google Sheets: Link Existing Spreadsheet
  const handleLinkExistingSheet = async () => {
    let id = manualSheetInput.trim();
    const urlMatch = id.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (urlMatch) id = urlMatch[1];
    if (!id) {
      setPopupMessage('올바른 스프레드시트 ID 또는 공유 URL을 입력하세요.');
      return;
    }
    setIsSyncingSheets(true);
    try {
      const token = await ensureSheetsToken();
      const details = await getSpreadsheetDetails(token, id);
      const chosenSheet = details.sheetNames[0] || 'Sheet1';
      const newConfig: SheetsConfig = {
        spreadsheetId: id,
        spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${id}/edit`,
        sheetName: chosenSheet,
        autoSync: true,
        updatedAt: new Date().toISOString()
      };
      setSheetsConfig(newConfig);
      try {
        localStorage.setItem('event_sheets_config_v1', JSON.stringify(newConfig));
      } catch {}
      try {
        await setDoc(doc(db, 'settings', 'sheets'), newConfig);
      } catch (e) {
        console.warn('Firestore settings write note:', e);
      }
      setIsLinkSheetModalOpen(false);
      setManualSheetInput('');
      setSheetsSyncStatus(`'${details.title}' 스프레드시트가 성공적으로 연동되었습니다.`);
      setTimeout(() => setSheetsSyncStatus(null), 4000);
    } catch (err: any) {
      console.error('Failed to link sheet:', err);
      setPopupMessage(err?.message || '스프레드시트 연결 실패: ID와 공유 권한을 확인해 주세요.');
    } finally {
      setIsSyncingSheets(false);
    }
  };

  // Google Sheets: Manual Full Sync
  const handleManualSync = () => {
    if (!sheetsConfig?.spreadsheetId) {
      setPopupMessage('연동된 구글 스프레드시트가 없습니다.');
      return;
    }
    setConfirmDialog({
      message: `현재 등록된 모든 참가자(${applications.length}명)의 데이터를 Google 스프레드시트에 누적 동기화하시겠습니까?`,
      onConfirm: async () => {
        setIsSyncingSheets(true);
        try {
          const token = await ensureSheetsToken();
          const existingCodes = await getExistingTicketCodes(
            token,
            sheetsConfig.spreadsheetId,
            sheetsConfig.sheetName
          );
          const unSynced = applications.filter(a => !existingCodes.has(a.code));
          if (unSynced.length === 0) {
            setPopupMessage('이미 모든 참가자 데이터가 구글 스프레드시트에 동기화되어 있습니다.');
            return;
          }
          const rowsToAppend = unSynced.map((app, idx) =>
            formatApplicationRow(app, existingCodes.size + idx + 1)
          );
          await appendApplicantRows(
            token,
            sheetsConfig.spreadsheetId,
            sheetsConfig.sheetName,
            rowsToAppend
          );
          setSheetsSyncStatus(`참가자 ${unSynced.length}명의 데이터가 구글 시트에 성공적으로 동기화되었습니다.`);
          setTimeout(() => setSheetsSyncStatus(null), 4000);
        } catch (err: any) {
          console.error('Manual sync failed:', err);
          setPopupMessage(err?.message || '구글 시트 동기화 중 오류가 발생했습니다.');
        } finally {
          setIsSyncingSheets(false);
        }
      }
    });
  };

  // Google Sheets: Unlink Sheet
  const handleUnlinkSheet = () => {
    setConfirmDialog({
      message: 'Google 스프레드시트 연동을 해제할까요? (스프레드시트 원본 파일은 삭제되지 않습니다.)',
      onConfirm: async () => {
        setSheetsConfig(null);
        try {
          localStorage.removeItem('event_sheets_config_v1');
        } catch {}
        try {
          await deleteDoc(doc(db, 'settings', 'sheets'));
        } catch (err) {
          console.warn('Firestore settings delete note:', err);
        }
        setSheetsSyncStatus('스프레드시트 연동이 해제되었습니다.');
        setTimeout(() => setSheetsSyncStatus(null), 3000);
      }
    });
  };

  // Helper to sanitize name
  const handleNameChange = (val: string, setter: (v: string) => void) => {
    const cleaned = val.replace(/[^a-zA-Z가-힣ㄱ-ㅎㅏ-ㅣ]/g, '');
    const sliced = [...cleaned].slice(0, 10).join('');
    setter(sliced);
    if ([...cleaned].length > 10) {
      setPopupMessage(MSG.nameMax);
    }
  };

  // Helper to format phone
  const handlePhoneChange = (val: string, setter: (v: string) => void) => {
    const digits = val.replace(/\D/g, '').slice(0, 11);
    const formatted =
      digits.length > 7
        ? `${digits.slice(0, 3)}-${digits.slice(3, 7)}-${digits.slice(7)}`
        : digits.length > 3
        ? `${digits.slice(0, 3)}-${digits.slice(3)}`
        : digits;
    setter(formatted);
  };

  // Validate form
  const validateApplication = (
    data: { name: string; phone: string; email: string; slot: string; memo: string; agree: boolean },
    excludeId?: string | null
  ): string | null => {
    if (!data.name.trim()) return MSG.nameEmpty;
    if (!/^[a-zA-Z가-힣]{1,10}$/.test(data.name.trim())) return MSG.nameChar;
    if (!/^\d{3}-\d{4}-\d{4}$/.test(data.phone.trim())) return MSG.phone;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email.trim())) return MSG.email;
    if (!SLOTS.includes(data.slot as any)) return MSG.slot;
    if ([...data.memo].length > 100) return MSG.memoMax;
    if (!data.agree) return MSG.agree;

    // Duplicate checks
    const others = applications.filter(x => x.id !== excludeId);
    if (others.some(x => x.phone === data.phone.trim())) return MSG.phoneDup;
    if (others.some(x => x.email.toLowerCase() === data.email.trim().toLowerCase())) return MSG.emailDup;

    return null;
  };

  // Submit Application (Participant)
  const handleSubmitApplication = async () => {
    const errorMsg = validateApplication(
      { name: formName, phone: formPhone, email: formEmail, slot: formSlot, memo: formMemo, agree: formAgree },
      isEditingMine ? myApplication?.id : null
    );
    if (errorMsg) {
      setPopupMessage(errorMsg);
      return;
    }

    setIsSubmitting(true);
    try {
      const now = new Date().toISOString();
      const targetId = isEditingMine && myApplication ? myApplication.id : newId();
      const code = isEditingMine && myApplication ? myApplication.code : generateTicketCode(applications);

      const docData = {
        code,
        name: formName.trim(),
        phone: formPhone.trim(),
        email: formEmail.trim(),
        slot: formSlot,
        memo: formMemo.trim(),
        agree: formAgree,
        createdAt: isEditingMine && myApplication ? myApplication.createdAt : now,
        updatedAt: now
      };

      const record: Application = {
        id: targetId,
        ...docData
      };

      // Set timeout race to prevent UI hanging under offline or connection issues
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error('네트워크 응답 시간이 초과되었습니다. 다시 시도해 주세요.')), 8000)
      );

      await Promise.race([
        setDoc(doc(db, 'applications', targetId), docData),
        timeoutPromise
      ]);

      setSubmittedApp(record);
      setMyAppId(targetId);
      try {
        localStorage.setItem(MY_APP_KEY, targetId);
      } catch {}

      setIsEditingMine(false);
      setJustSubmitted(true);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } catch (err: any) {
      console.error('Submission failed:', err);
      setPopupMessage(err?.message || '신청서 저장 중 오류가 발생했습니다. 네트워크 상태를 확인해 주세요.');
    } finally {
      setIsSubmitting(false);
    }
  };

  // Withdraw / Cancel Application (Participant)
  const handleWithdrawMine = () => {
    if (!myApplication) return;
    setConfirmDialog({
      message: '신청을 취소할까요? 입력한 내용이 실시간 명단에서 영구 삭제됩니다.',
      onConfirm: async () => {
        try {
          await deleteDoc(doc(db, 'applications', myApplication.id));
          setMyAppId(null);
          setSubmittedApp(null);
          try {
            localStorage.removeItem(MY_APP_KEY);
          } catch {}
          setFormName('');
          setFormPhone('');
          setFormEmail('');
          setFormSlot('');
          setFormMemo('');
          setFormAgree(false);
          setIsEditingMine(false);
        } catch (err) {
          handleFirestoreError(err, OperationType.DELETE, 'applications/' + myApplication.id);
        }
      }
    });
  };

  // Start new application
  const handleStartAnother = () => {
    setMyAppId(null);
    setSubmittedApp(null);
    try {
      localStorage.removeItem(MY_APP_KEY);
    } catch {}
    setFormName('');
    setFormPhone('');
    setFormEmail('');
    setFormSlot('');
    setFormMemo('');
    setFormAgree(false);
    setIsEditingMine(false);
  };

  // Filtered applications for admin search
  const filteredApplications = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    const qDigits = q.replace(/\D/g, '');
    if (!q) return applications;
    return applications.filter(app => {
      return (
        app.name.toLowerCase().includes(q) ||
        app.email.toLowerCase().includes(q) ||
        app.code.toLowerCase().includes(q) ||
        (qDigits.length >= 4 && app.phone.replace(/\D/g, '').includes(qDigits))
      );
    });
  }, [applications, searchQuery]);

  // Admin Save (Add or Edit)
  const handleAdminSave = async () => {
    const errorMsg = validateApplication(
      {
        name: editFormName,
        phone: editFormPhone,
        email: editFormEmail,
        slot: editFormSlot,
        memo: editFormMemo,
        agree: editFormAgree
      },
      adminEditApp ? adminEditApp.id : null
    );
    if (errorMsg) {
      setPopupMessage(errorMsg);
      return;
    }

    try {
      const now = new Date().toISOString();
      const targetId = adminEditApp ? adminEditApp.id : newId();
      const code = adminEditApp ? adminEditApp.code : generateTicketCode(applications);

      const docData = {
        code,
        name: editFormName.trim(),
        phone: editFormPhone.trim(),
        email: editFormEmail.trim(),
        slot: editFormSlot,
        memo: editFormMemo.trim(),
        agree: editFormAgree,
        createdAt: adminEditApp ? adminEditApp.createdAt : now,
        updatedAt: now
      };

      await setDoc(doc(db, 'applications', targetId), docData);
      setIsAdminEditOpen(false);
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, 'applications/' + (adminEditApp?.id || 'new'));
    }
  };

  // Admin Delete One
  const handleAdminDeleteOne = (app: Application) => {
    setConfirmDialog({
      message: `${maskName(app.name)}님(${app.code})의 신청을 실시간 데이터베이스에서 완전히 삭제할까요?`,
      onConfirm: async () => {
        try {
          await deleteDoc(doc(db, 'applications', app.id));
        } catch (err) {
          handleFirestoreError(err, OperationType.DELETE, 'applications/' + app.id);
        }
      }
    });
  };

  // Admin Clear All
  const handleAdminClearAll = () => {
    if (applications.length === 0) return;
    setConfirmDialog({
      message: `현재 등록된 참가자 ${applications.length}명의 모든 신청 정보를 Firestore에서 전체 삭제할까요? 이 작업은 되돌릴 수 없습니다.`,
      onConfirm: async () => {
        try {
          const batch = writeBatch(db);
          applications.forEach(app => {
            batch.delete(doc(db, 'applications', app.id));
          });
          await batch.commit();
        } catch (err) {
          handleFirestoreError(err, OperationType.DELETE, 'applications');
        }
      }
    });
  };

  // CSV Export with BOM
  const handleExportCsv = () => {
    if (!applications.length) {
      setPopupMessage('다운로드할 신청자가 없습니다.');
      return;
    }
    setConfirmDialog({
      message: `개인정보(이름·연락처·이메일)가 포함된 실시간 참가자 명단(${applications.length}건)을 CSV로 다운로드할까요?`,
      onConfirm: () => {
        const head = ['No', '접수 번호', '이름', '연락처', '이메일', '희망 일정', '남기실 말', '신청 시각', '최종 수정'];
        const rows = applications.map((d, i) => [
          i + 1,
          d.code,
          d.name,
          d.phone,
          d.email,
          d.slot,
          d.memo || '',
          fmtDate(d.createdAt),
          fmtDate(d.updatedAt)
        ]);
        const csvCell = (v: any) => {
          let t = String(v ?? '');
          if (/^[=+\-@\t\r]/.test(t)) t = "'" + t;
          return /[",\r\n]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t;
        };
        const csvContent =
          '\uFEFF' + [head, ...rows].map(row => row.map(csvCell).join(',')).join('\r\n');

        const now = new Date();
        const pad = (n: number) => String(n).padStart(2, '0');
        const filename = `참가자명단_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(
          now.getHours()
        )}${pad(now.getMinutes())}.csv`;

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }
    });
  };

  // Open Admin Edit
  const openAdminEdit = (target: Application | null) => {
    setAdminEditApp(target);
    if (target) {
      setEditFormName(target.name);
      setEditFormPhone(target.phone);
      setEditFormEmail(target.email);
      setEditFormSlot(target.slot);
      setEditFormMemo(target.memo || '');
      setEditFormAgree(target.agree);
    } else {
      setEditFormName('');
      setEditFormPhone('');
      setEditFormEmail('');
      setEditFormSlot('');
      setEditFormMemo('');
      setEditFormAgree(true);
    }
    setIsAdminEditOpen(true);
  };

  // Active view target for modal
  const viewTargetApp = useMemo(() => {
    if (!viewAppId) return null;
    return applications.find(a => a.id === viewAppId) || null;
  }, [applications, viewAppId]);

  return (
    <main className={route === 'admin' && isOrganizer ? 'wide' : ''}>
      {dbError && (
        <div className="notice" style={{ marginBottom: '16px', borderColor: 'var(--danger)', color: 'var(--danger)' }}>
          {dbError}
        </div>
      )}

      {/* ───────────────────────────────────────────────────────────
          1. 참가 신청 페이지 (루트 경로: '/')
         ─────────────────────────────────────────────────────────── */}
      {route === 'apply' && (
        <div>
          <div style={{ marginBottom: '28px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h1>참가 신청</h1>
              {isLiveConnected && (
                <span className="live-pill" title="Firestore 실시간 연결 완료">
                  <span className="live-dot" />
                  접수 중
                </span>
              )}
            </div>
            <p className="lead" style={{ margin: '6px 0 0' }}>
              {myApplication && !isEditingMine
                ? '참가 신청이 정상적으로 접수되었습니다.'
                : isEditingMine
                ? '신청 내용을 수정하신 후 저장해 주세요.'
                : '아래 내용을 입력하고 이벤트 참가를 신청하세요.'}
            </p>
          </div>

          {myApplication && !isEditingMine ? (
            /* 접수증 카드 */
            <section aria-live="polite">
              <div className={`ticket ${justSubmitted ? 'new' : ''}`}>
                <div className="body">
                  <h2>신청이 접수되었습니다</h2>
                  <dl>
                    <dt>이름</dt>
                    <dd>{maskName(myApplication.name)}</dd>
                    <dt>연락처</dt>
                    <dd>{maskPhone(myApplication.phone)}</dd>
                    <dt>이메일</dt>
                    <dd>{maskEmail(myApplication.email)}</dd>
                    <dt>희망 일정</dt>
                    <dd>{myApplication.slot}</dd>
                    <dt>신청 시각</dt>
                    <dd>{fmtDate(myApplication.updatedAt)}</dd>
                  </dl>
                </div>
                <div className="stub">
                  <small>접수 번호</small>
                  <strong>{myApplication.code}</strong>
                </div>
              </div>

              <div className="row">
                <button
                  type="button"
                  className="ghost"
                  onClick={() => setIsEditingMine(true)}
                >
                  내용 수정
                </button>
                <button
                  type="button"
                  className="danger"
                  onClick={handleWithdrawMine}
                >
                  신청 취소
                </button>
              </div>
              <div className="row">
                <button
                  type="button"
                  className="primary"
                  onClick={handleStartAnother}
                >
                  새 신청서 작성
                </button>
              </div>
            </section>
          ) : (
            /* 참가 신청 폼 */
            <div className="card">
              <div className="field">
                <label htmlFor="user-name">
                  이름<span className="req">*</span>
                </label>
                <input
                  id="user-name"
                  type="text"
                  value={formName}
                  onChange={e => handleNameChange(e.target.value, setFormName)}
                  placeholder="한글 또는 영어, 10자 이내"
                  autoComplete="off"
                />
              </div>

              <div className="field">
                <label htmlFor="user-phone">
                  연락처<span className="req">*</span>
                </label>
                <input
                  id="user-phone"
                  type="tel"
                  inputMode="numeric"
                  value={formPhone}
                  onChange={e => handlePhoneChange(e.target.value, setFormPhone)}
                  placeholder="010-0000-0000"
                  maxLength={13}
                  autoComplete="off"
                />
              </div>

              <div className="field">
                <label htmlFor="user-email">
                  이메일<span className="req">*</span>
                </label>
                <input
                  id="user-email"
                  type="email"
                  value={formEmail}
                  onChange={e => setFormEmail(e.target.value.replace(/\s/g, ''))}
                  placeholder="example@email.com"
                  maxLength={80}
                  autoComplete="off"
                />
              </div>

              <div className="field">
                <label htmlFor="user-slot">
                  희망 일정<span className="req">*</span>
                </label>
                <select
                  id="user-slot"
                  value={formSlot}
                  onChange={e => setFormSlot(e.target.value)}
                >
                  <option value="">선택하세요</option>
                  {SLOTS.map(s => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>

              <div className="field">
                <label htmlFor="user-memo">남기실 말</label>
                <textarea
                  id="user-memo"
                  value={formMemo}
                  onChange={e => {
                    const c = [...e.target.value];
                    if (c.length > 100) {
                      setFormMemo(c.slice(0, 100).join(''));
                      setPopupMessage(MSG.memoMax);
                    } else {
                      setFormMemo(e.target.value);
                    }
                  }}
                  placeholder="문의하실 내용이나 전달사항이 있으면 100자 이내로 적어 주세요."
                />
                <div className="count">{[...formMemo].length} / 100</div>
              </div>

              <div className="field">
                <label className="check">
                  <input
                    type="checkbox"
                    checked={formAgree}
                    onChange={e => setFormAgree(e.target.checked)}
                  />
                  신청 확인 연락을 위해 개인정보를 수집하는 데 동의합니다.
                  <span className="req">*</span>
                </label>
              </div>

              <button
                type="button"
                className="primary"
                disabled={isSubmitting}
                onClick={handleSubmitApplication}
              >
                {isSubmitting
                  ? '저장 중...'
                  : isEditingMine
                  ? '수정 내용 저장'
                  : '신청하기'}
              </button>

              {isEditingMine && (
                <div className="row">
                  <button
                    type="button"
                    className="ghost"
                    onClick={() => setIsEditingMine(false)}
                  >
                    수정 취소
                  </button>
                </div>
              )}
            </div>
          )}

          {/* 신청자에게 방해되지 않는 하단 주최자 관리 페이지 링크 */}
          <footer style={{ marginTop: '48px', paddingTop: '20px', borderTop: '1px solid var(--line)', textAlign: 'center', fontSize: '0.8rem', color: 'var(--muted)' }}>
            <span>© 이벤트 참가 신청 · </span>
            <a
              href="/admin"
              onClick={e => {
                e.preventDefault();
                navigateTo('admin');
              }}
              style={{ color: 'var(--muted)', textDecoration: 'underline', cursor: 'pointer' }}
            >
              주최자 관리자(/admin)
            </a>
          </footer>
        </div>
      )}

      {/* ───────────────────────────────────────────────────────────
          2. 주최자 관리자 페이지 (경로: '/admin')
         ─────────────────────────────────────────────────────────── */}
      {route === 'admin' && (
        <div>
          {/* 상단 네비게이션: 참가 신청 페이지로 돌아가기 */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
            <button
              type="button"
              className="ghost small"
              onClick={() => navigateTo('apply')}
              style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
            >
              ← 참가 신청 페이지로 이동
            </button>

            {isLiveConnected && (
              <span className="live-pill" title="Firestore 실시간 동기화 중">
                <span className="live-dot" />
                실시간 연동됨
              </span>
            )}
          </div>

          {authLoading ? (
            <div className="card lock">
              <p style={{ color: 'var(--muted)' }}>주최자 인증 상태를 확인하고 있습니다...</p>
            </div>
          ) : !currentUser ? (
            /* Google 로그인 화면 */
            <div className="card lock">
              <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '16px' }}>
                <svg width="44" height="44" viewBox="0 0 24 24">
                  <path
                    fill="#4285F4"
                    d="M23.745 12.27c0-.7-.06-1.4-.19-2.07H12v4.51h6.6c-.29 1.52-1.14 2.82-2.4 3.68v3.05h3.88c2.27-2.09 3.665-5.17 3.665-9.17z"
                  />
                  <path
                    fill="#34A853"
                    d="M12 24c3.24 0 5.95-1.08 7.93-2.91l-3.88-3.05c-1.08.72-2.45 1.16-4.05 1.16-3.12 0-5.77-2.1-6.72-4.93H1.25v3.15C3.26 21.36 7.36 24 12 24z"
                  />
                  <path
                    fill="#FBBC05"
                    d="M5.28 14.27c-.25-.72-.38-1.49-.38-2.27s.13-1.55.38-2.27V6.58H1.25C.45 8.18 0 9.98 0 12s.45 3.82 1.25 5.42l4.03-3.15z"
                  />
                  <path
                    fill="#EA4335"
                    d="M12 4.75c1.77 0 3.35.61 4.6 1.8l3.42-3.42C17.95 1.19 15.24 0 12 0 7.36 0 3.26 2.64 1.25 6.58l4.03 3.15c.95-2.83 3.6-4.98 6.72-4.98z"
                  />
                </svg>
              </div>
              <h2 style={{ margin: '0 0 8px', fontSize: '1.25rem' }}>주최자 전용 관리자 (/admin)</h2>
              <p style={{ color: 'var(--muted)', margin: '0 0 20px', fontSize: '0.95rem', lineHeight: 1.5 }}>
                주최자 Google 계정(<strong>{ORGANIZER_EMAIL}</strong>)으로 로그인하면 실시간 참가자 명단 확인 및 <strong>Google 스프레드시트 자동 누적 연동</strong>을 사용하실 수 있습니다.
              </p>
              {authError && (
                <div style={{ color: 'var(--danger)', marginBottom: '16px', fontSize: '0.9rem' }}>
                  {authError}
                </div>
              )}
              <button
                type="button"
                className="primary"
                disabled={isLoggingIn}
                onClick={handleGoogleSignIn}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '10px',
                  padding: '13px 20px'
                }}
              >
                <span>{isLoggingIn ? '로그인 창 여는 중...' : 'Google 계정으로 로그인 (시트 권한 포함)'}</span>
              </button>
            </div>
          ) : !isOrganizer ? (
            /* 권한 없는 계정으로 로그인된 경우 */
            <div className="card lock">
              <div className="icon" aria-hidden="true" style={{ fontSize: '2.5rem', marginBottom: '8px' }}>
                🚫
              </div>
              <h2 style={{ margin: '0 0 8px', fontSize: '1.25rem', color: 'var(--danger)' }}>
                접근 권한이 없습니다
              </h2>
              <p style={{ margin: '0 0 8px', fontSize: '0.95rem' }}>
                현재 로그인된 계정: <strong>{currentUser.email}</strong>
              </p>
              <p style={{ color: 'var(--muted)', margin: '0 0 20px', fontSize: '0.9rem', lineHeight: 1.5 }}>
                이 관리 페이지는 이벤트 주최자 계정(<strong>{ORGANIZER_EMAIL}</strong>)으로만 접근하실 수 있습니다.
              </p>
              <button
                type="button"
                className="primary"
                onClick={handleSignOut}
              >
                다른 Google 계정으로 로그인
              </button>
            </div>
          ) : (
            /* 승인된 주최자 실시간 신청자 명단 관리 화면 */
            <section className="admin">
              <div className="adminHead">
                <div>
                  <h2 style={{ fontSize: '1.4rem' }}>이벤트 참가자 관리 (/admin)</h2>
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginTop: '4px' }}>
                    <span className="live-pill">
                      <span className="live-dot" />
                      실시간 Firestore 동기화
                    </span>
                    <span style={{ fontSize: '0.85rem', color: 'var(--muted)' }}>
                      총 {applications.length}명 신청 완료
                    </span>
                  </div>
                </div>

                {/* 주최자 프로필 및 로그아웃 버튼 */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', background: 'var(--surface)', padding: '6px 12px', borderRadius: '10px', border: '1px solid var(--line)' }}>
                  {currentUser.photoURL && (
                    <img
                      src={currentUser.photoURL}
                      alt="주최자 프로필"
                      style={{ width: '32px', height: '32px', borderRadius: '50%', objectFit: 'cover' }}
                    />
                  )}
                  <div style={{ textAlign: 'left', lineHeight: 1.3 }}>
                    <div style={{ fontWeight: 600, fontSize: '0.85rem' }}>
                      {currentUser.displayName || '주최자'}
                    </div>
                    <div style={{ color: 'var(--muted)', fontSize: '0.75rem' }}>
                      {currentUser.email}
                    </div>
                  </div>
                  <button
                    type="button"
                    className="ghost small"
                    onClick={handleSignOut}
                    title="로그아웃"
                    style={{ marginLeft: '4px', padding: '5px 10px' }}
                  >
                    로그아웃
                  </button>
                </div>
              </div>

              {/* ───────────────────────────────────────────────────────────
                  Google Sheets 실시간 연동 관리 카드
                 ─────────────────────────────────────────────────────────── */}
              <div
                className="card"
                style={{
                  marginTop: '16px',
                  marginBottom: '16px',
                  border: '1px solid #107C41',
                  background: 'var(--surface)',
                  padding: '16px 20px'
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <div style={{ width: '32px', height: '32px', background: '#E6F4EA', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      <svg width="20" height="20" viewBox="0 0 24 24">
                        <path fill="#0F9D58" d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm2 14H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z" />
                      </svg>
                    </div>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: '1rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        Google 스프레드시트 실시간 연동
                        {sheetsConfig && (
                          <span style={{ fontSize: '0.75rem', background: '#E6F4EA', color: '#137333', padding: '2px 8px', borderRadius: '999px', fontWeight: 600 }}>
                            자동 누적 활성화됨
                          </span>
                        )}
                      </div>
                      <div style={{ fontSize: '0.85rem', color: 'var(--muted)' }}>
                        {sheetsConfig
                          ? '새로운 참가자 신청 접수 시 Google 스프레드시트에 행으로 실시간 자동 누적됩니다.'
                          : 'Google 스프레드시트를 연동하여 신청자가 접수될 때마다 즉시 행으로 자동 누적하세요.'}
                      </div>
                    </div>
                  </div>

                  {sheetsConfig ? (
                    <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                      <a
                        href={sheetsConfig.spreadsheetUrl}
                        target="_blank"
                        rel="noreferrer"
                        className="primary small"
                        style={{
                          textDecoration: 'none',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '6px',
                          background: '#0F9D58',
                          width: 'auto'
                        }}
                      >
                        시트 열기 ↗
                      </a>
                      <button
                        type="button"
                        className="ghost small"
                        disabled={isSyncingSheets}
                        onClick={handleManualSync}
                      >
                        {isSyncingSheets ? '동기화 중...' : '전체 동기화'}
                      </button>
                      <button
                        type="button"
                        className="danger small"
                        onClick={handleUnlinkSheet}
                      >
                        연동 해제
                      </button>
                    </div>
                  ) : (
                    <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                      <button
                        type="button"
                        className="primary small"
                        style={{ width: 'auto', background: '#0F9D58' }}
                        disabled={isSyncingSheets}
                        onClick={handleCreateNewSheet}
                      >
                        {isSyncingSheets ? '시트 생성 중...' : '+ 새 스프레드시트 생성 및 연동'}
                      </button>
                      <button
                        type="button"
                        className="ghost small"
                        onClick={() => setIsLinkSheetModalOpen(true)}
                      >
                        기존 시트 연결
                      </button>
                    </div>
                  )}
                </div>

                {sheetsSyncStatus && (
                  <div style={{ marginTop: '10px', fontSize: '0.85rem', color: '#137333', fontWeight: 600 }}>
                    ✓ {sheetsSyncStatus}
                  </div>
                )}
              </div>

              {/* 회차별 현황 요약 칩 */}
              <div className="crud" style={{ marginTop: '14px' }}>
                {SLOTS.map(s => {
                  const cnt = applications.filter(a => a.slot === s).length;
                  return (
                    <span key={s}>
                      {s}: <strong>{cnt}명</strong>
                    </span>
                  );
                })}
              </div>

              {/* 툴바: 검색, 추가, CSV */}
              <div className="toolbar">
                <input
                  type="search"
                  placeholder="이름·연락처·이메일·접수 번호 검색"
                  value={searchQuery}
                  onChange={e => setSearchQuery(e.target.value)}
                />
                <button
                  type="button"
                  className="primary"
                  style={{ width: 'auto' }}
                  onClick={() => openAdminEdit(null)}
                >
                  + 신청 추가
                </button>
                <button
                  type="button"
                  className="ghost"
                  disabled={applications.length === 0}
                  onClick={handleExportCsv}
                >
                  ⬇ CSV 전체 다운로드
                </button>
              </div>

              {/* 테이블 */}
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>No</th>
                      <th>접수 번호</th>
                      <th>이름</th>
                      <th>연락처</th>
                      <th>이메일</th>
                      <th>희망 일정</th>
                      <th>신청 시각</th>
                      <th>관리</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredApplications.length === 0 ? (
                      <tr>
                        <td colSpan={8} className="empty">
                          {applications.length === 0
                            ? '아직 접수된 참가자가 없습니다. 새 참가자가 등록하면 실시간으로 나타납니다.'
                            : '검색 조건과 일치하는 참가자가 없습니다.'}
                        </td>
                      </tr>
                    ) : (
                      filteredApplications.map((app, idx) => (
                        <tr key={app.id}>
                          <td data-label="No" className="masked nowrap">
                            {idx + 1}
                          </td>
                          <td data-label="접수 번호" className="masked nowrap font-bold">
                            {app.code}
                          </td>
                          <td data-label="이름" className="masked nowrap">
                            {maskName(app.name)}
                          </td>
                          <td data-label="연락처" className="masked nowrap">
                            {maskPhone(app.phone)}
                          </td>
                          <td data-label="이메일" className="masked">
                            {maskEmail(app.email)}
                          </td>
                          <td data-label="희망 일정">{app.slot}</td>
                          <td data-label="신청 시각" className="masked nowrap">
                            {fmtDate(app.createdAt)}
                          </td>
                          <td className="actcell">
                            <div className="acts">
                              <button
                                type="button"
                                className="ghost small"
                                onClick={() => {
                                  setViewAppId(app.id);
                                  setViewRevealed(false);
                                }}
                              >
                                보기
                              </button>
                              <button
                                type="button"
                                className="ghost small"
                                onClick={() => openAdminEdit(app)}
                              >
                                수정
                              </button>
                              <button
                                type="button"
                                className="danger small"
                                onClick={() => handleAdminDeleteOne(app)}
                              >
                                삭제
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              {applications.length > 0 && (
                <div className="row" style={{ marginTop: '24px' }}>
                  <button
                    type="button"
                    className="danger"
                    onClick={handleAdminClearAll}
                  >
                    명단 전체 삭제
                  </button>
                </div>
              )}
            </section>
          )}
        </div>
      )}

      {/* ───────── 기존 시트 연결 모달 ───────── */}
      {isLinkSheetModalOpen && (
        <div className="modal-overlay" role="dialog" aria-modal="true">
          <div className="modal-card">
            <h3>기존 Google 스프레드시트 연결</h3>
            <p style={{ color: 'var(--muted)', fontSize: '0.9rem', marginBottom: '14px' }}>
              연동할 Google 스프레드시트의 전체 URL 또는 스프레드시트 ID를 입력하세요.
            </p>
            <div className="field">
              <label htmlFor="sheet-id-input">스프레드시트 URL 또는 ID</label>
              <input
                id="sheet-id-input"
                type="text"
                placeholder="https://docs.google.com/spreadsheets/d/... 또는 ID"
                value={manualSheetInput}
                onChange={e => setManualSheetInput(e.target.value)}
              />
            </div>
            <div className="row">
              <button
                type="button"
                className="ghost"
                onClick={() => {
                  setIsLinkSheetModalOpen(false);
                  setManualSheetInput('');
                }}
              >
                취소
              </button>
              <button
                type="button"
                className="primary"
                disabled={isSyncingSheets || !manualSheetInput.trim()}
                onClick={handleLinkExistingSheet}
              >
                {isSyncingSheets ? '연결 확인 중...' : '연결하기'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ───────── 상세 보기 모달 (Organizer R) ───────── */}
      {viewTargetApp && (
        <div className="modal-overlay" role="dialog" aria-modal="true">
          <div className="modal-card wide">
            <h3>신청 상세 정보</h3>
            <dl>
              <dt>접수 번호</dt>
              <dd>
                <strong>{viewTargetApp.code}</strong>
              </dd>
              <dt>이름</dt>
              <dd>{viewRevealed ? viewTargetApp.name : maskName(viewTargetApp.name)}</dd>
              <dt>연락처</dt>
              <dd>{viewRevealed ? viewTargetApp.phone : maskPhone(viewTargetApp.phone)}</dd>
              <dt>이메일</dt>
              <dd>{viewRevealed ? viewTargetApp.email : maskEmail(viewTargetApp.email)}</dd>
              <dt>희망 일정</dt>
              <dd>{viewTargetApp.slot}</dd>
              <dt>남기실 말</dt>
              <dd>{viewTargetApp.memo || '—'}</dd>
              <dt>신청 시각</dt>
              <dd>{fmtDate(viewTargetApp.createdAt)}</dd>
              <dt>최종 수정</dt>
              <dd>{fmtDate(viewTargetApp.updatedAt)}</dd>
            </dl>

            <div className="row">
              <button
                type="button"
                className="ghost"
                onClick={() => setViewRevealed(!viewRevealed)}
              >
                {viewRevealed ? '다시 가리기' : '개인정보 보기'}
              </button>
              <button
                type="button"
                className="ghost"
                onClick={() => {
                  const target = viewTargetApp;
                  setViewAppId(null);
                  openAdminEdit(target);
                }}
              >
                수정
              </button>
            </div>
            <div className="row">
              <button
                type="button"
                className="primary"
                onClick={() => setViewAppId(null)}
              >
                닫기
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ───────── 주최자 추가·수정 모달 (C·U) ───────── */}
      {isAdminEditOpen && (
        <div className="modal-overlay" role="dialog" aria-modal="true">
          <div className="modal-card wide">
            <h3>{adminEditApp ? `신청 수정 · ${adminEditApp.code}` : '참가자 직접 추가'}</h3>

            <div className="field">
              <label htmlFor="edit-name">
                이름<span className="req">*</span>
              </label>
              <input
                id="edit-name"
                type="text"
                value={editFormName}
                onChange={e => handleNameChange(e.target.value, setEditFormName)}
                placeholder="한글 또는 영어, 10자 이내"
              />
            </div>

            <div className="field">
              <label htmlFor="edit-phone">
                연락처<span className="req">*</span>
              </label>
              <input
                id="edit-phone"
                type="tel"
                value={editFormPhone}
                onChange={e => handlePhoneChange(e.target.value, setEditFormPhone)}
                placeholder="010-0000-0000"
                maxLength={13}
              />
            </div>

            <div className="field">
              <label htmlFor="edit-email">
                이메일<span className="req">*</span>
              </label>
              <input
                id="edit-email"
                type="email"
                value={editFormEmail}
                onChange={e => setEditFormEmail(e.target.value.replace(/\s/g, ''))}
                maxLength={80}
              />
            </div>

            <div className="field">
              <label htmlFor="edit-slot">
                희망 일정<span className="req">*</span>
              </label>
              <select
                id="edit-slot"
                value={editFormSlot}
                onChange={e => setEditFormSlot(e.target.value)}
              >
                <option value="">선택하세요</option>
                {SLOTS.map(s => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            </div>

            <div className="field">
              <label htmlFor="edit-memo">남기실 말</label>
              <textarea
                id="edit-memo"
                value={editFormMemo}
                onChange={e => {
                  const c = [...e.target.value];
                  if (c.length > 100) {
                    setEditFormMemo(c.slice(0, 100).join(''));
                    setPopupMessage(MSG.memoMax);
                  } else {
                    setEditFormMemo(e.target.value);
                  }
                }}
              />
              <div className="count">{[...editFormMemo].length} / 100</div>
            </div>

            <div className="field">
              <label className="check">
                <input
                  type="checkbox"
                  checked={editFormAgree}
                  onChange={e => setEditFormAgree(e.target.checked)}
                />
                신청자에게 개인정보 수집 동의를 받았습니다.<span className="req">*</span>
              </label>
            </div>

            <div className="row">
              <button
                type="button"
                className="ghost"
                onClick={() => setIsAdminEditOpen(false)}
              >
                취소
              </button>
              <button
                type="button"
                className="primary"
                onClick={handleAdminSave}
              >
                저장
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ───────── 알림 팝업 ───────── */}
      {popupMessage && (
        <div className="modal-overlay" role="dialog" aria-modal="true">
          <div className="modal-card">
            <p style={{ margin: '0 0 20px', fontSize: '1.02rem' }}>{popupMessage}</p>
            <button
              type="button"
              className="primary"
              onClick={() => setPopupMessage(null)}
            >
              확인
            </button>
          </div>
        </div>
      )}

      {/* ───────── 예/아니요 확인 팝업 ───────── */}
      {confirmDialog && (
        <div className="modal-overlay" role="dialog" aria-modal="true">
          <div className="modal-card">
            <p style={{ margin: '0 0 20px', fontSize: '1.02rem' }}>{confirmDialog.message}</p>
            <div className="row" style={{ marginTop: 0 }}>
              <button
                type="button"
                className="ghost"
                onClick={() => setConfirmDialog(null)}
              >
                아니요
              </button>
              <button
                type="button"
                className="danger"
                onClick={() => {
                  const action = confirmDialog.onConfirm;
                  setConfirmDialog(null);
                  action();
                }}
              >
                예
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
