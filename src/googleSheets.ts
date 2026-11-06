import type { Application } from './types.ts';

export const SHEETS_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.file'
];

let inMemoryAccessToken: string | null = null;

export function setSheetsToken(token: string | null) {
  inMemoryAccessToken = token;
}

export function getSheetsToken(): string | null {
  return inMemoryAccessToken;
}

export interface SheetsConfig {
  spreadsheetId: string;
  spreadsheetUrl: string;
  sheetName: string;
  autoSync: boolean;
  updatedAt: string;
}

function parseGoogleApiError(errData: any, status: number): string {
  const msg = errData?.error?.message || '';
  if (status === 403) {
    if (msg.includes('Google Sheets API has not been used') || msg.includes('disabled')) {
      return 'Google Cloud 콘솔에서 "Google Sheets API"가 활성화되지 않았습니다. 아래 링크에서 [사용] 버튼을 누른 후 다시 시도해 주세요:\nhttps://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=671224416058';
    }
    if (msg.includes('insufficient') || msg.includes('scope')) {
      return 'Google 스프레드시트 접근 권한이 체크되지 않았습니다. 로그아웃 후 다시 로그인하여 권한을 허용해 주세요.';
    }
    return `권한 오류 (403): ${msg}`;
  }
  if (status === 401) {
    return 'Google 인증 세션이 만료되었습니다. 다시 로그인해 주세요.';
  }
  if (status === 404) {
    return '스프레드시트를 찾을 수 없습니다. 스프레드시트 ID 또는 공유 URL을 확인해 주세요.';
  }
  return msg || `Google API 오류 (${status})`;
}

/**
 * Creates a dedicated Google Spreadsheet for event applications
 * and initializes column headers.
 */
export async function createEventSpreadsheet(
  accessToken: string,
  customTitle?: string
): Promise<{ spreadsheetId: string; spreadsheetUrl: string; sheetName: string }> {
  const title = customTitle || `이벤트 참가 신청 명단 (실시간)`;
  const sheetName = '참가자명단';

  const createRes = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      properties: { title },
      sheets: [
        {
          properties: {
            title: sheetName,
            gridProperties: {
              frozenRowCount: 1
            }
          }
        }
      ]
    })
  });

  if (!createRes.ok) {
    const err = await createRes.json().catch(() => ({}));
    throw new Error(parseGoogleApiError(err, createRes.status));
  }

  const data = await createRes.json();
  const spreadsheetId = data.spreadsheetId;
  const spreadsheetUrl = data.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

  // Initialize header row
  const headers = [
    'No',
    '접수 번호',
    '이름',
    '연락처',
    '이메일',
    '희망 일정',
    '남기실 말',
    '신청 시각',
    '동기화 시각'
  ];

  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!A1:I1?valueInputOption=USER_ENTERED`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        values: [headers]
      })
    }
  );

  return { spreadsheetId, spreadsheetUrl, sheetName };
}

/**
 * Fetch spreadsheet metadata to get valid sheet tab name
 */
export async function getSpreadsheetDetails(
  accessToken: string,
  spreadsheetId: string
): Promise<{ title: string; sheetNames: string[] }> {
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(parseGoogleApiError(err, res.status));
  }

  const data = await res.json();
  const title = data.properties?.title || '스프레드시트';
  const sheetNames = (data.sheets || []).map((s: any) => s.properties?.title || 'Sheet1');
  return { title, sheetNames };
}

/**
 * Reads existing ticket codes in Column B to prevent duplicate entries
 */
export async function getExistingTicketCodes(
  accessToken: string,
  spreadsheetId: string,
  sheetName: string
): Promise<Set<string>> {
  try {
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!B2:B`,
      {
        headers: { Authorization: `Bearer ${accessToken}` }
      }
    );

    if (!res.ok) return new Set();
    const data = await res.json();
    const rows = data.values || [];
    const codeSet = new Set<string>();
    rows.forEach((r: any[]) => {
      if (r && r[0]) codeSet.add(String(r[0]).trim());
    });
    return codeSet;
  } catch {
    return new Set();
  }
}

/**
 * Appends rows to a spreadsheet
 */
export async function appendApplicantRows(
  accessToken: string,
  spreadsheetId: string,
  sheetName: string,
  rows: (string | number)[][]
): Promise<number> {
  if (rows.length === 0) return 0;

  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!A:I:append?valueInputOption=USER_ENTERED`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        values: rows
      })
    }
  );

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(parseGoogleApiError(err, res.status));
  }

  return rows.length;
}

/**
 * Formats an application into a spreadsheet row
 */
export function formatApplicationRow(app: Application, index: number): (string | number)[] {
  const fmtDate = (iso: string) => {
    try {
      return new Date(iso).toLocaleString('ko-KR', { dateStyle: 'short', timeStyle: 'short' });
    } catch {
      return iso;
    }
  };

  return [
    index,
    app.code,
    app.name,
    app.phone,
    app.email,
    app.slot,
    app.memo || '',
    fmtDate(app.createdAt),
    fmtDate(new Date().toISOString())
  ];
}
