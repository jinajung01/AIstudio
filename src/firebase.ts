import { initializeApp, getApps, getApp } from 'firebase/app';
import { getAuth } from 'firebase/auth';
import { getFirestore, doc, getDocFromServer } from 'firebase/firestore';
import oauthFirebaseConfig from '../firebase-applet-config.json';

// Provisioned Cloud Firestore project and database
export const firestoreProjectConfig = {
  projectId: 'hypnagogic-ascent-spz3c',
  apiKey: 'AIzaSyC6Ili8V8GbVBQZfFIdpbvBuTCpHBdRo58',
  authDomain: 'hypnagogic-ascent-spz3c.firebaseapp.com',
  appId: '1:5900119638:web:71244ffec9a7af90fbd5f7',
  firestoreDatabaseId: 'ai-studio-0e82a620-9a1f-4748-996e-f2582b9c7250'
};

// Initialize dedicated Firebase app for Firestore
const firestoreApp = getApps().find(a => a.name === 'firestore-app')
  || initializeApp(firestoreProjectConfig, 'firestore-app');

export const db = getFirestore(firestoreApp, firestoreProjectConfig.firestoreDatabaseId);

// Initialize default Firebase app for Google OAuth / Sheets Auth
const authApp = getApps().find(a => a.name === '[DEFAULT]')
  || initializeApp(oauthFirebaseConfig);

export const auth = getAuth(authApp);

export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

export interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  };
}

export function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null): never {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

export async function testConnection(): Promise<boolean> {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
    return true;
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      console.warn('Firestore is offline or unreachable:', error.message);
      return false;
    }
    return true;
  }
}
