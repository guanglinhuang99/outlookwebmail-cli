import type { OutlookState } from '../outlook/state.js';

export interface PageInfo {
  url?: string;
  title?: string;
  w?: number;
  h?: number;
  sx?: number;
  sy?: number;
  pw?: number;
  ph?: number;
  dialog?: unknown;
}

export interface BrowserStatus {
  connected: boolean;
  taskSpaceId: string | number | null;
  url: string | null;
  title: string | null;
  page: PageInfo;
  snapshot: string;
  browserName?: string | null;
  browserSession?: 'launched' | 'reused' | 'external-cdp' | 'shared-edge' | null;
}

export interface MailReadyProbe {
  ready: boolean;
  url: string | null;
  title: string | null;
  searchInputs: number;
  inboxFolders: number;
  mailLists: number;
  loginFrames: number;
  busy: boolean;
}

export interface LoginHandoffResult {
  taskSpaceId: string | number | null;
  url: string | null;
  handedOff: boolean;
}

export interface StatusResult {
  backend: 'ego-lite' | 'playwright';
  browser?: string | null;
  url: string;
  title: string | null;
  state: OutlookState;
  mailReady: boolean;
  browserSession?: 'launched' | 'reused' | 'external-cdp' | 'shared-edge' | null;
}

export interface DoctorCheck {
  name: 'node' | 'ego-lite' | 'playwright' | 'authentication' | 'dom';
  status: 'pass' | 'warn' | 'fail';
  message: string;
}

export interface DoctorResult {
  ok: boolean;
  checks: DoctorCheck[];
}

export interface InspectResult {
  backend: 'ego-lite' | 'playwright';
  capturedAt: string;
  state: OutlookState;
  page: PageInfo;
  snapshot: string;
  domInventory: unknown[];
  listCandidates: unknown[];
  messageRowCandidates: unknown[];
  scrollCandidates: unknown[];
  iframes: unknown[];
}

export interface MessageInspectResult {
  backend: 'ego-lite' | 'playwright';
  capturedAt: string;
  state: OutlookState;
  page: PageInfo;
  snapshot: string;
  bodyCandidates: unknown[];
  headerCandidates: unknown[];
  attachmentCandidates: unknown[];
  iframes: unknown[];
}
