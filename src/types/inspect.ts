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
}

export interface StatusResult {
  backend: 'ego-lite';
  url: string;
  title: string | null;
  state: OutlookState;
}

export interface InspectResult {
  backend: 'ego-lite';
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
  backend: 'ego-lite';
  capturedAt: string;
  state: OutlookState;
  page: PageInfo;
  snapshot: string;
  bodyCandidates: unknown[];
  headerCandidates: unknown[];
  attachmentCandidates: unknown[];
  iframes: unknown[];
}
