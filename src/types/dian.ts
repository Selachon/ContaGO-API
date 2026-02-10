export interface DownloadRequest {
  token_url: string;
  start_date?: string;
  end_date?: string;
  session_uid?: string;
}

export interface DocumentInfo {
  id: string;
  docnum: string;
  nit: string;
}

export interface ProgressData {
  step: string;
  current: number;
  total: number;
  detalle?: string;
}
