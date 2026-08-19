export type RunStatus = "queued" | "running" | "partial" | "completed" | "failed" | "cancelled";
export interface Pharmacy { id: string; name: string; url: string }
export interface Profile { id: string; name: string; pharmacies: Pharmacy[]; reference_pharmacy_id: string | null; archived: boolean }
export interface Settings {
  schema_version: number; onboarding_complete: boolean; legacy_migrated: boolean; theme: "light" | "dark";
  output_directory: string; file_name_template: string; report_format: "xlsm" | "xlsx";
  green: string; red: string; retention: number | null; check_updates: boolean;
  window_width: number; window_height: number; profiles: Profile[]; selected_profile_id: string | null;
}
export interface CredentialStatus { configured: boolean; backend: string; masked_cookie?: string; warning?: string }
export interface Run {
  id: string; profile_id: string; parent_run_id: string | null; status: RunStatus;
  started_at: string; finished_at: string | null; reference_pharmacy_id: string;
  pharmacy_count: number; successful_pharmacies: number; product_count: number;
  pinned: boolean; report_path: string | null; warning_count: number;
}
export interface Bootstrap { version: string; settings: Settings; credentials: CredentialStatus; active_run_id: string | null; history_size_bytes: number; legacy_config_present: boolean }
export interface ProgressEvent { sequence: number; run_id: string; kind: string; pharmacy_id?: string; stage: string; message: string; current?: number; total?: number; timestamp: string }
