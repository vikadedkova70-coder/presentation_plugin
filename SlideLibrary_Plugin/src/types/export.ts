
export interface ExportFormData {
    title: string;
    description: string;
    tags: string[];
}

export interface ExportResult {
    id: number;
    title: string;
    status: "success";
    sha256: string;
    drive_file_id: string;
    drive_url: string;
}

export interface ApiErrorDetail {
    message?: string;
    drive_url?: string;
}

export interface ApiErrorResponse {
    detail?: string | ApiErrorDetail;
}

export type FormErrors = Partial<Record<keyof ExportFormData, string>>;
