export interface ExportFormData{
    title: string;
    description: string;
    tags: string[];
}

export interface ExportResult {
    id: string;
    success: boolean;
    link: string[];
}

export type FormErrors = Partial<Record<keyof ExportFormData, string>>;
