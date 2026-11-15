
import { CONFIG } from "../types/config";
import type { ExportFormData, FormErrors } from "../types/export";

export function validateExportForm(data: ExportFormData): FormErrors {
    const errors: FormErrors = {};

    if (!data.title.trim()) {
        errors.title = "Введите название презентации.";
    } else if (data.title.trim().length < CONFIG.TITLE_MIN) {
        errors.title = "Название слишком короткое.";
    } else if (data.title.trim().length > CONFIG.TITLE_MAX) {
        errors.title = `Название не должно превышать ${CONFIG.TITLE_MAX} символов.`;
    }

    if (data.description.length > CONFIG.DESCRIPTION_MAX) {
        errors.description = `Описание не должно превышать ${CONFIG.DESCRIPTION_MAX} символов.`;
    }

    if (data.tags.length > CONFIG.TAGS_MAX) {
        errors.tags = `Допускается не более ${CONFIG.TAGS_MAX} тегов.`;
    } else if (data.tags.some(tag => tag.length > CONFIG.TAG_MAX_LENGTH)) {
        errors.tags = `Каждый тег не должен превышать ${CONFIG.TAG_MAX_LENGTH} символов.`;
    }

    return errors;
}
