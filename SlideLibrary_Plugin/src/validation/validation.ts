import { CONFIG } from "../types/config";
import type { ExportFormData, FormErrors } from "../types/export";
import { ExportError } from "./errors";

const toMB = (bytes: number) => (bytes / 1024 / 1024).toFixed(1);

export async function validatePresentationFile(blob: Blob): Promise<void> {
    if (blob.size === 0) {
        throw new ExportError("FILE_EMPTY", "Файл пустой. Сохраните презентацию и повторите экспорт.");
    }
    if (blob.size < CONFIG.MIN_FILE_SIZE) {
        throw new ExportError("FILE_INVALID", "Файл повреждён или слишком мал для презентации.");
    }
    if (blob.size > CONFIG.MAX_FILE_SIZE) {
        throw new ExportError(
            "FILE_TOO_LARGE",
            `Файл слишком большой (${toMB(blob.size)} МБ). Максимум — ${toMB(CONFIG.MAX_FILE_SIZE)} МБ.`,
        );
    }

    const head = new Uint8Array(await blob.slice(0, 4).arrayBuffer());
    const isZip = head[0] === 0x50 && head[1] === 0x4b && head[2] === 0x03 && head[3] === 0x04;
    if (!isZip) {
        throw new ExportError("FILE_INVALID", "Файл не похож на презентацию PowerPoint (.pptx).");
    }
}

export function parseTags(raw: string): string[] {
    const seen = new Set<string>();
    return raw
        .split(/[,;\n]+/)
        .map((tag) => tag.trim().replace(/^#+/, "").replace(/\s+/g, " "))
        .filter((tag) => {
            if (!tag) return false;
            const key = tag.toLowerCase();
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
}

export function normalizeFormData(data: ExportFormData): ExportFormData {
    return {
        title: data.title.trim().replace(/\s+/g, " "),
        description: data.description.trim(),
        tags: data.tags,
    };
}

export function validateExportForm(data: ExportFormData): FormErrors {
    const errors: FormErrors = {};
    const title = data.title.trim();

    if (!title) {
        errors.title = "Введите название презентации.";
    } else if (title.length < CONFIG.TITLE_MIN) {
        errors.title = `Название слишком короткое (минимум ${CONFIG.TITLE_MIN} символа).`;
    } else if (title.length > CONFIG.TITLE_MAX) {
        errors.title = `Название не должно превышать ${CONFIG.TITLE_MAX} символов.`;
    }

    if (data.description.length > CONFIG.DESCRIPTION_MAX) {
        errors.description = `Описание не должно превышать ${CONFIG.DESCRIPTION_MAX} символов.`;
    }

    if (data.tags.length > CONFIG.TAGS_MAX) {
        errors.tags = `Допускается не более ${CONFIG.TAGS_MAX} тегов.`;
    } else if (data.tags.some((tag) => tag.length > CONFIG.TAG_MAX_LENGTH)) {
        errors.tags = `Каждый тег не должен превышать ${CONFIG.TAG_MAX_LENGTH} символов.`;
    }

    return errors;
}

export const hasErrors = (errors: FormErrors): boolean => Object.keys(errors).length > 0;
