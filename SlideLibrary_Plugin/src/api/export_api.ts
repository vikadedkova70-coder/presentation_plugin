import { CONFIG } from "../types/config";
import type { ApiErrorResponse, ExportFormData, ExportResult } from "../types/export";
import type { PresentationFileData } from "../office/presentation";

export class ExportApiError extends Error {
    constructor(
        message: string,
        public readonly status: number,
        public readonly driveUrl?: string,
        public readonly cancelled = false,
    ) {
        super(message);
        this.name = "ExportApiError";
        Object.setPrototypeOf(this, ExportApiError.prototype);
    }
}

const STATUS_MESSAGES: Record<number, string> = {
    400: "Файл пустой или повреждён.",
    409: "Такая презентация уже загружена.",
    413: "Превышен максимальный размер файла.",
    415: "Поддерживаются только файлы PPTX.",
    422: "Проверьте заполнение обязательных полей.",
    502: "Ошибка сохранения в Google Drive.",
    503: "Сервер временно недоступен. Попробуйте позже.",
    504: "Сервер не ответил вовремя. Попробуйте позже.",
};

function isHttpUrl(value: unknown): value is string {
    if (typeof value !== "string") return false;
    try {
        const { protocol } = new URL(value);
        return protocol === "http:" || protocol === "https:";
    } catch {
        return false;
    }
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
    typeof value === "object" && value !== null && !Array.isArray(value);

const isAbortError = (error: unknown): boolean =>
    (error as { name?: string } | null)?.name === "AbortError";

export function buildExportFormData(file: PresentationFileData, form: ExportFormData): FormData {
    const formData = new FormData();
    formData.append("file", file.blob, file.fileName);
    formData.append("title", form.title);
    formData.append("description", form.description);
    formData.append("tags", form.tags.join(", "));
    return formData;
}

function buildHttpError(status: number, detail: ApiErrorResponse["detail"]): ExportApiError {
    if (status === 409 && isRecord(detail)) {
        return new ExportApiError(
            typeof detail.message === "string" ? detail.message : STATUS_MESSAGES[409],
            status,
            isHttpUrl(detail.drive_url) ? detail.drive_url : undefined,
        );
    }

    const message =
        STATUS_MESSAGES[status] ??
        (typeof detail === "string" ? detail : `Ошибка сервера: ${status}`);

    return new ExportApiError(message, status);
}

export async function uploadPresentation(
    formData: FormData,
    signal?: AbortSignal,
): Promise<ExportResult> {
    const controller = new AbortController();
    let timedOut = false;

    const timeout = window.setTimeout(() => {
        timedOut = true;
        controller.abort();
    }, CONFIG.REQUEST_TIMEOUT_MS);

    const onExternalAbort = () => controller.abort();
    if (signal?.aborted) controller.abort();
    signal?.addEventListener("abort", onExternalAbort);

    try {
        const response = await fetch(`${CONFIG.API_BASE_URL}/api/exports`, {
            method: "POST",
            body: formData,
            signal: controller.signal,
        });

        const body = (await response.json().catch(() => ({}))) as ExportResult & ApiErrorResponse;

        if (!response.ok) {
            throw buildHttpError(response.status, body.detail);
        }

        if (!isHttpUrl(body.drive_url) || !body.drive_file_id) {
            throw new ExportApiError(
                "Сервер не вернул информацию о загруженной презентации.",
                response.status,
            );
        }

        return body;
    } catch (error) {
        if (error instanceof ExportApiError) throw error;

        if (isAbortError(error)) {
            if (timedOut) {
                throw new ExportApiError("Превышено время ожидания ответа сервера.", 0);
            }
            throw new ExportApiError("Экспорт отменён.", 0, undefined, true);
        }

        throw new ExportApiError("Не удалось подключиться к серверу. Проверьте подключение.", 0);
    } finally {
        window.clearTimeout(timeout);
        signal?.removeEventListener("abort", onExternalAbort);
    }
}
