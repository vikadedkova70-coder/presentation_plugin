
import { CONFIG } from "../types/config";
import type { ApiErrorResponse, ExportResult } from "../types/export";

export class ExportApiError extends Error {
    constructor(
        message: string,
        public readonly status: number,
        public readonly driveUrl?: string,
    ) {
        super(message);
        this.name = "ExportApiError";
    }
}

export async function uploadPresentation(formData: FormData): Promise<ExportResult> {
    const controller = new AbortController();

    const timeout = window.setTimeout(
        () => controller.abort(),
        CONFIG.REQUEST_TIMEOUT_MS,
    );

    try {
        const response = await fetch(`${CONFIG.API_BASE_URL}/api/exports`, {
            method: "POST",
            body: formData,
            signal: controller.signal,
        });

        const body = await response.json().catch(() => ({})) as ExportResult & ApiErrorResponse;

        if (!response.ok) {
            const detail = body.detail;

            if (typeof detail === "object" && detail !== null) {
                throw new ExportApiError(
                    detail.message || "Презентация уже существует.",
                    response.status,
                    detail.drive_url,
                );
            }

            const messages: Record<number, string> = {
                400: "Файл пустой или повреждён.",
                409: "Такая презентация уже загружена.",
                413: "Превышен максимальный размер файла.",
                415: "Поддерживаются только файлы PPTX.",
                422: "Проверьте заполнение обязательных полей.",
                502: "Ошибка сохранения в Google Drive.",
            };

            throw new ExportApiError(
                typeof detail === "string"
                    ? detail
                    : messages[response.status] || `Ошибка сервера: ${response.status}`,
                response.status,
            );
        }

        if (!body.drive_url || !body.drive_file_id) {
            throw new ExportApiError(
                "Сервер не вернул информацию о загруженной презентации.",
                response.status,
            );
        }

        return body;
    } catch (error) {
        if (error instanceof ExportApiError) {
            throw error;
        }

        if (error instanceof DOMException && error.name === "AbortError") {
            throw new ExportApiError(
                "Превышено время ожидания ответа сервера.",
                0,
            );
        }

        throw new ExportApiError(
            "Не удалось подключиться к серверу. Проверьте подключение.",
            0,
        );
    } finally {
        window.clearTimeout(timeout);
    }
}
