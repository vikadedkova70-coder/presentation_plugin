// обработка ошибок, вывод сообщений пользователю

export type ExportErrorCode =
    | "FILE_READ_FAILED" | "FILE_EMPTY" | "FILE_INVALID" | "FILE_TOO_LARGE"
    | "VALIDATION" | "DUPLICATE" | "NETWORK" | "TIMEOUT"
    | "UPSTREAM_UNAVAILABLE" | "SERVER_ERROR" | "CANCELLED" | "UNKNOWN";

export class ExportError extends Error {
    constructor(
        public readonly code: ExportErrorCode,
        message: string,
        public readonly details?: unknown,
    ) {
        super(message);
        this.name = "ExportError";
        Object.setPrototypeOf(this, ExportError.prototype);
    }
}

export function toUserMessage(error: unknown): string {
    return error instanceof ExportError
        ? error.message
        : "Неизвестная ошибка. Попробуйте ещё раз или обратитесь в поддержку.";
}
