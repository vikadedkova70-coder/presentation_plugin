import { CONFIG } from "../types/config";
import { ExportError } from "../validation/errors";

export interface PresentationFileData {
    blob: Blob;
    fileName: string;
}

const PPTX_MIME =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const SLICE_SIZE = 1024 * 1024; // 1 МБ
const toMB = (bytes: number) => (bytes / 1024 / 1024).toFixed(1) + " МБ";

export async function getPresentationBlob(title: string): Promise<PresentationFileData> {
    const blob = await readCompressedFile();
    return { blob, fileName: buildFileName(title) };
}

function readCompressedFile(): Promise<Blob> {
    return new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(
            Office.FileType.Compressed,
            { sliceSize: SLICE_SIZE },
            (result) => {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new ExportError(
                        "FILE_READ_FAILED",
                        "Не удалось получить презентацию из PowerPoint. Попробуйте ещё раз.",
                        result.error,
                    ));
                    return;
                }

                const file = result.value;

                if (file.size === 0 || file.sliceCount === 0) {
                    file.closeAsync();
                    reject(new ExportError("FILE_EMPTY", "Презентация пустая. Добавьте слайды и повторите экспорт."));
                    return;
                }
                if (file.size > CONFIG.MAX_FILE_SIZE) {
                    file.closeAsync();
                    reject(new ExportError(
                        "FILE_TOO_LARGE",
                        `Файл слишком большой (${toMB(file.size)}). Максимум — ${toMB(CONFIG.MAX_FILE_SIZE)}.`,
                    ));
                    return;
                }

                const parts: Uint8Array<ArrayBuffer>[] = new Array(file.sliceCount);

                const readSlice = (index: number) => {
                    file.getSliceAsync(index, (slice) => {
                        if (slice.status !== Office.AsyncResultStatus.Succeeded) {
                            file.closeAsync();
                            reject(new ExportError(
                                "FILE_READ_FAILED",
                                "Ошибка при чтении презентации. Попробуйте ещё раз.",
                                slice.error,
                            ));
                            return;
                        }

                        parts[index] = new Uint8Array(slice.value.data) as Uint8Array<ArrayBuffer>;

                        if (index + 1 < file.sliceCount) {
                            readSlice(index + 1);
                        } else {
                            file.closeAsync();
                            resolve(new Blob(parts, { type: PPTX_MIME }));
                        }
                    });
                };

                readSlice(0);
            },
        );
    });
}

const FORBIDDEN_CHARS = /[<>:"/\\|?*\u0000-\u001F]/g;

export function sanitizeFileName(raw: string): string {
    return raw
        .replace(FORBIDDEN_CHARS, "_")
        .replace(/\s+/g, " ")
        .replace(/^[.\s]+|[.\s]+$/g, "")
        .slice(0, 120);
}

export function getSavedBaseName(): string | null {
    const url = Office.context.document.url;
    if (!url) return null;

    const last = url.split(/[?#]/)[0].split(/[/\\]/).pop();
    if (!last) return null;

    let decoded = last;
    try {
        decoded = decodeURIComponent(last);
    } catch {
    }

    const base = sanitizeFileName(decoded.replace(/\.(pptx|pptm|ppsx|potx|pps|ppt)$/i, ""));
    return base || null;
}

function timestampName(): string {
    const d = new Date();
    const p = (n: number) => String(n).padStart(2, "0");
    return `Презентация_${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}-${p(d.getMinutes())}`;
}

export function buildFileName(title: string): string {
    const base = sanitizeFileName(title) || getSavedBaseName() || timestampName();
    return `${base}.pptx`;
}
