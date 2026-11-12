/* Получение .pptx и имя файла */

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

/* Получение текущего .pptx и имя файла */
export async function getPresentationBlob(): Promise<PresentationFileData> {
    const blob = await readCompressedFile();
    const baseName = "TestName";
    return { blob, fileName: `${baseName}.pptx` };
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
                        `Файл слишком большой (${toMB(file.size)} МБ). Максимум — ${toMB(CONFIG.MAX_FILE_SIZE)} МБ.`,
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
