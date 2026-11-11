/* global Office */

export interface PresentationFileData {
    blob: Blob;
    fileName: string;
}

/**
 * Получает текущую открытую презентацию PowerPoint в виде Blob (.pptx)
 */
export function getPresentationBlob(): Promise<PresentationFileData> {
    return new Promise((resolve, reject) => {
        // Запрос на сжатый файл (.pptx) кусками по 1 МБ (1048576 байт)
        Office.context.document.getFileAsync(
            Office.FileType.Compressed,
            { sliceSize: 1048576 },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    reject(new Error(`Не удалось получить файл: ${result.error.message}`));
                    return;
                }

                const file = result.value;
                const sliceCount = file.sliceCount;
                const slices: Uint8Array[] = [];
                let slicesReceived = 0;

                // Вспомогательная функция для рекурсивного чтения слайсов по порядку
                const getNextSlice = (sliceIndex: number) => {
                    file.getSliceAsync(sliceIndex, (sliceResult) => {
                        if (sliceResult.status === Office.AsyncResultStatus.Failed) {
                            file.closeAsync();
                            reject(new Error(`Ошибка чтения части файла: ${sliceResult.error.message}`));
                            return;
                        }

                        // Превращает данные слайса в Uint8Array
                        const sliceData = new Uint8Array(sliceResult.value.data);
                        slices[sliceIndex] = sliceData;
                        slicesReceived++;

                        if (slicesReceived === sliceCount) {
                            // Когда все куски получены, закрываем дескриптор файла и создаём блоб
                            file.closeAsync();

                            const blobParts: BlobPart[] = slices.map((slice) => {
                                const copy = new Uint8Array(slice.byteLength);
                                copy.set(slice);
                                return copy.buffer;
                            });

                            const blob = new Blob(blobParts, {
                                type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            });

                            const fileName = extractFileName();
                            resolve({ blob, fileName });
                        } else {
                            getNextSlice(sliceIndex + 1);
                        }
                    });
                };

                // Запуск чтения с нулевого слайса
                getNextSlice(0);
            }
        );
    });
}

/**
 * Достает имя текущего файла из URL документа (если презентация сохранена)
 */
function extractFileName(): string {
    const docUrl = Office.context.document.url;
    if (!docUrl) {
        return "presentation.pptx";
    }

    const cleanUrl = docUrl.split("?")[0];
    const parts = cleanUrl.split(/[/\\]/);
    const lastPart = parts[parts.length - 1];

    if (lastPart && lastPart.toLowerCase().endsWith(".pptx")) {
        return decodeURIComponent(lastPart);
    }

    return "presentation.pptx";
}
