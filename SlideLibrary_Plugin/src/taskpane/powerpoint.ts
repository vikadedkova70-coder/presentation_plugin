/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, document, fetch, FormData, Blob */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        const form = document.getElementById("export-form");
        if (form) {
            form.addEventListener("submit", handleExportSubmit);
        }
    }
});

async function handleExportSubmit(event: Event) {
    event.preventDefault();

    const titleInput = document.getElementById("title") as HTMLInputElement;
    const descInput = document.getElementById("description") as HTMLTextAreaElement;
    const tagsInput = document.getElementById("tags") as HTMLInputElement;
    const exportBtn = document.getElementById("export-btn") as HTMLButtonElement;

    if (exportBtn) exportBtn.disabled = true;
    updateStatus("Извлечение файла из PowerPoint...", "info");

    try {
        const presentationSlice = await getPresentationFileBlob();

        const formData = new FormData();
        formData.append("file", presentationSlice, `${titleInput?.value || "presentation"}.pptx`);
        formData.append("title", titleInput?.value || "");
        formData.append("description", descInput?.value || "");
        formData.append("tags", tagsInput?.value || "");

        updateStatus("Отправка на сервер...", "info");

        const BACKEND_URL = "https://localhost:8000/api/exports"; 
        const response = await fetch(BACKEND_URL, {
            method: "POST",
            body: formData
        });

        if (!response.ok) {
            throw new Error(`Ошибка сервера: ${response.statusText}`);
        }

        const data = await response.json();
        updateStatus("Файл успешно экспортирован!", "success");
        showResultLink(data.file_url || data.url || "#");

    } catch (err: any) {
        updateStatus(`Ошибка: ${err.message || "Не удалось экспортировать"}`, "error");
    } finally {
        if (exportBtn) exportBtn.disabled = false;
    }
}

function getPresentationFileBlob(): Promise<Blob> {
    return new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(
            Office.FileType.Compressed,
            { sliceSize: 65536 },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    reject(new Error(result.error.message));
                    return;
                }

                const file = result.value;
                const slicesCount = file.sliceCount;
                const slicePromises: Promise<ArrayBuffer>[] = [];

                for (let i = 0; i < slicesCount; i++) {
                    slicePromises.push(
                        new Promise((res, rej) => {
                            file.getSliceAsync(i, (sliceResult) => {
                                if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                    res(sliceResult.value.data);
                                } else {
                                    rej(sliceResult.error.message);
                                }
                            });
                        })
                    );
                }

                Promise.all(slicePromises)
                    .then((buffers) => {
                        file.closeAsync();
                        const blob = new Blob(buffers, {
                            type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        });
                        resolve(blob);
                    })
                    .catch((err) => {
                        file.closeAsync();
                        reject(err);
                    });
            }
        );
    });
}

function updateStatus(message: string, type: "info" | "success" | "error") {
    const statusBox = document.getElementById("status");
    const statusMsg = document.getElementById("status-message");

    if (!statusBox || !statusMsg) return;

    statusBox.className = `status-box ${type}`;
    statusMsg.textContent = message;
    statusBox.classList.remove("hidden");
}

function showResultLink(url: string) {
    const link = document.getElementById("result-link") as HTMLAnchorElement;
    if (link && url && url !== "#") {
        link.href = url;
        link.classList.remove("hidden");
    }
}