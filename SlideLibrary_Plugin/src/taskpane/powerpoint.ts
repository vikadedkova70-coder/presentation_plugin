/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, document, fetch, FormData */

import { getPresentationBlob } from "../office/presentation";

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
        const { blob: presentationBlob } = await getPresentationBlob();

        const formData = new FormData();
        formData.append("file", presentationBlob, `${titleInput?.value || "presentation"}.pptx`);
        formData.append("title", titleInput?.value || "");
        formData.append("description", descInput?.value || "");
        formData.append("tags", tagsInput?.value || "");

        updateStatus("Отправка на сервер...", "info");
        // TODO: Вынести в export_api.ts при интеграции с бэкендом
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