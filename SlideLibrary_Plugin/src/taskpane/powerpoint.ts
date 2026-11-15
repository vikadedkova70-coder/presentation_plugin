
/* global Office, document, FormData */

import { getPresentationBlob } from "../office/presentation";
import { uploadPresentation, ExportApiError } from "../api/export_api";
import { validateExportForm } from "../validation/validation";

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("export-form")
            ?.addEventListener("submit", handleExportSubmit);
    }
});

async function handleExportSubmit(event: Event) {
    event.preventDefault();

    const titleInput = document.getElementById("title") as HTMLInputElement;
    const descInput = document.getElementById("description") as HTMLTextAreaElement;
    const tagsInput = document.getElementById("tags") as HTMLInputElement;
    const exportBtn = document.getElementById("export-btn") as HTMLButtonElement;

    const title = titleInput.value.trim();
    const description = descInput.value.trim();
    const tags = tagsInput.value
        .split(",")
        .map(tag => tag.trim())
        .filter(Boolean);

    const errors = validateExportForm({
        title,
        description,
        tags,
    });

    const firstError = Object.values(errors)[0];

    if (firstError) {
        updateStatus(firstError, "error");
        return;
    }

    exportBtn.disabled = true;

    const resultLink = document.getElementById("result-link") as HTMLAnchorElement;
    resultLink.classList.add("hidden");

    updateStatus("Извлечение презентации из PowerPoint...", "info");

    try {
        const presentation = await getPresentationBlob(title);

        const formData = new FormData();

        formData.append("title", title);
        formData.append("description", description);
        formData.append("tags", tags.join(", "));
        formData.append("file", presentation.blob, presentation.fileName);

        updateStatus("Отправка презентации в Google Drive...", "info");

        const result = await uploadPresentation(formData);

        updateStatus("Презентация успешно сохранена в Google Drive.", "success");

        showResultLink(result.drive_url);
    } catch (error) {
        const message = error instanceof Error
            ? error.message
            : "Не удалось экспортировать презентацию.";

        updateStatus(`Ошибка: ${message}`, "error");

        if (error instanceof ExportApiError && error.driveUrl) {
            showResultLink(error.driveUrl);
        }
    } finally {
        exportBtn.disabled = false;
    }
}

function updateStatus(
    message: string,
    type: "info" | "success" | "error",
) {
    const statusBox = document.getElementById("status");
    const statusMsg = document.getElementById("status-message");

    if (!statusBox || !statusMsg) {
        return;
    }

    statusBox.className = `status-box ${type}`;
    statusMsg.textContent = message;
    statusBox.classList.remove("hidden");
}

function showResultLink(url: string) {
    const link = document.getElementById("result-link") as HTMLAnchorElement;

    if (link && url) {
        link.href = url;
        link.classList.remove("hidden");
    }
}
