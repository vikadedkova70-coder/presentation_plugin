import { getPresentationBlob, getSavedBaseName } from "../office/presentation";
import { uploadPresentation, buildExportFormData, ExportApiError } from "../api/export_api";
import {
    validateExportForm,
    validatePresentationFile,
    parseTags,
    normalizeFormData,
    hasErrors,
} from "../validation/validation";

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("export-form")
            ?.addEventListener("submit", handleExportSubmit);

        prefillTitle();
    }
});

function prefillTitle(): void {
    const titleInput = document.getElementById("title") as HTMLInputElement | null;
    const name = getSavedBaseName();
    if (titleInput && name && !titleInput.value) {
        titleInput.value = name;
    }
}

async function handleExportSubmit(event: Event): Promise<void> {
    event.preventDefault();

    const titleInput = document.getElementById("title") as HTMLInputElement;
    const descInput = document.getElementById("description") as HTMLTextAreaElement;
    const tagsInput = document.getElementById("tags") as HTMLInputElement;
    const exportBtn = document.getElementById("export-btn") as HTMLButtonElement;

    hideResultLink();

    const form = normalizeFormData({
        title: titleInput.value,
        description: descInput.value,
        tags: parseTags(tagsInput.value),
    });

    const errors = validateExportForm(form);
    if (hasErrors(errors)) {
        const firstError = Object.values(errors)[0] ?? "Проверьте заполнение формы.";
        updateStatus(firstError, "error");
        return;
    }

    exportBtn.disabled = true;

    try {
        updateStatus("Извлечение презентации из PowerPoint...", "info");
        const file = await getPresentationBlob(form.title);
        await validatePresentationFile(file.blob);

        updateStatus("Отправка презентации в Google Drive...", "info");
        const result = await uploadPresentation(buildExportFormData(file, form));

        updateStatus("Презентация успешно сохранена в Google Drive.", "success");
        showResultLink(result.drive_url, "Открыть в Google Drive");
    } catch (error) {
        console.error("Экспорт не удался:", error);

        const message = error instanceof Error
            ? error.message
            : "Не удалось экспортировать презентацию.";

        updateStatus(`Ошибка: ${message}`, "error");

        if (error instanceof ExportApiError && error.driveUrl) {
            showResultLink(error.driveUrl, "Открыть существующую презентацию");
        }
    } finally {
        exportBtn.disabled = false;
    }
}

function updateStatus(message: string, type: "info" | "success" | "error"): void {
    const statusBox = document.getElementById("status");
    const statusMsg = document.getElementById("status-message");

    if (!statusBox || !statusMsg) {
        return;
    }

    statusBox.className = `status-box ${type}`;
    statusMsg.textContent = message;
    statusBox.classList.remove("hidden");
}

function showResultLink(url: string, text: string): void {
    const link = document.getElementById("result-link") as HTMLAnchorElement | null;

    if (link && url) {
        link.href = url;
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        link.textContent = text;
        link.classList.remove("hidden");
    }
}

function hideResultLink(): void {
    document.getElementById("result-link")?.classList.add("hidden");
}
