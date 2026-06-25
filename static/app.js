const input = document.getElementById("sales-file");
const zone = document.getElementById("drop-zone");
const label = document.getElementById("file-label");
const form = document.getElementById("upload-form");
const button = document.getElementById("analyse-button");

function showFile(file) {
    if (!file) return;
    label.textContent = file.name;
    zone.classList.add("has-file");
}

input?.addEventListener("change", () => showFile(input.files[0]));

["dragenter", "dragover"].forEach((eventName) => {
    zone?.addEventListener(eventName, (event) => {
        event.preventDefault();
        zone.classList.add("is-dragging");
    });
});

["dragleave", "drop"].forEach((eventName) => {
    zone?.addEventListener(eventName, (event) => {
        event.preventDefault();
        zone.classList.remove("is-dragging");
    });
});

zone?.addEventListener("drop", (event) => {
    if (!event.dataTransfer.files.length) return;
    const transfer = new DataTransfer();
    transfer.items.add(event.dataTransfer.files[0]);
    input.files = transfer.files;
    showFile(input.files[0]);
});

form?.addEventListener("submit", () => {
    button.disabled = true;
    button.querySelector("span:first-child").textContent = "Анализируем…";
});
