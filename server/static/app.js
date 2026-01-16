const dropzone = document.getElementById("dropzone");
const input = document.getElementById("fileInput");
const statusBox = document.getElementById("statusMessage");

dropzone.addEventListener("click", () => input.click());

dropzone.addEventListener("dragover", e => {
    e.preventDefault();
    dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", e => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    uploadFile(e.dataTransfer.files[0]);
});

input.addEventListener("change", () => {
    uploadFile(input.files[0]);
});


function uploadFile(file) {
    if (!file) return;

    statusBox.className = "status";
    statusBox.textContent = "⏳ Wysyłanie pliku...";
    statusBox.classList.remove("hidden");

    const form = new FormData();
    form.append("file", file);

    fetch("/upload", {
        method: "POST",
        body: form
    })
    .then(async r => {
        const data = await r.json();
        if (!r.ok) throw data;
        return data;
    })
    .then(() => {
        statusBox.textContent = "✅ Plik został nadpisany pomyślnie";
        statusBox.classList.add("success");
        refreshStatus();
    })
    .catch(err => {
        statusBox.textContent = "❌ " + (err.error || "Błąd wysyłania pliku");
        statusBox.classList.add("error");
    });
}


// ===== STATUS =====

function refreshStatus() {
    fetch("/status")
        .then(r => r.json())
        .then(data => {
            document.getElementById("currentStatus").textContent =
                data.current
                    ? `${data.current.number} (do ${data.current.end})`
                    : "Brak aktywnego przekierowania";

            document.getElementById("nextStatus").textContent =
                data.next
                    ? `${data.next.number} (od ${data.next.start})`
                    : "Brak kolejnych wpisów";
        })
        .catch(() => {
            document.getElementById("currentStatus").textContent = "Błąd odczytu";
            document.getElementById("nextStatus").textContent = "—";
        });
}

refreshStatus();
setInterval(refreshStatus, 30000);
