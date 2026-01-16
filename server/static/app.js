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

/* ===== FILE UPLOAD ===== */

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
            statusBox.textContent =
                "❌ " + (err.error || "Błąd wysyłania pliku");
            statusBox.classList.add("error");
        });
}

/* ===== STATUS ===== */

function refreshStatus() {
    fetch("/status")
        .then(r => r.json())
        .then(data => {
            const current = document.getElementById("currentStatus");
            const next = document.getElementById("nextStatus");
            const dot = document.getElementById("activeDot");

            if (data.current) {
                current.textContent =
                    `${data.current.person} • do ${data.current.end}`;
                dot.classList.remove("hidden");
            } else {
                current.textContent = "Brak aktywnego przekierowania";
                dot.classList.add("hidden");
            }

            if (data.next) {
                next.textContent =
                    `${data.next.person} • od ${data.next.start}`;
            } else {
                next.textContent = "Brak kolejnych wpisów";
            }
        })
        .catch(() => {
            document.getElementById("currentStatus").textContent = "Błąd odczytu";
            document.getElementById("nextStatus").textContent = "—";
            document.getElementById("activeDot").classList.add("hidden");
        });
}

refreshStatus();
setInterval(refreshStatus, 30000);
