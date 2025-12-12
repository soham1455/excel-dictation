// app.js
window.addEventListener("load", () => {

    const startBtn = document.getElementById("startBtn");
    const stopBtn  = document.getElementById("stopBtn");
    const statusEl = document.getElementById("status");
    const historyEl = document.getElementById("history");
    const langSel = document.getElementById("langSel");
    const copyBtn = document.getElementById("copyBtn");

    function updateStatus(msg) {
        statusEl.textContent = "Status: " + msg;
    }

    function addTranscript(text) {
        const div = document.createElement("div");
        div.textContent = text;
        historyEl.prepend(div);
    }

    // Mic + SpeechRecognition support
    const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;

    if (!SpeechRec) {
        updateStatus("SpeechRecognition is not supported in this browser.");
        startBtn.disabled = true;
        return;
    }

    let recognition = new SpeechRec();
    recognition.interimResults = true;
    recognition.continuous = true;

    // EVENTS
    recognition.onstart = () => updateStatus("listening...");
    recognition.onend   = () => {
        updateStatus("stopped");
        startBtn.disabled = false;
        stopBtn.disabled = true;
    };

    recognition.onerror = (e) => {
        console.error("Recognition error:", e);
        updateStatus("error: " + e.error);
    };

    recognition.onresult = (event) => {
        let finalText = "";
        for (let i = event.resultIndex; i < event.results.length; i++) {
            const result = event.results[i];
            if (result.isFinal) {
                finalText += result[0].transcript;
            }
        }
        if (finalText) {
            addTranscript(finalText);
        }
    };

    // START BUTTON
    startBtn.addEventListener("click", async () => {
        try {
            // Explicit mic permission
            await navigator.mediaDevices.getUserMedia({ audio: true });

            recognition.lang = langSel.value;
            recognition.start();
            startBtn.disabled = true;
            stopBtn.disabled = false;
            updateStatus("starting...");
        } 
        catch (err) {
            console.error("Mic permission error:", err);
            alert("Microphone access blocked. Please allow mic permission in Chrome.");
            updateStatus("mic blocked");
        }
    });

    // STOP BUTTON
    stopBtn.addEventListener("click", () => {
        recognition.stop();
    });

    // COPY BUTTON
    copyBtn.addEventListener("click", async () => {
        const first = historyEl.firstElementChild;
        if (!first) {
            alert("No transcript yet!");
            return;
        }

        const text = first.textContent;

        try {
            await navigator.clipboard.writeText(text);
            copyBtn.textContent = "Copied ✓";
            setTimeout(() => copyBtn.textContent = "Copy Last Transcript", 1200);
        } 
        catch (err) {
            alert("Copy failed. Please select manually.");
        }
    });

    updateStatus("idle");
});
