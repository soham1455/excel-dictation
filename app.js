// app.js

window.addEventListener("load", () => {

  const startBtn = document.getElementById("startBtn");
  const stopBtn  = document.getElementById("stopBtn");
  const statusEl = document.getElementById("status");
  const historyEl = document.getElementById("history");
  const langSel = document.getElementById("langSel");
  const copyBtn = document.getElementById("copyBtn");

  function updateStatus(msg) { if (statusEl) statusEl.textContent = "Status: " + msg; }
  function addTranscript(text) {
    if (!historyEl) return;
    const div = document.createElement("div");
    div.textContent = text;
    historyEl.prepend(div);
  }

  // If Office.js exists, use it to write into Excel cells
  async function writeToSelectedCell(text) {
    if (!window.Office) {
      console.warn("Office.js not available; skipping writeToSelectedCell.");
      return;
    }
    try {
      await Office.onReady();
      Office.context.document.setSelectedDataAsync(
        text,
        { coercionType: Office.CoercionType.Text },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Write failed:", asyncResult.error && asyncResult.error.message);
            updateStatus("Write failed: " + (asyncResult.error && asyncResult.error.message));
          } else {
            console.log("Wrote to Excel:", text);
          }
        }
      );
    } catch (e) {
      console.error("Office write error:", e);
    }
  }

  // SpeechRecognition setup
  const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRec) {
    updateStatus("SpeechRecognition not supported. Use Chrome/Edge.");
    if (startBtn) startBtn.disabled = true;
    return;
  }

  let recognition = new SpeechRec();
  recognition.interimResults = true;
  recognition.continuous = true;

  recognition.onstart = () => updateStatus("listening...");
  recognition.onend   = () => {
    updateStatus("stopped");
    if (startBtn) startBtn.disabled = false;
    if (stopBtn) stopBtn.disabled = true;
  };
  recognition.onerror = (e) => {
    console.error("Recognition error:", e);
    updateStatus("error: " + (e.error || e.message || e.name));
  };

  recognition.onresult = (event) => {
    // Collect finals from result batch
    let finalText = "";
    for (let i = event.resultIndex; i < event.results.length; i++) {
      const r = event.results[i];
      if (r.isFinal) finalText += r[0].transcript;
    }
    if (finalText) {
      addTranscript(finalText);
      // write into Excel when possible
      writeToSelectedCell(finalText);
    }
  };

  // Start — explicitly request mic to avoid silent block in iframes
  if (startBtn) {
    startBtn.addEventListener("click", async () => {
      try {
        updateStatus("requesting mic...");
        await navigator.mediaDevices.getUserMedia({ audio: true });
        recognition.lang = langSel.value || "en-US";
        recognition.start();
        startBtn.disabled = true;
        stopBtn.disabled = false;
        updateStatus("starting...");
      } catch (e) {
        console.error("Mic permission or start failed", e);
        updateStatus("Could not start: " + (e.message || e.name));
        alert("Microphone required — please allow microphone permissions for this site.");
      }
    });
  }

  if (stopBtn) stopBtn.addEventListener("click", () => {
    try { recognition.stop(); } catch (e) { console.warn(e); }
  });

  if (copyBtn) {
    copyBtn.addEventListener("click", async () => {
      const first = historyEl && historyEl.firstElementChild;
      if (!first) { alert("No transcript yet!"); return; }
      const text = first.textContent;
      try {
        await navigator.clipboard.writeText(text);
        copyBtn.textContent = "Copied ✓";
        setTimeout(()=> copyBtn.textContent = "Copy Last Transcript", 1200);
      } catch (err) {
        alert("Auto-copy failed. Select the transcript and press Ctrl+C.\n\nTranscript: " + text);
      }
    });
  }

  updateStatus("idle");
});
