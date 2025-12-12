// app.js

window.addEventListener('load', () => {
  const startBtn = document.getElementById('startBtn');
  const stopBtn  = document.getElementById('stopBtn');
  const statusEl = document.getElementById('status');
  const historyEl = document.getElementById('history');
  const langSel = document.getElementById && document.getElementById('langSel');
  const copyBtn = document.getElementById && document.getElementById('copyBtn');

  function updateStatus(s){ if (statusEl) statusEl.textContent = 'Status: ' + s; }
  function addTranscript(t){ if (!historyEl) return; const d=document.createElement('div'); d.textContent=t; historyEl.prepend(d); }

  // If Office is available, use dialog flow. Otherwise use embedded speech (direct page).
  const isOffice = !!(window.Office && Office.context && Office.context.ui && Office.context.document);

  // Helper to write to selected Excel cell
  async function writeToSelectedCell(text) {
    if (!window.Office) { console.warn('Office not available — not writing to Excel'); return; }
    try {
      await Office.onReady();
      Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, function (res) {
        if (res.status === Office.AsyncResultStatus.Failed) {
          console.error('Write failed', res.error && res.error.message);
          updateStatus('Write failed: ' + (res.error && res.error.message));
        } else {
          console.log('Wrote to Excel:', text);
        }
      });
    } catch(e) { console.error('Office write error', e); }
  }

  // DIALOG: open popup and listen for messages from dialog
  let dialog;
  async function openDictationDialog() {
    return new Promise((resolve, reject) => {
      const dialogUrl = 'https://soham1455.github.io/excel-dictation/dialog.html';
      Office.context.ui.displayDialogAsync(dialogUrl, { height: 60, width: 40 }, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error('Dialog open failed', asyncResult.error);
          reject(asyncResult.error);
          return;
        }
        dialog = asyncResult.value;
        // When dialog sends message via messageParent
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
          const message = args.message; // message is the transcript text
          console.log('Dialog transcript:', message);
          addTranscript(message);
          // write to Excel cell
          writeToSelectedCell(message);
        });
        // monitor closure
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
          console.log('Dialog event', arg);
        });
        resolve(dialog);
      });
    });
  }

  // If not Office, keep existing direct-page behaviour (speech in iframe)
  const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
  let recognition;
  if (!isOffice && SpeechRec) {
    recognition = new SpeechRec();
    recognition.interimResults = true;
    recognition.continuous = true;
    recognition.onstart = () => updateStatus('listening...');
    recognition.onend = () => { updateStatus('stopped'); if (startBtn) startBtn.disabled=false; if (stopBtn) stopBtn.disabled=true; };
    recognition.onerror = (e) => updateStatus('error: ' + (e.error||e.message||e.name));
    recognition.onresult = (ev) => {
      let finalText='';
      for (let i = ev.resultIndex; i < ev.results.length; i++) {
        if (ev.results[i].isFinal) finalText += ev.results[i][0].transcript;
      }
      if (finalText) {
        addTranscript(finalText);
      }
    };
  }

  // Start button handler: if running inside Office, open dialog; else use inline recognition
  if (startBtn) {
    startBtn.addEventListener('click', async () => {
      try {
        if (isOffice) {
          updateStatus('opening popup for mic...');
          await openDictationDialog();
          updateStatus('popup open - speak in the popup');
        } else {
          // direct page behavior
          if (!recognition) return updateStatus('SpeechRecognition not supported in this browser.');
          await navigator.mediaDevices.getUserMedia({ audio: true });
          recognition.lang = (langSel && langSel.value) || 'en-US';
          recognition.start();
          startBtn.disabled = true;
          if (stopBtn) stopBtn.disabled = false;
          updateStatus('starting...');
        }
      } catch (e) {
        console.error('Start failed', e);
        updateStatus('Could not start: ' + (e.message||e.name));
      }
    });
  }

  if (stopBtn) {
    stopBtn.addEventListener('click', () => {
      try { if (recognition) recognition.stop(); } catch(e){ console.warn(e); }
      // also try to close dialog if open
      try { if (dialog) dialog.close(); } catch(e){ /* ignore */ }
      updateStatus('stopped');
    });
  }

  // copy button still works on parent page
  if (copyBtn) {
    copyBtn.addEventListener('click', async () => {
      const first = historyEl && historyEl.firstElementChild;
      if (!first) { alert('No transcript yet!'); return; }
      const text = first.textContent;
      try {
        await navigator.clipboard.writeText(text);
        copyBtn.textContent = 'Copied ✓';
        setTimeout(()=> copyBtn.textContent = 'Copy Last Transcript', 1200);
      } catch (err) {
        alert('Auto-copy failed. Select and press Ctrl+C: ' + text);
      }
    });
  }

  updateStatus('idle');
});
