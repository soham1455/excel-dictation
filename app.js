// app.js
let recognition;
const startBtn = document.getElementById('startBtn');
const stopBtn  = document.getElementById('stopBtn');
const statusEl = document.getElementById('status');
const historyEl = document.getElementById('history');
const langSel = document.getElementById('langSel');

function updateStatus(s){ statusEl.textContent = 'Status: ' + s; }
function appendHistory(text) {
  const p = document.createElement('div');
  p.textContent = text;
  historyEl.prepend(p);
}

const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
if (!SpeechRec) {
  updateStatus('Browser does not support Web Speech API. Use Chrome or Edge (Chromium).');
  startBtn.disabled = true;
} else {
  recognition = new SpeechRec();
  recognition.interimResults = true;
  recognition.continuous = true;

  recognition.onstart = () => updateStatus('listening...');
  recognition.onerror = (e) => updateStatus('error: ' + e.error);
  recognition.onend = () => {
    updateStatus('stopped');
    startBtn.disabled = false;
    stopBtn.disabled = true;
  };

  recognition.onresult = (event) => {
    let interim = '';
    for (let i = event.resultIndex; i < event.results.length; ++i) {
      const transcript = event.results[i][0].transcript;
      if (event.results[i].isFinal) {
        appendHistory(transcript);
        writeToSelectedCell(transcript);
      } else {
        interim += transcript;
      }
    }
    updateStatus('listening... (interim: ' + interim + ')');
  };
}

startBtn.addEventListener('click', ()=>{
  try {
    recognition.lang = langSel.value || 'en-US';
    recognition.start();
    startBtn.disabled = true;
    stopBtn.disabled = false;
  } catch (e) {
    updateStatus('Could not start: ' + e.message);
  }
});

stopBtn.addEventListener('click', ()=>{
  recognition.stop();
});

function writeToSelectedCell(text) {
  if (!window.Office) {
    console.warn('Office.js not available; text:', text);
    return;
  }
  Office.onReady().then(() => {
    Office.context.document.setSelectedDataAsync(text, {coercionType: Office.CoercionType.Text},
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error('Write failed: ' + asyncResult.error.message);
          updateStatus('Write failed: ' + asyncResult.error.message);
        } else {
          console.log('Wrote:', text);
        }
      });
  });
}
