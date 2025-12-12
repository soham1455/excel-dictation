// app.js
console.log('app.js loaded at', new Date().toISOString());
window.addEventListener('load', () => {
  console.log('window.load fired');
  const startBtn = document.getElementById('startBtn');
  const stopBtn = document.getElementById('stopBtn');
  const statusEl = document.getElementById('status');
  const historyEl = document.getElementById('history');
  const langSel = document.getElementById('langSel');

  console.log({startBtn, stopBtn, statusEl, historyEl, langSel});

  if (!startBtn) {
    console.error('startBtn is missing in DOM — check index.html IDs');
    return;
  }

  // minimal working handler so we can test quickly
  startBtn.addEventListener('click', async () => {
    try {
      await navigator.mediaDevices.getUserMedia({audio:true});
      statusEl && (statusEl.textContent = 'Status: listening...');
      console.log('Mic access OK — recognition would start now');
      // For test: append a fake transcript
      const p = document.createElement('div'); p.textContent = 'TEST TRANSCRIPT (mic OK)';
      historyEl && historyEl.prepend(p);
    } catch (e) {
      console.error('Mic request failed', e);
      alert('Mic access failed: ' + (e.message || e.name));
    }
  });

  if (stopBtn) stopBtn.addEventListener('click', ()=>{ statusEl && (statusEl.textContent='Status: stopped'); });
});
