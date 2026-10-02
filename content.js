(() => {
  const channel = 'teams-vtt-download-v1';
  let revision = -1;
  let pendingRequest = null;
  let timeout = null;
  const host = document.createElement('div');
  host.hidden = true;
  host.style.cssText = 'position:fixed;right:20px;bottom:20px;z-index:2147483647;';
  const panel = host.attachShadow({ mode: 'open' });
  panel.innerHTML = `
    <style>
      :host { color-scheme: light; font: 14px/1.4 system-ui, sans-serif; }
      section { width: 260px; padding: 14px; border: 1px solid #d8d8e6; border-radius: 12px;
        background: white; color: #242424; box-shadow: 0 4px 24px #0002; }
      strong { display: block; margin-bottom: 10px; }
      button { width: 100%; padding: 10px; border: 0; border-radius: 7px;
        background: #464eb8; color: white; font: inherit; font-weight: 600; cursor: pointer; }
      button:focus-visible { outline: 3px solid #242424; outline-offset: 3px; }
      button:disabled { opacity: .65; cursor: wait; }
      p { margin: 10px 0 0; font-size: 12px; overflow-wrap: anywhere; }
      @media (max-width: 340px) { section { width: 210px; } }
    </style>
    <section aria-label="Teams transcript download">
      <strong>Recording transcript</strong>
      <button type="button">Download VTT</button>
      <p role="status" aria-live="polite">Open the Transcript panel if the download is unavailable.</p>
    </section>`;
  const button = panel.querySelector('button');
  const status = panel.querySelector('p');
  document.documentElement.append(host);

  function finish(message) {
    clearTimeout(timeout);
    pendingRequest = null;
    button.disabled = false;
    button.textContent = 'Download VTT';
    status.textContent = message;
  }

  button.addEventListener('click', () => {
    pendingRequest = crypto.randomUUID();
    button.disabled = true;
    button.textContent = 'Downloading…';
    status.textContent = 'Getting the transcript from Microsoft…';
    timeout = setTimeout(() => finish('No response. Refresh the recording and try again.'), 45000);
    window.postMessage({ channel, type: 'download', requestId: pendingRequest, revision }, location.origin);
  });

  window.addEventListener('message', event => {
    if (event.source !== window || event.origin !== location.origin || event.data?.channel !== channel) return;
    const data = event.data;
    if (data.type === 'state' && Number.isInteger(data.revision) && data.revision >= revision) {
      if (data.revision !== revision) finish('Open the Transcript panel if the download is unavailable.');
      revision = data.revision;
      host.hidden = !data.active;
      if (!host.isConnected) document.documentElement.append(host);
      return;
    }
    if (data.type !== 'result' || data.requestId !== pendingRequest || data.revision !== revision) return;
    if (typeof data.error === 'string') {
      finish(data.error);
      return;
    }
    if (typeof data.vtt !== 'string' || !data.vtt.startsWith('WEBVTT') || data.vtt.length > 20 * 1024 * 1024) {
      finish('Microsoft returned an unsupported transcript.');
      return;
    }
    const filename = typeof data.filename === 'string' ? data.filename : 'Teams-transcript.vtt';
    const url = URL.createObjectURL(new Blob([data.vtt], { type: 'text/vtt;charset=utf-8' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = filename.replace(/[<>:"/\\|?*\u0000-\u001F\u007F]/g, '_').slice(0, 200);
    document.documentElement.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    finish('Transcript sent to your browser’s Downloads folder.');
  });

  window.postMessage({ channel, type: 'hello' }, location.origin);
})();
