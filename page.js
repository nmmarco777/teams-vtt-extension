(() => {
  const channel = 'teams-vtt-download-v1';
  const helpers = globalThis.TeamsVtt;
  const originalFetch = window.fetch;
  let pageAddress = location.href;
  let context = null;
  let downloadUrl = null;
  let revision = 0;
  let downloading = false;
  let lastFileAddress = null;

  function post(type, values = {}) {
    window.postMessage({ channel, type, revision, ...values }, location.origin);
  }

  function publish() {
    post('state', { active: Boolean(context) });
  }

  function setContext(address, title) {
    const key = helpers.itemKey(address);
    if (context?.key === key) {
      context.title = title || context.title;
      return;
    }
    revision++;
    context = { address, key, title: title || document.title || 'Teams recording' };
    downloadUrl = null;
    publish();
  }

  function syncPage() {
    if (location.href !== pageAddress) {
      pageAddress = location.href;
      revision++;
      context = null;
      downloadUrl = null;
      publish();
    }
    const info = window.g_fileInfo;
    if (!info?.['.spItemUrl']) return;
    try {
      const address = helpers.itemAddress(info['.spItemUrl'], location.href);
      const isVideo = /\.(mp4|webm|mov)$/i.test(info.name || '')
        || typeof info.hasTranscripts === 'boolean' || Boolean(document.querySelector('video'));
      if (address && isVideo && address !== lastFileAddress) {
        lastFileAddress = address;
        setContext(address, info.name || info.displayName || info.title);
      }
    } catch {
      // Other SharePoint files may expose a different address format.
    }
  }

  function rememberMetadata(data) {
    const transcript = helpers.metadata(data);
    if (!transcript) return;
    downloadUrl = helpers.microsoftUrl(transcript.temporaryDownloadUrl, location.href).href;
    publish();
  }

  async function inspectResponse(response, requestedUrl, requestRevision, requestedPage) {
    try {
      const url = helpers.microsoftUrl(requestedUrl, requestedPage);
      if (!/transcripts/i.test(url.href) || /\/cdnmedia\//i.test(url.pathname) || !response.ok) return;
      const address = helpers.itemAddress(url.href);
      if (!address || requestedPage !== location.href || requestRevision !== revision) return;
      const data = await response.clone().json();
      if (requestedPage !== location.href || requestRevision !== revision || !helpers.metadata(data)) return;
      setContext(address);
      rememberMetadata(data);
    } catch {
      // Transcript discovery must not change the player's response or errors.
    }
  }

  window.fetch = function (...args) {
    syncPage();
    const requestRevision = revision;
    const requestedPage = location.href;
    const request = args[0];
    const requestedUrl = typeof request === 'string' || request instanceof URL ? String(request) : request?.url;
    const promise = originalFetch.apply(this, args);
    promise.then(response => inspectResponse(response, requestedUrl, requestRevision, requestedPage), () => {});
    return promise;
  };

  async function fetchText(url) {
    const response = await originalFetch.call(window, url, {
      credentials: 'same-origin', signal: AbortSignal.timeout(20000)
    });
    if (response.status === 401 || response.status === 403) {
      throw new Error('Microsoft denied access. Open the Transcript panel, then try again.');
    }
    if (!response.ok) throw new Error(`Microsoft returned error ${response.status}. Refresh the recording and try again.`);
    if (Number(response.headers.get('content-length')) > helpers.maxTranscriptLength) {
      throw new Error('The transcript is too large.');
    }
    const text = await response.text();
    if (text.length > helpers.maxTranscriptLength) throw new Error('The transcript is too large.');
    return text;
  }

  async function download(requestId, expectedRevision) {
    syncPage();
    if (downloading || !context || expectedRevision !== revision) return;
    downloading = true;
    const startedRevision = revision;
    const startedPage = location.href;
    const currentContext = context;
    try {
      // Refresh short-lived links on every click when cookie access is available.
      let metadataText;
      try {
        const metadataAddress = currentContext.address.replace(/\/_api\/v[\d.]+\//, '/_api/v2.1/');
        metadataText = await fetchText(`${metadataAddress}/media/transcripts`);
      } catch (error) {
        if (!downloadUrl) throw error;
      }
      if (startedRevision !== revision || startedPage !== location.href) throw new Error('The recording changed. Try again.');
      if (metadataText) {
        const data = JSON.parse(metadataText);
        if (!helpers.metadata(data)) {
          downloadUrl = null;
          throw new Error('No transcript is available.');
        }
        rememberMetadata(data);
      }
      const url = helpers.microsoftUrl(downloadUrl, location.href);
      url.searchParams.set('format', 'json');
      const vtt = helpers.toVtt(await fetchText(url.href));
      syncPage();
      if (startedRevision !== revision || startedPage !== location.href) throw new Error('The recording changed. Try again.');
      post('result', { requestId, vtt, filename: helpers.filename(currentContext.title) });
    } catch (error) {
      let message = error.message;
      if (error.name === 'TimeoutError') message = 'Microsoft took too long to respond. Try again.';
      if (error.name === 'SyntaxError') message = 'Microsoft returned an unsupported transcript format.';
      if (error.name === 'TypeError') message = 'Could not reach the transcript. Open the Transcript panel or refresh the recording, then try again.';
      post('result', { requestId, error: message });
    } finally {
      downloading = false;
    }
  }

  window.addEventListener('message', event => {
    if (event.source !== window || event.origin !== location.origin || event.data?.channel !== channel) return;
    if (event.data.type === 'hello') {
      syncPage();
      publish();
    }
    if (event.data.type === 'download' && typeof event.data.requestId === 'string') {
      void download(event.data.requestId, event.data.revision);
    }
  });
  setInterval(syncPage, 1000);
})();
