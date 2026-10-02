(() => {
  const maxTranscriptLength = 20 * 1024 * 1024;

  function microsoftUrl(value, base) {
    const url = new URL(value, base);
    const allowedHost = url.hostname.endsWith('.sharepoint.com') || url.hostname.endsWith('.svc.ms');
    if (url.protocol !== 'https:' || !allowedHost || url.username || url.password || url.port) {
      throw new Error('The transcript address is not a supported Microsoft address.');
    }
    return url;
  }

  function itemAddress(value, base) {
    const url = microsoftUrl(value, base);
    const match = url.pathname.match(/^(.*\/_api\/v[\d.]+\/drives\/[^/]+\/items\/[^/]+)/);
    if (!match) return null;
    return new URL(match[1], url.origin).href;
  }

  function itemKey(address) {
    return address?.replace(/\/_api\/v[\d.]+\//, '/_api/');
  }

  function metadata(data) {
    const transcripts = data?.media?.transcripts ?? data?.value;
    if (Array.isArray(transcripts)) {
      return transcripts.find(item => item?.isDefault && item?.temporaryDownloadUrl)
        ?? transcripts.find(item => item?.temporaryDownloadUrl) ?? null;
    }
    return data?.temporaryDownloadUrl ? data : null;
  }

  function milliseconds(offset) {
    if (typeof offset !== 'string') throw new Error('A transcript timestamp is missing.');
    const match = offset.match(/^(\d+):([0-5]\d):([0-5]\d)(?:\.(\d+))?$/);
    if (!match) throw new Error('A transcript timestamp is invalid.');
    const fraction = (match[4] ?? '').padEnd(3, '0').slice(0, 3);
    const value = (Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3])) * 1000 + Number(fraction);
    if (!Number.isSafeInteger(value)) throw new Error('A transcript timestamp is too large.');
    return value;
  }

  function timestamp(value) {
    const hours = Math.floor(value / 3600000);
    const minutes = Math.floor(value / 60000) % 60;
    const seconds = Math.floor(value / 1000) % 60;
    return [hours, minutes, seconds].map(part => String(part).padStart(2, '0')).join(':')
      + '.' + String(value % 1000).padStart(3, '0');
  }

  function escapeText(value) {
    return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function toVtt(text) {
    if (typeof text !== 'string' || text.length > maxTranscriptLength) {
      throw new Error('The transcript is missing or too large.');
    }
    const normalized = text.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
    if (/^WEBVTT(?:[ \t][^\n]*)?\n/.test(normalized)) {
      if (!normalized.includes(' --> ')) throw new Error('This transcript has no captions.');
      return normalized;
    }
    const data = JSON.parse(normalized);
    if (!Array.isArray(data.entries) || !data.entries.length) {
      throw new Error('This recording has no transcript entries.');
    }
    if (data['@odata.nextLink'] || data.nextLink) {
      throw new Error('Only part of the transcript was loaded. Open the full recording in Stream.');
    }
    const cues = data.entries.map((entry, index) => {
      const start = milliseconds(entry.startOffset);
      const end = milliseconds(entry.endOffset);
      if (end <= start || typeof entry.text !== 'string') {
        throw new Error('A transcript entry is invalid.');
      }
      const words = escapeText(entry.text.replace(/\s+/g, ' ').trim());
      const speaker = typeof entry.speakerDisplayName === 'string'
        ? escapeText(entry.speakerDisplayName.replace(/\s+/g, ' ').trim()) : '';
      const caption = speaker ? `<v ${speaker}>${words}</v>` : words;
      return { start, index, text: `${timestamp(start)} --> ${timestamp(end)}\n${caption}` };
    });
    cues.sort((first, second) => first.start - second.start || first.index - second.index);
    return 'WEBVTT\n\n' + cues.map((cue, index) => `${index + 1}\n${cue.text}\n`).join('\n');
  }

  function filename(title) {
    const name = String(title || 'Teams recording').replace(/\.(mp4|vtt)$/i, '')
      .replace(/[<>:"/\\|?*\u0000-\u001F\u007F]/g, '_').replace(/[. ]+$/g, '').trim().slice(0, 160);
    return `${name || 'Teams recording'}-transcript.vtt`;
  }

  globalThis.TeamsVtt = Object.freeze({ microsoftUrl, itemAddress, itemKey, metadata, toVtt, filename, maxTranscriptLength });
})();
