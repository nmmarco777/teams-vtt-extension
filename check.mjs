import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';

const read = name => readFile(new URL(name, import.meta.url), 'utf8');
const sandbox = vm.createContext({ URL });
vm.runInContext(await read('transcript.js'), sandbox);
const { toVtt, metadata, microsoftUrl, itemAddress, itemKey, filename } = sandbox.TeamsVtt;
const entries = [
  { startOffset: '00:01:00.000001', endOffset: '00:01:02.25', text: 'Next line', speakerDisplayName: 'Nate' },
  { startOffset: '00:00:59.9999999', endOffset: '00:01:00.0000001', text: 'Hello <team> & all.\n Welcome!', speakerDisplayName: 'A > B' }
];
const vtt = toVtt(JSON.stringify({ entries }));
assert.equal(vtt, 'WEBVTT\n\n1\n00:00:59.999 --> 00:01:00.000\n<v A &gt; B>Hello &lt;team&gt; &amp; all. Welcome!</v>\n\n2\n00:01:00.000 --> 00:01:02.250\n<v Nate>Next line</v>\n');
assert.equal(toVtt('\uFEFFWEBVTT\r\n\r\n00:00:00.000 --> 00:00:01.000\r\nHi\r\n'), 'WEBVTT\n\n00:00:00.000 --> 00:00:01.000\nHi\n');
assert.throws(() => toVtt('WEBVTT\n\n'));
assert.throws(() => toVtt('<html>Sign in</html>'));
assert.throws(() => toVtt(JSON.stringify({ entries: [] })));
assert.throws(() => toVtt(JSON.stringify({ entries, '@odata.nextLink': '/more' })));
assert.throws(() => toVtt(JSON.stringify({ entries: [{ ...entries[0], startOffset: '00:61:00' }] })));
assert.throws(() => toVtt(JSON.stringify({ entries: [{ ...entries[0], endOffset: '00:00:01' }] })));
assert.equal(toVtt(JSON.stringify({ entries: [{ ...entries[0], startOffset: '125:00:00', endOffset: '125:00:01', speakerDisplayName: null }] })), 'WEBVTT\n\n1\n125:00:00.000 --> 125:00:01.000\nNext line\n');
const first = { temporaryDownloadUrl: 'https://tenant.sharepoint.com/transcript-a' };
const preferred = { temporaryDownloadUrl: 'https://tenant.sharepoint.com/transcript-b', isDefault: true };
assert.equal(metadata({ value: [first, preferred] }), preferred);
assert.equal(metadata({ media: { transcripts: [first] } }), first);
assert.equal(metadata(preferred), preferred);
assert.equal(metadata({ value: [] }), null);
for (const address of ['https://sharepoint.com.evil.test/file', 'http://tenant.sharepoint.com/file', 'https://evilsharepoint.com/file', 'javascript:alert(1)', 'https://user@tenant.sharepoint.com/file']) {
  assert.throws(() => microsoftUrl(address));
}
assert.equal(microsoftUrl('https://media.svc.ms/file').hostname, 'media.svc.ms');
const item = 'https://tenant.sharepoint.com/sites/team/_api/v2.1/drives/drive/items/item';
assert.equal(itemAddress(item + '/media/transcripts?token=private'), item);
assert.equal(itemKey(item), itemKey(item.replace('v2.1', 'v2.0')));
assert.equal(filename('Planning / meeting.mp4'), 'Planning _ meeting-transcript.vtt');

const manifest = JSON.parse(await read('manifest.json'));
assert.equal(manifest.manifest_version, 3);
assert.equal(manifest.permissions, undefined);
for (const script of manifest.content_scripts) {
  assert.equal(script.all_frames, true);
  for (const file of script.js) new vm.Script(await read(file));
}

const listeners = [];
const messages = [];
let tick;
const location = { href: 'https://tenant.sharepoint.com/_layouts/15/stream.aspx?id=one', origin: 'https://tenant.sharepoint.com' };
const requests = [];
const window = {
  g_fileInfo: { '.spItemUrl': item.replace('v2.1', 'v2.0'), name: 'Planning.mp4', hasTranscripts: true },
  fetch: async address => {
    requests.push(String(address));
    if (String(address).endsWith('/media/transcripts')) return new Response(JSON.stringify({ value: [first] }));
    return new Response(JSON.stringify({ entries }));
  },
  postMessage: message => messages.push(message),
  addEventListener: (type, listener) => listeners.push(listener)
};
const context = vm.createContext({ window, location, URL, Response, AbortSignal, document: { title: 'Recording', querySelector: () => null }, setInterval: callback => { tick = callback; } });
vm.runInContext(await read('transcript.js'), context);
vm.runInContext(await read('page.js'), context);
const send = data => listeners[0]({ source: window, origin: location.origin, data: { channel: 'teams-vtt-download-v1', ...data } });
send({ type: 'hello' });
assert.equal(messages.at(-1).active, true);
const currentRevision = messages.at(-1).revision;
send({ type: 'download', requestId: 'test', revision: currentRevision });
for (let attempt = 0; attempt < 100 && !messages.some(message => message.type === 'result'); attempt++) {
  await new Promise(resolve => setImmediate(resolve));
}
const result = messages.find(message => message.type === 'result');
assert.equal(result?.vtt, vtt);
assert.equal(result?.filename, 'Planning-transcript.vtt');
assert.equal(requests.length, 2);
assert.equal(requests[0], item + '/media/transcripts');
assert.match(requests[1], /format=json/);
location.href = 'https://tenant.sharepoint.com/sites/team/Documents';
tick();
assert.equal(messages.at(-1).active, false);
send({ type: 'download', requestId: 'stale', revision: currentRevision });
assert.equal(requests.length, 2);
window.g_fileInfo = { '.spItemUrl': item.replace(/item$/, 'other'), name: 'Other.mp4', hasTranscripts: true };
tick();
assert.equal(messages.at(-1).active, true);
assert.ok(messages.at(-1).revision > currentRevision);
console.log('Passed: VTT conversion, metadata, address validation, manifest, download flow, and recording changes.');
