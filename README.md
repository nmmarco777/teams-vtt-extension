# Teams VTT Download

A small Chrome extension that adds **Download VTT** to a recording page. VTT is a caption file containing the words, speaker names, and timestamps.

## Install

1. Open `chrome://extensions` in Chrome.
2. Turn on **Developer mode**.
3. Click **Load unpacked** and select this `teams-vtt-extension` folder.
4. Refresh any recording pages that were already open.

No build step or package installation is needed. Requires Chrome 111 or later.

## Use

1. Open a Teams recording in Chrome, either inside Teams or on its SharePoint/Stream page.
2. Click **Download VTT** in the bottom-right corner of the recording player.
3. If access fails, open Microsoft's **Transcript** panel and try again. The extension can use the transcript link loaded by the player.

Chrome saves a file such as `Planning-transcript.vtt`. Chrome's download settings determine whether it asks where to save it.

## Scope

- Supports the SharePoint/OneDrive Stream player, including matching frames embedded in Teams in Chrome.
- Runs on `*.sharepoint.com`, `teams.microsoft.com`, and `teams.cloud.microsoft`. The button appears only after a recording is identified.
- Uses the default transcript language when several transcripts are available.
- Converts Microsoft's full JSON transcript to VTT, preserving speaker names and timestamps. Existing VTT responses are saved directly.
- Uses your existing Microsoft session. No separate sign-in, application registration, external server, analytics, or persistent transcript storage.
- Reads transcript metadata responses and requests transcript content when you click. Does not request video or audio files.
- Does not generate a transcript for a meeting that was never transcribed.

This first version targets the Stream player used for recordings. A Teams recap that shows only a transcript without loading a Stream player is not supported. The Teams desktop application, retired Stream Classic, government cloud domains, and Microsoft Defender proxy domains are also outside this version's scope. Guest sessions may need the native Transcript panel to load the signed transcript link first. Microsoft must permit the transcript request.

Microsoft's internal player endpoints may change. This extension has passed local checks and a simulated recording-page test. A real signed-in player was inspected, and transcript discovery was adjusted for its API address format. The extension itself still needs a real download test in Chrome. The in-app browser's inspection controls cannot install or run a Chrome extension.

## Check

Run `node check.mjs` from this folder. It checks transcript conversion, metadata formats, unsafe addresses, the manifest, the download flow, and clearing the old transcript when the recording changes.

To verify in Chrome, load the extension, open a transcribed recording, download the VTT, and check the final caption against the end of the native transcript. Then switch to another recording and confirm the filename and text change. Also try a recording without a transcript.

## References

- [Microsoft recording storage and permissions](https://learn.microsoft.com/en-us/microsoftteams/tmr-meeting-recording-change)
- [Chrome content scripts](https://developer.chrome.com/docs/extensions/develop/concepts/content-scripts)
- [Stream endpoint observations](https://github.com/brendangooden/ms-teams-sharepoint-downloader)

The endpoint observations informed compatibility handling. This extension's source was written independently.
