# Email to Azure DevOps — Chrome Extension (local, unpacked)

A toolbar popup that reads the email you have open in Outlook Web and creates
an Azure DevOps work item from it (title, description, and any files you
attach), without needing admin approval — this runs as a locally-loaded
extension, not through your organization's managed add-in store.

## Install (one-time, ~1 minute)

1. Unzip this folder somewhere permanent (don't delete it after — Chrome
   loads the extension from this folder every time).
2. Chrome → `chrome://extensions`
3. Toggle **Developer mode** on (top right).
4. Click **Load unpacked** → select this folder.
5. Pin it: click the puzzle-piece icon in Chrome's toolbar → pin **Email to
   Azure DevOps**.

> If your org has locked down "Developer mode" or unpacked extensions via
> policy, this approach won't work either and you'd need IT to allow it —
> but that's a separate, usually much smaller ask than the add-in store
> approval process.

## Set up your Azure DevOps connection

1. In Azure DevOps: profile icon (top right) → **Personal access tokens** →
   **New Token** → scope **Work Items: Read & Write** → set an expiry → copy it.
2. Open an email in Outlook Web, click the extension icon, expand **Settings**,
   and fill in:
   - **Organization** — the `contoso` part of `dev.azure.com/contoso`
   - **Project**
   - **PAT** from step 1
   - Default work item type
   - Default area path (populated once org/project/PAT are filled in — pick
     whichever one you create work items under most often)
3. **Save**. This is remembered locally on this computer (`chrome.storage.local`)
   — it's not synced to Chrome sync or sent anywhere except directly to
   `dev.azure.com` when you create a work item.
4. The **Area path** dropdown loads automatically once Organization/Project/PAT
   are set, pulling your project's actual area path tree from Azure DevOps.
   It's cached locally for a day so the popup opens instantly; use **↻ Refresh
   area paths** if you've just added/renamed an area in Azure DevOps and want
   the dropdown to pick it up immediately.

## Using it

1. Open/select an email in Outlook Web.
2. Click the extension icon. It tries to auto-fill **Title** and
   **Description** from the open message — review and edit as needed (it's a
   best-effort scrape of Outlook's page, not an official API, so it won't
   always be perfect).
3. Pick an **Area path** from the dropdown if you want one other than your
   configured default (it starts pre-selected to whatever you set as the
   default area path in Settings, but you can change it per work item).
3. If the email has attachments you want on the work item, use the **Attach
   files** picker to select them from disk (Outlook Web doesn't expose
   attachment content to page scripts, so this is a manual step — drag the
   file from the Outlook attachment onto your desktop first, or use Outlook's
   "Save As" on the attachment, then pick it here).
4. **Inline images** in the email body (e.g. a screenshot pasted into the
   message, a logo) are detected automatically and listed under the
   **Include inline images from email body** checkbox, which is on by
   default. Untick it if you don't want them attached.
5. The **Description** field has the signature/sign-off trimmed out by
   default (see below) — untick **Trim signature / sign-off from
   description** to see the untrimmed text if you need something from it.
6. Click **Create work item**. You'll get a link when it's done.

### About signature trimming

This looks for a line that's essentially just a closing greeting — "Best
regards,", "Lep pozdrav," "Mit freundlichen Grüßen," "Sent from my iPhone,"
a lone "--" delimiter, or the start of a confidentiality disclaimer — and
cuts the description there. It's deliberately narrow (matching whole lines
against a known list of sign-offs in English/Slovenian/German/Croatian,
rather than any substring) so it won't accidentally cut a real sentence
that happens to start with "Thanks" or "Regards" mid-paragraph. Still, it's
a heuristic guess, not a parser — if it trims something it shouldn't (or
misses a signature in another language/format), just toggle the checkbox
off to get the full text back and edit by hand.

### About inline image detection

The content script looks at every `<img>` inside the message body, skips
anything that looks like a 1x1 tracking pixel or spacer, **and skips any
image that appears at or after the detected signature line** (so a company
logo or social icons in someone's signature block won't get attached) —
this is checked by actual document position, not just text, so it's
independent of whether the signature was found in the Description text.
It then tries to fetch each remaining image's bytes directly from within
the Outlook page (so it carries your session automatically for images
Outlook itself is hosting). Two limitations worth knowing:
- Images hosted on a different domain than Outlook (e.g. a marketing email's
  remotely-hosted images) will fail this fetch due to normal cross-origin
  restrictions and are silently skipped — only images Outlook serves from
  its own domain reliably come through.
- This wasn't tested against a live email containing an inline image (I
  didn't have one on hand to check against), so treat it as best-effort like
  the rest of the scraping — if it misses an image you expected, the file
  picker in step 3 still works as a manual fallback.

## Why scraping the email is "best effort" (and what's actually confirmed)

Text extracted from the message body is also normalized before it lands in
Description: Outlook Web often wraps each line in its own `<div>` and pads
with non-breaking spaces, which otherwise shows up as doubled blank lines
between paragraphs and stray double spaces. Runs of 2+ blank lines collapse
to one, and runs of spaces/non-breaking spaces collapse to a single space —
without touching real line breaks or intentional structure.

Outlook Web's HTML isn't a published, stable API — Microsoft changes class
names and structure over time, and different tenants can even land on
different domains (this build matches `outlook.office.com`,
`outlook.office365.com`, `outlook.live.com`, and `outlook.cloud.microsoft` —
that last one is a newer domain some tenants get redirected to).

**Confirmed against a live Outlook Web session:**
- Subject: the first `[role="heading"]` inside the `[role="main"][aria-label="Reading Pane"]` landmark. Scoping the search to that landmark (rather than the whole page) is what keeps it from accidentally grabbing sidebar headings like "Navigation pane" or "Inbox".
- Body: the first `[role="document"][aria-label="Message body"]` inside that same landmark. In a conversation thread there are several (one per message) — the first is the currently-open/topmost message.

**Still unverified / best-effort:** attachment name detection. I wasn't able
to find a live email with real attachments during testing to confirm the
selector against, so `findAttachmentNames()` is a reasonable guess at likely
markup, not a confirmed one. It's fine if it comes up empty — the file
picker in the popup works regardless and is what actually gets uploaded. If
you want this tightened up: open an email with an attachment, press F12,
right-click the attachment chip → Inspect, and send me the HTML you see.

If subject/body ever come up blank on your setup (e.g., after a Microsoft
UI update), the same fix applies — inspect and send me the markup.

## Files

```
manifest.json   Extension configuration (MV3)
content.js      Runs on outlook.office.com/live.com — scrapes the open email
popup.html/js/css   The toolbar popup UI and Azure DevOps API calls
icons/          Toolbar icons
```

## Security notes

- The PAT lives only in this browser's local extension storage on this
  machine. Treat it like a password: give it an expiry and rotate it.
- All network calls go straight from your browser to `dev.azure.com` — there
  is no third-party server in this flow.
- If you ever want to share this with teammates, each person should generate
  their **own** PAT rather than sharing one, so Azure DevOps permissions and
  audit history stay per-person.
