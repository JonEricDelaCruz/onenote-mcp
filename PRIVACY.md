# Privacy and data handling

**Short version:** your notes go straight from Microsoft to your own computer.
Nothing is sent anywhere else. There is no server, no account, no analytics, no
tracking, and no telemetry of any kind. Nobody — including the person who wrote
this tool — can see your notes, your Microsoft account, or that you use this at
all.

This document explains exactly how, so you can verify it rather than trust it.

Last updated: 2026-08-11 · Applies to: OneNote for Claude v2.1.1

---

## The permission screen, line by line

When you sign in, Microsoft shows a consent screen. Here's what each line
actually means and why it's needed.

### "View and modify your OneNote notebooks"

> *OneNote for Claude will be able to view the titles of your OneNote notebooks
> and sections, view all pages, modify all pages and create new pages. It cannot
> access password protected sections on your behalf.*

This is the Microsoft permission `Notes.ReadWrite`. It's the narrowest permission
that allows both reading and writing your own notes.

**What it covers:** notebooks you own, in your own account.

**What it does NOT cover:**

- Notebooks other people shared with you
- Group or team notebooks
- Password-protected sections (Microsoft blocks these outright — even with this
  permission, they cannot be opened)
- Your email, calendar, contacts, files, Teams messages, or anything else in your
  Microsoft account

Microsoft's wording says "view all pages" because the permission is granted at
the account level. It does not mean the tool reads everything — it means it
*could* read a page if asked to. In practice it only ever fetches the specific
pages you ask about. See "What actually happens" below.

**Want read-only?** Open Settings → Extensions → OneNote for Claude and turn off
**"Allow Claude to change your notes."** For a stronger guarantee, set the
permissions field to `Notes.Read offline_access` and sign in again — then
Microsoft itself will refuse any write, regardless of what the software does.

### "View your basic profile"

Your name and username, so the tool can tell you which account is signed in.

This one comes from the standard OpenID sign-in scopes (`openid`, `profile`) that
Microsoft includes in every sign-in. **This tool does not request the separate
`User.Read` Graph permission**, and it makes no API call to read your profile —
it reads your name from the sign-in token already on your machine. Your profile
photo, job title, manager, and directory details are never requested or fetched.

### "Maintain access to data you have given OneNote for Claude access to"

This is `offline_access`. It's what lets you sign in once instead of every hour.

Microsoft's wording — *"even when you are not currently using the app"* — sounds
alarming and is worth being precise about. It means the tool holds a **refresh
token** so it can renew its own access without prompting you again. It does
**not** mean anything runs in the background. This software only ever executes
when your AI app launches it, and only does something when you ask it to.

Without this, you'd re-authenticate roughly every hour.

### "unverified"

Microsoft shows this for any app whose publisher hasn't paid for Microsoft
Partner Network verification. It is a statement about a commercial verification
programme, **not** a security finding. Plenty of legitimate open-source tools show
it. The counter-evidence available to you: the entire source code is public and
auditable, which is more than most verified apps offer.

---

## What actually happens when you use it

Every network request this software can make, exhaustively:

| When | Where it goes | What is sent | What comes back |
|---|---|---|---|
| Signing in | `login.microsoftonline.com` | Your sign-in, handled by Microsoft's own library | A token, stored on your machine |
| Listing structure | `graph.microsoft.com` | Your token | Notebook and section names |
| Reading a page | `graph.microsoft.com` | Your token, the page ID | That page's content |
| Reading a page's images | `graph.microsoft.com` | Your token, the image ID | The image file |
| Searching | `graph.microsoft.com` | Your token | Page titles, plus page bodies if you enabled content search |
| Writing | `graph.microsoft.com` | Your token, the text you asked to save | Confirmation |

**That is the complete list.** Two hostnames, both Microsoft's. No third party,
no analytics endpoint, no error reporting service, no update check, no
"anonymous usage statistics."

You can verify this yourself:

```bash
grep -rE "https?://" onenote-mcp.mjs onenote-cli.mjs src/*.mjs
```

The only hosts that appear are `graph.microsoft.com`,
`login.microsoftonline.com`, `localhost`, and documentation links in comments.

### Where your notes go

Microsoft → your computer → the AI app you're using.

The note content becomes part of your conversation, so it's handled under **your
AI provider's** privacy policy (Anthropic's, if you're using Claude) exactly like
anything else you type or paste. That's the one place your note content leaves
your machine, and it happens because you asked a question about it.

**This can include images, but only in specific cases.** Microsoft runs OCR on
your images so OneNote's own search can find them, but does not expose that text
through the API. So the image file itself is passed to your AI app, which reads
it directly. That is how text inside a screenshot becomes available at all.

Because that is both expensive and more of your data leaving your machine, it is
not the default. By default (`auto`):

- Images are downloaded and sent **only when the page has too little text to
  answer from** — under 400 characters, meaning the page is essentially just a
  picture. At most two per read.
- On every other page, **no image is downloaded**. The tool counts them from the
  page HTML it already has and tells your assistant they exist, so you can ask.
  Counting costs no additional request and transmits nothing.

You are always in control:

- *"Read it without the images"* → nothing is downloaded
- Set **When to read images** to `never` in the extension settings → images are
  never fetched or sent under any circumstances
- Set it to `always` → every read includes them

A screenshot sent this way is treated exactly like one you dragged into the chat
yourself. Images go nowhere except your AI provider, and are never written to
disk.

Nothing routes through any server belonging to this tool's author. There is no
such server.

### What is read, and when

Only what your request requires:

- *"What notebooks do I have?"* → names only, no page content
- *"Read my Q3 page"* → that one page
- *"Search for 'budget'"* → page titles; page bodies only if you enable content search
- *"Summarize my recent notes"* → the specific pages it lists for you first

Nothing is bulk-downloaded, indexed, or copied. There is no local database of
your notes.

### Pages you have read are remembered on your computer

This is the one place note content is stored, so it is worth being exact.

When you read a page, the extracted text is kept on your own machine. Reading it
again later reuses that copy instead of downloading and re-processing the page.
That is the whole reason it exists: re-reading an unchanged page is wasted work.

**Where it is kept**

| | |
|---|---|
| **macOS** | `~/Library/Application Support/onenote-mcp/cache/pages-<hash>.json` |
| **Windows** | `%APPDATA%\onenote-mcp\cache\pages-<hash>.json` |
| **Linux** | `~/.config/onenote-mcp/cache/pages-<hash>.json` |

Same protection as your sign-in token: file mode `0600` inside a `0700`
directory, readable only by your user account, written atomically.

**What is in it**

The plain text of pages you asked about — nothing else. No images (those are
never written to disk). No credentials. Not your whole notebook: only pages you
actually opened.

**It is yours alone**

- The `<hash>` in the filename is derived from your Microsoft account ID, so two
  accounts on the same computer cannot read each other's cached pages — and the
  filename itself does not spell out who you are.
- **It is never transmitted anywhere.** `src/cache.mjs` contains no network code
  at all. There is no server to send it to.
- It is not shared between users, machines, or installs. Someone else running
  this tool has their own cache of their own notes, and no path exists between
  the two.

**It cannot go stale.** Before reusing a stored page, the tool compares it
against the page's `lastModifiedDateTime` — a value Microsoft already returned
in the listing it just fetched. If you edited the page by so much as a
character, the stored copy is discarded and the page is re-downloaded. There is
no expiry window during which you might get an old answer.

**Limits.** At most 500 pages or about 4 MB of text, whichever comes first,
after which the least recently used entries are dropped.

**How to see it, wipe it, or turn it off**

```bash
npx @joneericdelacruz/onenote-mcp cache status   # what is stored, and where
npx @joneericdelacruz/onenote-mcp cache clear    # delete all of it
```

Signing out deletes it too — ask Claude to sign out, and the cached text goes
with the credentials.

To stop it being written at all, open Settings → Extensions → OneNote for Claude
and set **Remember pages between sessions** to:

- `memory` — kept only while your AI app is running, nothing written to disk
- `off` — no caching of any kind

**Also, a 60-second memory cache.** Separately from the above, pages read during
a search are held in the running process's memory for 60 seconds so a follow-up
question doesn't re-fetch them. It holds at most 200 pages, is never written to
disk, and vanishes when you quit your AI app.

---

## What's stored on your computer

Two things: your Microsoft sign-in token, and the text of pages you have read.

### Your sign-in token

| | |
|---|---|
| **macOS** | `~/Library/Application Support/onenote-mcp/token-cache.json` |
| **Windows** | `%APPDATA%\onenote-mcp\token-cache.json` |
| **Linux** | `~/.config/onenote-mcp/token-cache.json` |

It's created with permissions `0600` (readable only by your user account) inside
a `0700` directory, and written atomically so it can't be left half-finished. The
file is managed by MSAL, Microsoft's own official authentication library — this
project never handles your password, and never sees it.

**Note text is written to disk in one place only:** the page cache described
above, in your own config folder, which you can inspect, wipe, or switch off. No
index, no database, no logs of content, and nothing written anywhere else.

**To delete it:** ask Claude to sign out, or run `npx @joneericdelacruz/onenote-mcp signout`,
or just delete the file above. Signing out removes the cached page text as well.

### Remembered page text

Covered in full under ["Pages you have read are remembered on your
computer"](#pages-you-have-read-are-remembered-on-your-computer) above: your
config folder, mode `0600`, scoped to your account, never transmitted, wiped by
`cache clear` or by signing out.

---

## Things this tool does not do

Stated plainly, because "we don't track you" is easy to say and worth being
specific about:

- ❌ No analytics or telemetry of any kind
- ❌ No crash or error reporting to any service
- ❌ No usage statistics, anonymous or otherwise
- ❌ No update checks that phone home
- ❌ No advertising or advertising identifiers
- ❌ No user accounts, licence keys, or registration
- ❌ No data sold, shared, or transmitted to any third party
- ❌ No cached note text leaving your machine — the cache is local-only, per
  account, and readable by no one but you
- ❌ No background processes, daemons, or scheduled tasks
- ❌ No network access to any host other than Microsoft's

**Files it touches, completely:** its own token cache (above), and — only if you
run the optional `onenote-cli setup` or `uninstall` command yourself — your AI
app's MCP configuration file, so it can add or remove its own entry. That edit
takes a timestamped backup first and leaves every other entry untouched. The MCP
server itself never touches either file beyond the token cache.

The author receives **nothing**. Not your email, not a count of installs, not an
error message. There is no mechanism by which they could.

The only signal that exists anywhere: Microsoft's admin portal shows the app
owner a **number** — how many accounts have consented. No names, no addresses, no
content, no activity. That's a Microsoft feature, not something this tool reports.

---

## Safety measures

**Deleting a page requires confirmation.** The AI must state the page's exact
current title before a delete goes through. A wrong or stale ID fails safely
rather than destroying the wrong page.

**Write tools can be turned off entirely** in the extension settings, which
disables page creation, editing, and deletion in one switch.

**Page content is treated as untrusted.** Your notes are parsed by a small
purpose-built parser (`src/parse-html.mjs`) that cannot run scripts, load
resources, or make network requests. A note containing malicious HTML cannot do
anything.

**Minimal dependencies.** Four packages, all from official publishers: Microsoft's
authentication library, the Model Context Protocol SDK, a schema validator, and a
browser-opener. None runs code at install time. Automated checks fail the build if
that ever changes.

---

## Verifying any of this yourself

You don't have to take it on faith.

**Read the source.** It's at
[github.com/JonEricDelaCruz/onenote-mcp](https://github.com/JonEricDelaCruz/onenote-mcp).
The runtime is about 3,700 lines of commented JavaScript, plus a similar amount of tests. The files that
matter for privacy:

- `src/onenote.mjs` — every Microsoft API call
- `src/auth.mjs` — sign-in and token storage
- `src/cache.mjs` — what is remembered on disk, and how it is protected
- `onenote-mcp.mjs` — the tools exposed to the AI

**Watch its traffic.** Use Little Snitch (macOS), Wireshark, or any firewall. The
only outbound connections are to Microsoft.

**Check what's stored.** Open the token cache file. It's JSON, and it contains
Microsoft tokens — nothing else.

**Revoke access anytime.** Go to
[microsoft.com/consent](https://microsoft.com/consent), find "OneNote for Claude,"
and remove it. Access stops immediately, regardless of what's on your machine.

---

## Reporting a problem

If you find anything in this document that isn't accurate, that's a bug and it
matters more than a broken feature. Please
[open an issue](https://github.com/JonEricDelaCruz/onenote-mcp/issues).

For something security-sensitive, mark the issue clearly and it'll be handled
first.

---

## Changes to this document

This file is versioned in the repository, so every change to it is publicly
visible in the commit history. If the data handling ever changes, that change is
part of the permanent record.
