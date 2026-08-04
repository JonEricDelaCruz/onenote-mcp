# Privacy and data handling

**Short version:** your notes go straight from Microsoft to your own computer.
Nothing is sent anywhere else. There is no server, no account, no analytics, no
tracking, and no telemetry of any kind. Nobody — including the person who wrote
this tool — can see your notes, your Microsoft account, or that you use this at
all.

This document explains exactly how, so you can verify it rather than trust it.

Last updated: 2026-08-04 · Applies to: OneNote for Claude v2.0.0

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

**One exception, stated plainly: a short-lived memory cache.** When you search
inside your notes, the pages that were read are held **in memory for 60
seconds**, so a follow-up question ("now find X in those same notes") doesn't
re-download everything. Specifics:

- It lives in the running process's memory only — **never written to disk**
- It holds at most 200 pages and expires after 60 seconds
- It disappears entirely the moment the process stops, which is when you quit
  your AI app
- It is never transmitted anywhere

If you'd rather it didn't exist at all, it's about fifteen lines in
`src/onenote.mjs` (search for `CONTENT_CACHE_TTL_MS`) and the tool works without
it — just more slowly.

---

## What's stored on your computer

One file: your Microsoft sign-in token.

| | |
|---|---|
| **macOS** | `~/Library/Application Support/onenote-mcp/token-cache.json` |
| **Windows** | `%APPDATA%\onenote-mcp\token-cache.json` |
| **Linux** | `~/.config/onenote-mcp/token-cache.json` |

It's created with permissions `0600` (readable only by your user account) inside
a `0700` directory, and written atomically so it can't be left half-finished. The
file is managed by MSAL, Microsoft's own official authentication library — this
project never handles your password, and never sees it.

**No notes are ever written to disk.** No cache, no index, no logs of content.

**To delete it:** ask Claude to sign out, or run `npx @joneericdelacruz/onenote-mcp signout`,
or just delete the file above.

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
