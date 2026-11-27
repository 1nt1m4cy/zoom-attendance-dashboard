# Zoom Attendance Harvester

A real-time web dashboard that joins a Zoom meeting, extracts participant names and student IDs, and exports a clean Excel attendance sheet. Built for lecturers, TAs, and class reps who need automated attendance without Zoom's paid reports.

---

## What It Does

1. Logs into your Zoom account via a stealth headless browser.
2. Joins the meeting you specify.
3. Continuously scrapes the participant list as people join and leave.
4. Applies smart filtering:
   - **Mixed class** (multiple groups detected): collects only Group A members.
   - **Own class** (single group or no groups): collects only names with a valid student ID.
   - Names that do not meet the active filter criteria are displayed as "Ignored" with a reason.
5. Displays everything on a professional dark-themed dashboard in your browser.
6. Exports a styled `.xlsx` file to your Desktop on save.

---

## Features

<table>
  <tr>
    <td><strong>Real-time dashboard</strong></td>
    <td>Flask + Server-Sent Events. Stats, tables, and activity log update live as participants join.</td>
  </tr>
  <tr>
    <td><strong>Smart group detection</strong></td>
    <td>Automatically switches between group-based and ID-based filtering depending on who is in the meeting.</td>
  </tr>
  <tr>
    <td><strong>Stealth browser</strong></td>
    <td>Anti-detection flags, spoofed navigator properties, human-like email typing. Zoom does not flag the session.</td>
  </tr>
  <tr>
    <td><strong>Credential protection</strong></td>
    <td>Credentials loaded from environment variables, never hardcoded. Password injected via JS value setter (zero keystroke events). Cleared from memory and DOM immediately after login. See <a href="#security-model">Security Model</a>.</td>
  </tr>
  <tr>
    <td><strong>Noise filtering</strong></td>
    <td>Strips Zoom UI text, chat messages, navigation labels, initials, and other non-name artifacts from the participant list.</td>
  </tr>
  <tr>
    <td><strong>All ID delimiters</strong></td>
    <td>Extracts student IDs separated by space, underscore, dash, parentheses, slash, pipe, or comma.</td>
  </tr>
  <tr>
    <td><strong>Scan animation</strong></td>
    <td>Full-screen green scan overlay fires when the bot enters the meeting.</td>
  </tr>
  <tr>
    <td><strong>Search</strong></td>
    <td>Search bars on both Collected and Ignored tables.</td>
  </tr>
  <tr>
    <td><strong>Refresh-safe</strong></td>
    <td>Browser refresh recovers full state (all names, stats, and phase) from the server.</td>
  </tr>
  <tr>
    <td><strong>Cross-platform</strong></td>
    <td>Detects Desktop path on Windows (including OneDrive fallback), macOS, and Linux.</td>
  </tr>
  <tr>
    <td><strong>Styled Excel export</strong></td>
    <td>Color-coded sections, borders, headers. Ready to submit or share.</td>
  </tr>
</table>

---

## Requirements

- Python 3.9 or later
- A Zoom account (free tier works)
- Chromium (installed automatically by Playwright)

---

## Installation

```bash
git clone https://github.com/YOUR_USERNAME/zoom-attendance-dashboard.git
cd zoom-attendance-dashboard

pip install -r requirements.txt
python -m playwright install chromium
```

---

## Configuration

Copy the example environment file and fill in your Zoom credentials:

```bash
cp .env.example .env
```

Edit `.env`:

```
ZOOM_EMAIL=your_zoom_email@example.com
ZOOM_PASSWORD=your_zoom_password
BOT_NAME=Attendance Bot
```

Alternatively, export them directly:

```bash
export ZOOM_EMAIL='your_zoom_email@example.com'
export ZOOM_PASSWORD='your_zoom_password'
```

On Windows (PowerShell):

```powershell
$env:ZOOM_EMAIL = "your_zoom_email@example.com"
$env:ZOOM_PASSWORD = "your_zoom_password"
```

The script will refuse to start if `ZOOM_EMAIL` or `ZOOM_PASSWORD` are not set.

---

## Usage

```bash
python harvester.py
```

A browser tab opens automatically. On the landing page:

1. Paste the Zoom meeting URL.
2. Enter an output filename (optional -- defaults to a timestamped name).
3. Click **Start Harvesting**.

The dashboard shows:

- **Status pill** -- Connecting, Live, Waiting Room, Done, or Error.
- **Timer** -- elapsed time since start.
- **Stats cards** -- In Meeting, Collected, Ignored, Groups.
- **Collected table** -- names that passed the active filter, with student ID badges. Searchable.
- **Ignored table** -- names that did not pass, with the specific reason. Searchable.
- **Activity log** -- timestamped log of every action the bot takes.

When you are done, click **Save & Export**. The `.xlsx` file is saved to your Desktop and downloaded through the browser.

---

## How It Works

```
harvester.py
    |
    +-- Flask (main thread)
    |     |-- GET  /           Dashboard HTML
    |     |-- POST /start      Launches bot thread
    |     |-- GET  /stream     SSE event stream
    |     |-- GET  /state      Full state snapshot (refresh recovery)
    |     |-- POST /stop       Signals bot to stop
    |     |-- POST /save       Generates and saves .xlsx
    |     +-- GET  /download   Serves the saved file
    |
    +-- Playwright (background thread)
          |-- Login (stealth, credential-hardened)
          |-- Join meeting
          |-- Poll participant list every 15s
          |-- Scroll ReactVirtualized list to capture all names
          |-- Extract name, student ID, group tag
          |-- Apply group/ID filter, emit events to SSE queue
          +-- Detect meeting end
```

**Participant extraction** scrolls the virtualized list in steps, reads `span` and `div` text from each visible row, splits on newlines, and feeds each candidate through the noise filter and name cleaner. The cleaner strips IDs (all delimiter formats), group tags, initials, "(Me)", "(Host)", and join/leave messages.

**Filtering logic** runs after every new name:

| Condition | Filter mode | Collected if | Ignored if |
|---|---|---|---|
| Multiple groups (A, B, C, ...) detected | Group-based | Group A members | Any other group, or no group + no ID |
| Only Group A or no groups detected | ID-based | Has a 7-13 digit student ID | No student ID |

The filter is re-evaluated every time a new group tag appears, and both the Collected and Ignored tables are rebuilt from scratch so names move between tables as the classification changes.

---

<h2 id="security-model">Security Model</h2>

Credentials are protected at every layer:

<table>
  <tr>
    <th>Attack vector</th>
    <th>Protection</th>
  </tr>
  <tr>
    <td><strong>Source code inspection</strong></td>
    <td>Credentials are never in the code. They are loaded from environment variables or a <code>.env</code> file, which is excluded from version control via <code>.gitignore</code>.</td>
  </tr>
  <tr>
    <td><strong>Headful browser / screen recording</strong></td>
    <td>The password is set via a JavaScript property setter (<code>HTMLInputElement.prototype.value.set</code>), not via keyboard input. No keystroke events are emitted. The password field displays only the browser's native masked dots. After form submission, the field value is cleared from the DOM.</td>
  </tr>
  <tr>
    <td><strong>Keystroke loggers / input event listeners</strong></td>
    <td>The JS setter approach bypasses <code>keydown</code>, <code>keyup</code>, and <code>keypress</code> events entirely. Only <code>input</code> and <code>change</code> events are dispatched (required for React form binding), and those carry no character data.</td>
  </tr>
  <tr>
    <td><strong>Playwright trace / HAR recording</strong></td>
    <td>Traces are not enabled. The password value is set via <code>page.evaluate</code>, which Playwright records as a JS snippet, not as a typed string. The snippet contains the value at recording time, but tracing is off by default and no trace is saved.</td>
  </tr>
  <tr>
    <td><strong>MITM / Wireshark / network interception</strong></td>
    <td>Zoom login uses HTTPS with certificate verification (Playwright default). The script verifies the page URL contains <code>zoom.us</code> before entering any credential. If the domain does not match (redirect, DNS hijack, proxy interception page), credentials are withheld and the bot aborts.</td>
  </tr>
  <tr>
    <td><strong>Process memory inspection</strong></td>
    <td>The password string is overwritten with dummy bytes and dereferenced immediately after the login form is submitted. <code>gc.collect()</code> is called to accelerate cleanup. The email is also dereferenced after use. No module-level variable holds either value after login completes.</td>
  </tr>
  <tr>
    <td><strong>DevTools / console access in headful mode</strong></td>
    <td>The password input's <code>.value</code> is set to an empty string after form submission. Reading it from the Elements panel or console returns nothing.</td>
  </tr>
  <tr>
    <td><strong>Code modification to exfiltrate credentials</strong></td>
    <td>Cannot be prevented in an open-source model. However, since credentials are loaded from the user's own environment, a modified script can only expose the credentials of the person running it -- credentials they already know. No other user's credentials are at risk.</td>
  </tr>
</table>

---

## Supported Name Formats

The ID extractor handles all common Zoom display name patterns:

```
Elvis Kwashie 1703839183           (space)
Elvis Kwashie_1703839183           (underscore)
Elvis Kwashie (1703839183)         (parentheses)
Elvis Kwashie - 1703839183         (dash with spaces)
Elvis Kwashie-1703839183           (dash)
Elvis Kwashie/1703839183           (slash)
Elvis Kwashie|1703839183           (pipe)
Elvis Kwashie,1703839183           (comma)
EK Elvis Kwashie 1703839183 Grp A (initials + ID + group)
```

Group tags are detected in formats: `Group A`, `Grp A`, `GrpA`, `GA`, or a trailing single letter after the student ID.

---

## Troubleshooting

**"Zoom credentials not found"** -- Set `ZOOM_EMAIL` and `ZOOM_PASSWORD` in your environment or `.env` file. The script will not start without them.

**Bot gets stuck on "Connecting"** -- Zoom may be showing a CAPTCHA or rate-limiting. Wait a few minutes and try again. If running headful (modify `headless=True` to `headless=False` for debugging), watch what the browser shows.

**Names not being captured** -- The participant panel must be open. The bot clicks the Participants button each poll cycle, but Zoom UI changes can break the selector. File an issue with the Zoom client version.

**"Unexpected domain" error** -- The script detected a redirect away from zoom.us before entering credentials. This is a safety check. Verify your meeting URL is correct.

---

## License

MIT
