# Factory Access DB → Odoo 15 Sync Agent

Reads new rows from `full_data` inside `pro.accdb` (password `2214`) and
pushes them as JSON to an Odoo 15 endpoint. Runs as a single standalone
`.exe` — **no Python required on the factory computer.**

---

## 1. How it decides what's "new" (the triggering logic)

MS Access has no external "row was just inserted" event you can listen
to from outside the application that owns the file — and modifying
whatever software currently writes into `pro.accdb` to fire such an
event would be invasive and risky.

Instead, this agent **polls by ID**:

1. It remembers the highest row ID already sent, in `sync_state.json`.
2. Every `poll_interval_seconds` (default: 60), it queries for rows with
   a higher ID than that.
3. It sends only the new rows to Odoo.
4. It only updates the remembered ID **after** Odoo confirms receipt.

This means: if the factory PC restarts, if Odoo is briefly unreachable,
or if the script crashes mid-cycle — nothing is lost and nothing is
ever sent twice. It also opens the database **read-only**, so it won't
fight for a file lock with whatever software is already using
`pro.accdb`.

If `full_data` doesn't have a numeric ID column, tell me the actual
column names and I'll switch the logic to a timestamp column instead —
same idea, just compares dates instead of IDs.

---

## 2. Prerequisites on the factory PC

- **Microsoft Access Database Engine Redistributable** must be present.
  - If Microsoft Office (with Access) is already installed on that PC,
    you almost certainly already have this — skip this step.
  - If not, download "Microsoft Access Database Engine 2016
    Redistributable" from Microsoft's site. It's free, lightweight, and
    needs no Office license. Match the **bitness** (32-bit vs 64-bit)
    to how you build the `.exe` below.
- Network access from the factory PC to your Odoo server.

## 3. Building the standalone .exe — no Windows PC needed

You have two options. **Option A is recommended if you don't own a
Windows machine.**

### Option A — GitHub Actions (builds it in the cloud, free)

This folder already includes `.github/workflows/build.yml`, which tells
GitHub to spin up a temporary Windows machine, build the exe there, and
let you download the finished file. You never touch Windows yourself.

1. Create a free GitHub account if you don't have one (github.com).
2. Create a new **private** repository (e.g. `factory-sync`).
3. Upload the entire contents of this `factory_sync` folder into that
   repo (drag-and-drop works fine on github.com, or use `git push`).
4. Go to the repo's **Actions** tab. The build starts automatically
   (or click **"Build sync_agent.exe" → Run workflow** if it doesn't).
5. Wait ~2 minutes. Click into the finished run → scroll to
   **Artifacts** → download **sync_agent-exe**.
6. Unzip it — inside is `sync_agent.exe`. That's the file you copy to
   the factory PC.

Whenever you change `sync_agent.py` in the future, just push the
update to GitHub and a fresh exe builds automatically.

### Option A2 — GitLab CI (same idea, if you prefer GitLab)

This folder also includes `.gitlab-ci.yml` for the same purpose.

1. Create a free account at gitlab.com if needed.
2. Create a new **private** project.
3. Upload everything from this `factory_sync` folder into it, making
   sure `.gitlab-ci.yml` ends up at the project **root**.
4. Go to **Build → Pipelines** — it runs automatically.
5. Open the finished job → **Job artifacts** panel → download
   `dist/sync_agent.exe`.

Note: GitLab's Windows runners are currently in beta and can be a
little slower to start than GitHub's, but work the same way.

### Option B — Build it yourself on any Windows PC with Python

```bash
cd factory_sync
pip install -r requirements.txt
pyinstaller --onefile --name sync_agent --console sync_agent.py
```

Or double-click `build.bat` on that machine. Produces
`dist/sync_agent.exe`.

## 4. Deploying to the factory PC

1. Create a folder, e.g. `C:\FactorySync\`
2. Copy into it:
   - `sync_agent.exe`
   - `config.json`
3. Edit `config.json`:
   - `access_db_path` → full path to `pro.accdb` on that PC
   - `odoo_url` → your real Odoo server address + `/api/full_data/import`
   - `odoo_api_key` → a long random string (set the **same** value in
     Odoo — see the module README)
   - `id_column` → confirm this matches the actual primary key column
     name in `full_data` (commonly `ID`)
4. Double-click `sync_agent.exe` to test it. Check `sync_agent.log` in
   the same folder to confirm it's polling and sending successfully.

## 5. Running it automatically (pick one)

**Option A — simplest: Startup folder**
Press `Win+R`, type `shell:startup`, drop a shortcut to
`sync_agent.exe` in there. It launches every time the PC logs in and
keeps running (it has its own internal loop — no need to relaunch it).

**Option B — Task Scheduler, "At log on" trigger**
More control over restarts: Task Scheduler → Create Task → Trigger:
"At log on" → Action: start `sync_agent.exe`. Tick "Restart if it fails"
under Settings for extra resilience.

**Option C — Windows Service (most robust, for 24/7 unattended PCs)**
Use a tool like NSSM (`nssm install FactorySync`) to wrap
`sync_agent.exe` as a proper Windows service that auto-restarts on
crash and starts before any user logs in. Recommended if this PC is a
dedicated, always-on machine.

## 6. Files this creates next to the exe

- `sync_agent.log` — rotating log (5MB × 5 files), check this first if
  something looks wrong
- `sync_state.json` — remembers the last synced row ID; delete it only
  if you intentionally want to re-send everything from scratch

## 7. Security notes

- The Access password and the Odoo API key both live in plain text in
  `config.json`. Restrict folder/file permissions on the factory PC so
  only the account running the agent can read it.
- The Odoo endpoint checks the `X-API-KEY` header against a value
  stored in Odoo's System Parameters — see the module README for setup.
