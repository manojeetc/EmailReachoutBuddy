# EmailRechoutBuddy: Professor Research Outreach Tracker

A **local Flask app** to track professor research outreach from draft through reply. Includes **18 contacts imported** from your uploaded `Om_Research_Outreach_Columbus_Colleges.xlsx` Faculty Directory, with their research fit, scores, source links and caveats. All professor data is editable. **No emails are automatically sent.**

## Features

- Dashboard: professors, contacted, confirmed replies, follow-ups due and possible replies.
- Research-oriented contact profiles with university, lab research, relevance, priority, verification and private notes.
- Editable email templates; save drafts; review exact subject/body/recipient before clicking Send.
- Sends via Microsoft Graph authenticated as your `@live.com` account; records date, content, and conversation ID.
- Manual follow-up draft generation, default due after 7 days; never auto-sends.
- Sync your inbox on demand; exact Graph conversation matches are marked as replies; other messages from matching professor addresses are flagged **possible** until you review.
- Import-compatible Excel export (using the same headers as Excel import), Supabase PostgreSQL database, encrypted persistent MSAL token cache, workspace-password protection, CSRF checking.
- Professor photo (paste an image URL; initials shown as fallback), a manually set 13-step **Status** pipeline, and a separate **Email Status**.
- AI drafting via the OpenAI API (`OPENAI_API_KEY`, `OPENAI_MODEL`, `SENDER_NAME`, `SENDER_PROFILE` in `.env`): generates the first email, a follow-up, or a reply from the email chain and research profile. Drafts are never sent automatically.
- Reply to a professor's message from the timeline (sent in the same Outlook thread). Run **Sync inbox** to load full message text.

## 1. Register Microsoft OAuth

In Microsoft Entra App registrations, select your **personal Microsoft accounts** app:

1. **Authentication** → Add a platform → **Web** → `http://localhost:5000/callback` (do not register under Mobile/Desktop). If the URI already exists on another platform, remove that copy first.
2. **Certificates & secrets** → New client secret → copy the **Value** (not Secret ID).
3. **API permissions** → Microsoft Graph → **Delegated**: `User.Read`, `Mail.Read`, `Mail.Send`, `Mail.ReadWrite`. `offline_access` is requested by the OAuth flow for refresh tokens. For Microsoft Graph, ensure users can consent to required delegated permissions.
4. Copy your Application (client) ID from the Overview page.

Personal `@live.com` works with **delegated** OAuth login. Application-only mail access is not used.

## 2. Supabase setup, data migration, and run

1. **Rotate any database password shared in chat.** It must be considered compromised. Reset it in the Supabase Dashboard before connecting this app. Do not send the replacement password in chat or commit it.
2. Copy the connection host/port/database/user from **Project Settings → Database → Connect**. The project may require its session pooler settings if a direct connection is unavailable.
3. Add either `SUPABASE_DB_CON_STRING` with the **Session Pooler** URI, or the individual `SUPABASE_DB_HOST`, `SUPABASE_DB_PORT`, `SUPABASE_DB_NAME`, `SUPABASE_DB_USER`, and `SUPABASE_DB_PASSWORD` values to the existing `.env` (keep the generated app secrets already there). The direct database host can be IPv6-only; the Session Pooler is the recommended option when direct DNS/connection fails. Never share the connection string or put a database password or service-role key in browser code. The `sb_publishable_...` key is not a database password and is not used by this server-side Flask app.
4. Install the driver, apply [`supabase/schema.sql`](supabase/schema.sql) through the Supabase SQL Editor or with `python apply_supabase_schema.py`, then move the existing SQLite records (professors, drafts, email history, inbox history, and settings) to Supabase. Run these commands **before** starting the app:

	```sh
	source .venv/bin/activate
	python -m pip install -r requirements.txt
	python apply_supabase_schema.py
	python migrate_sqlite_to_supabase.py
	```

	The migration is safe to rerun: existing rows are not overwritten, and the local `outreach.sqlite3` is left unchanged. It copies the encrypted Outlook token cache from `settings`; keeping the same `TOKEN_ENCRYPTION_KEY` preserves the ability to use it.

5. Start the website:

	```sh
	python app.py
	```

Browse **http://localhost:5000**, enter your **APP_PASSWORD**, then click **Connect Microsoft Outlook** and sign in as `manojeet@live.com`.

## 3. Daily workflow

1. Open a professor. Verify the address and research details before sending (imported spreadsheet information is not automatically revalidated).
2. Edit the outreach email placeholders and click **Save draft**.
3. Click **Review & send**. Confirm recipient, subject and body, then **Send via Outlook**.
4. Click **Sync inbox** to look for replies. Matches on the original Graph `conversationId` count as confirmed. Other messages from that professor received after outreach are **possible** and should be checked. Follow-up due reflects last sent date and a customizable delay; there is no background scheduler.
5. Create an editable follow-up draft as needed. Sending follow-ups uses a new message with `Re:` subject, not guaranteed to thread in every Outlook client.

## Limitations and operational notes

- **Manual inbox sync** only; no automatic scheduled polling or live webhook. Inbox sync searches **up to 2,000 newest Inbox messages per click**. Emails in Junk or other folders, or older than this cap, may not be detected.
- A matched sender with a different conversation ID may be unrelated: review possible replies. A professor using a different address will not match. Microsoft Graph may not preserve the same conversation ID across services; follow-up matching is best-effort.
- Graph accepting a send request does not guarantee final delivery. Confirm delivery/failed delivery via Outlook.
- The app creates a Graph draft, then asks Graph to send it. If the send succeeds but the network response fails, the tracker may be uncertain; **check Outlook Sent Items before retrying**.
- One local workspace/user only. Keep the Flask server bound to localhost. For public hosting, implement proper user authentication, HTTPS, a persistent session store and additional rate limiting; do not expose this dev app to the public internet.
- `.env` and `outreach.sqlite3` contain sensitive secrets/data; keep them private and back up securely. Changing `TOKEN_ENCRYPTION_KEY` makes stored OAuth tokens unreadable (disconnect/reset as needed).
- There is no automatic deletion or deduplication across contacts besides unique email addresses.
- By default the app seeds contacts only when the database is first created. Existing changes are not overwritten on restart.
- Data reflects the user's spreadsheet, not independently verified professor availability, role, or current contact details.

## Data storage

Supabase PostgreSQL is the live backend. The schema is in [`supabase/schema.sql`](supabase/schema.sql); the original SQLite file remains as a local backup after migration. Do not delete it until you have verified the imported records in the website and Supabase. `seed_professors.json` is the spreadsheet-import snapshot; the original workbook isn't modified.
