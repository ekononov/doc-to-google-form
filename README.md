# doc-to-google-form

A Telegram bot built on [n8n](https://n8n.io) that turns a `.docx` file with a test into a Google Form quiz.

The user sends a `.docx` to the bot. The workflow then:

1. parses the questions, options, correct answers and points;
2. creates a Google Form in quiz mode with auto-grading;
3. creates a Google Spreadsheet for the results and links it to the form;
4. shares both documents with the user's email (editor access);
5. sends the user links to the form and to the results.

## Contents

- [How it works](#how-it-works)
- [Bot commands](#bot-commands)
- [`.docx` format](#docx-format)
- [Repository structure](#repository-structure)
- [Setup](#setup)
- [Parser API](#parser-api)
- [Limitations](#limitations)
- [License](#license)

## How it works

```
Telegram ──► n8n workflow ──► doc-parser (FastAPI)       .docx → JSON
                 │
                 ├──► Google Forms API     create form, add questions (quiz mode)
                 ├──► Google Sheets API    create results spreadsheet
                 ├──► Apps Script API      link the form to the spreadsheet
                 ├──► Google Drive API     share form and spreadsheet with the user
                 └──► Telegram             reply with links
```

Services from `n8n/docker-compose.yml` that this workflow uses:

| Service  | Purpose                                                       |
|----------|---------------------------------------------------------------|
| `n8n`    | Runs the workflow (port `5678`)                               |
| `runners`| Executes Code nodes (external task runner)                    |
| `parser` | `doc-parser`: FastAPI service that parses `.docx` (port 8010, internal) |

The compose file also starts `sandbox-*` and `searxng`. They belong to the n8n AI assistant setup and are not used by this workflow.

Conversation state is stored in two n8n Data Tables:

| Table           | Columns                                          | Purpose                                   |
|-----------------|--------------------------------------------------|-------------------------------------------|
| `quiz_sessions` | `chat_id`, `message`, `command`, `state`, `data`, `key` | Current step of the dialog per chat |
| `user_emails`   | `username`, `email`                              | Telegram username → email for sharing     |

A user is identified by their Telegram **username**.

### Google Form structure

- Quiz mode enabled.
- Question 1: *Class Name* (short text, required).
- Question 2: *Full Name* (short text, required).
- Then one required question per parsed question: radio buttons for a single correct answer, checkboxes for several. Points and correct answers are set from the file.

The bot replies in English or Ukrainian depending on the user's Telegram language (English by default). The form field titles use the same language.

## Bot commands

| Command      | Action                                                             |
|--------------|--------------------------------------------------------------------|
| `/start`     | Asks for the user's email                                          |
| `/set_email` | Asks for the email again (replaces the stored one)                 |
| `/new`       | Starts a new quiz: the bot asks for a name and a `.docx` file      |
| `/reset`     | Intended to clear the chat session. **Not working yet** (see [Limitations](#limitations)) |

Typical flow:

1. Send `/start`, then your email.
2. Send `/new`.
3. Send the `.docx` as a document with the **quiz name in the caption**.
4. Receive links to the form and the results spreadsheet.

## `.docx` format

The parser is heuristic and is not tied to one template. Supported: single-choice and multiple-choice questions; English, Russian and Ukrainian (can be mixed).

**Questions:** `1. ...`, `1) ...`, `Question 1. ...`, a top-level Word list item, or a plain paragraph.

**Options:** `a) ...`, `A. ...`, `(b) ...`, `- ...`, `• ...`, or a nested Word list item.

**Correct answers** (in order of priority):

1. A line after the options: `Answer: a, c`, `Correct answer: 2,3`, or the text of an option.
2. A mark inside the option: `*`, `+`, `✓` at the start or end, `=` before the label (`=c) text`), or `(correct)` at the end.
3. Formatting: bold, underline, highlight or font color on some of the options. If all options are formatted the same way, formatting is ignored.

**Points:** `(2 points)` or `[3 pts]` in the question text, or a separate line `Points: 2`. Default is 1.

**Service lines** (en/ru/uk) are removed: a "question text" header, an "options" header, and a "comment/explanation" block after the options.

**No markers at all:** blocks separated by an empty line (first line is the question, the rest are options).

Example:

```text
1. What is the capital of France? (2 points)
a) Berlin
b) Paris
c) Madrid
Answer: b

2. Which of these are prime numbers?
a) 2
b) 4
c) 7
Answer: a, c
```

## Repository structure

```
.
├── doc_to_google_form_workflow.json   # n8n workflow (import into n8n)
├── doc-parser/                        # .docx → JSON service
│   ├── main.py                        # FastAPI app and parsing logic
│   ├── requirements.txt
│   └── Dockerfile
├── n8n/
│   ├── docker-compose.yml             # n8n, runners, parser, sandbox, searxng
│   ├── .env.sample                    # environment template
│   └── searxng-settings.yml
└── scripts/
    └── docs_link.js                   # Apps Script: link a form to a spreadsheet
```

## Setup

### Prerequisites

- Docker with Compose.
- A Telegram bot token from [@BotFather](https://t.me/BotFather).
- A public HTTPS URL for n8n. The Telegram Trigger uses a webhook, so Telegram must be able to reach n8n.
- A Google account and a Google Cloud project.

### 1. Start n8n and the parser

```bash
cd n8n
cp .env.sample .env
# edit .env: replace the placeholder secrets and set WEBHOOK_URL / N8N_HOST / N8N_PROTOCOL
docker compose up -d --build
```

n8n is available at `http://localhost:5678`. `.env` is git-ignored.

### 2. Import the workflow

In n8n: **Workflows → Import from file** → `doc_to_google_form_workflow.json`.

### 3. Create the Data Tables

Create two Data Tables, all columns of type string:

- `quiz_sessions`: `chat_id`, `message`, `command`, `state`, `data`, `key`
- `user_emails`: `username`, `email`

Then open each Data Table node in the workflow and select these tables. The imported workflow contains table IDs from the original instance.

### 4. Google Cloud and credentials

1. In your Google Cloud project, enable: **Google Forms API**, **Google Sheets API**, **Google Drive API**, **Apps Script API**.
2. Create an OAuth client and an n8n **Google OAuth2 API** credential (this is the `Google account` credential used by the HTTP Request nodes). Authorize it with scopes for Forms, Sheets, Drive and Apps Script execution.
3. Create an n8n **Telegram API** credential with the bot token and attach it to the Telegram nodes.

### 5. Apps Script for linking the form and the spreadsheet

The Forms API cannot set a response destination, so the workflow calls an Apps Script function.

1. Create an Apps Script project (script.google.com) and paste `scripts/docs_link.js`.
2. Link the script to the same Google Cloud project as the OAuth client (**Project Settings → Google Cloud Platform project**).
3. Deploy it as an **API Executable**.
4. In the workflow node **Link Form and Spreadshet**, replace the script ID in the URL (`.../scripts/<SCRIPT_ID>:run`) with yours.

The node calls the script with `"devMode": true`, so it runs the latest saved version of the script. The script must be owned by the same Google account that authorized the n8n credential.

### 6. Activate

Activate the workflow and send `/start` to the bot.

## Parser API

The `parser` service is internal to the compose network (`http://parser:8010`). To run it locally:

```bash
cd doc-parser
pip install -r requirements.txt
uvicorn main:app --port 8010
```

| Method | Path     | Description                                                   |
|--------|----------|---------------------------------------------------------------|
| `GET`  | `/health`| Returns `{"status": "ok"}`                                    |
| `POST` | `/parse` | Multipart field `file` (`.docx`). Query: `default_points` (default `1`). Returns `{"questions": [{"number", "text", "type", "points", "options": [{"text", "correct"}]}]}`. `type` is `single` or `multiple`. |

```bash
curl -F "file=@test.docx" "http://localhost:8010/parse?default_points=1"
```

## Limitations

- Users without a Telegram username cannot use the bot.
- The email is not validated. If it is not a Google-compatible address, sharing fails.
- Only single-choice and multiple-choice questions are supported, and only `.docx` input.
- The quiz name is taken from the document caption. A file sent without a caption fails.
- `/reset` does nothing: the `If row exists` and `Reply after data clear` nodes are disabled and no node deletes the session data on this path.
- The workflow has no error handling branch. If any step fails, the user gets no reply and the session row stays in `quiz_sessions`.
- The workflow hard-codes the parser URL `http://parser:8010/parse`. `PARSER_URL` / `PARSER_PORT` from `.env.sample` are not read by it.

## License

[MIT](LICENSE)
