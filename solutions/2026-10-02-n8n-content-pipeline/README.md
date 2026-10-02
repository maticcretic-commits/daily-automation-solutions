# BlogFlow — n8n Content Pipeline Automation

A production-ready **content publishing pipeline**: a queue (Google Sheets tabs)
feeds an AI article generator (Claude/OpenAI), which auto-publishes to WordPress —
with the resilience rules real clients demand. Built as the demo solution for a
client brief asking for an n8n multi-workflow content automation system.

**What it automates** (the classic 9-workflow pattern, core engine here):
- **Queue** — content ideas live in a sheet; the pipeline pulls **oldest first**
- **AI generation** — article body written by an LLM via API
- **Publishing** — immediate `publish` to WordPress via REST (never draft), with
  SEO title + meta description (Rank Math style) and **image alt text = focus keyword**
- **Sitemap ping** after every publish
- **Downstream queues** — every published post is appended to the social
  (Pinterest/LinkedIn) and newsletter (MailerLite) queues for follow-up workflows
- **Resilience** — daily publish limit read from the Settings tab (never
  hardcoded), failed items retry up to 3 times then are marked Failed, and **one
  item's failure never stops the rest of the run**

## Files

| File | What it is |
|---|---|
| `blogflow.py` | The engine: queue store, AI + WordPress clients, pipeline. Stdlib only. |
| `workflow-n8n.json` | The same pipeline as an **importable n8n workflow** (14 nodes). |
| `test_blogflow.py` | 19 tests covering FIFO order, limits, retries, failure isolation, SEO rules. |

## Use it

```bash
# 1. Run the test suite
python3 test_blogflow.py

# 2. Enqueue content ideas (KIND KEYWORD FOCUS-KEYWORD)
python3 blogflow.py --enqueue article "compound interest" "compound interest"
python3 blogflow.py --enqueue calculator "sip calculator" "sip calculator"

# 3. Run one daily publish cycle (respects daily_limit from settings)
python3 blogflow.py

# 4. Preview what a run would publish (no side effects)
python3 blogflow.py --dry-run

# 5. Daily summary report
python3 blogflow.py --summary
```

The store lives in `./blogflow-data/` — five CSVs mirroring the Google Sheet
tabs (`queue`, `settings`, `log`, `social_queue`, `newsletter_queue`).

### Real API mode

Swap the mock clients for the real ones (keys via environment):

```bash
export ANTHROPIC_API_KEY="sk-ant-..."   # Claude article generation
export WP_USER="editor" WP_APP_PASSWORD="xxxx xxxx xxxx xxxx"  # WordPress
python3 - <<'EOF'
from blogflow import QueueStore, ClaudeClient, WordPressClient, run_pipeline
store = QueueStore("./blogflow-data")
wp = WordPressClient("https://your-site.com")
run_pipeline(store, ClaudeClient(), wp)
EOF
```

### Import into n8n

Workflows → Import from File → choose `workflow-n8n.json`. Reconnect the four
credential placeholders (Google Sheets OAuth, Anthropic API key, WordPress app
password), set the `WP_SITE_URL` env var, paste your sheet ID, and activate the
daily schedule. The sticky note inside the workflow documents every rule.

## Behavior contract (what a client can verify)

1. Oldest queued item publishes first (FIFO).
2. `daily_limit` in Settings is the only throttle — change it without touching code.
3. An item that fails is retried on later runs; after 3 attempts it is marked Failed.
4. A failing item never blocks siblings — the run always completes.
5. Every publish sets SEO title + meta description and `alt_text = focus_keyword`.
6. The sitemap is pinged after every publish.

## License

MIT — free for commercial use.
