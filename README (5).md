# Hair Serum Creator Outreach Pipeline

Automates the loop you have been running manually: handle → video URL → transcript → drafted email.

## What it does

For each creator in your input CSV:
1. Pulls the last 10 videos from their TikTok via Apify
2. Picks the best one (hair-relevant, recent, has spoken content)
3. Gets the transcript via the Fusion MCP
4. Drafts a Be Bodywise outreach email using the locked template
5. Writes everything to an XLSX file with two sheets

You review the low-confidence rows on the Review sheet, edit if needed, and feed the Outreach sheet into your mail merge tool.

## Setup

```bash
pip install apify-client anthropic openpyxl
```

You need two API keys:

**Apify** — sign up at apify.com (free tier covers ~16,000 results/month, enough for 1,500+ creators). Get token from console.apify.com → Settings → Integrations.

**Anthropic** — sign up at console.anthropic.com. Generate an API key.

Set them as environment variables:

```bash
export APIFY_TOKEN=apify_api_xxxxxxxxxxxx
export ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxx
```

## Run

Drop your creators into `creators_input.csv` (columns: handle, display_name, email, niche, bio).

```bash
python outreach_pipeline.py --input creators_input.csv --output creators_output.xlsx
```

## Output

`creators_output.xlsx` contains two sheets:

### Sheet 1: Outreach (mail-merge ready)

Exactly the 8 columns your mail merge tool expects:

| Column | Source |
|---|---|
| Email Address | from input |
| Name | display_name from input |
| Subject | generated |
| Body | generated |
| Image | empty (fill in if attaching anything) |
| Status | "Pending" / "DO_NOT_SEND" / "Error" |
| Sent At | empty (your mail merge fills this) |
| From | empty (set globally in your mail merge) |

### Sheet 2: Review (internal QA)

Use this to scan and quality-check before sending. Columns: Handle, Email Address, Name, Niche, Bio, Video URL, Confidence (high/medium/low), Reasoning (one-line note on what X and Y were grounded in), Status.

Sort by Confidence and review `low` rows manually. The `high` ones are usually ready to go as-is.

## Cost per run

Roughly:
- Apify: $0.001 per creator profile (free tier covers this)
- Claude API: ~$0.07 per creator (transcript fetch + email generation)

For 50 creators per round: ~$3.50. For 500 creators: ~$35.

## Notes on video selection logic

Videos are scored on:
- Hair-relevant caption keywords (hair, wig, scalp, edges, etc.)
- Duration 15-120s (short enough to be a real post, long enough to have spoken content)
- Recency (last 90 days, with extra boost for last 30)
- View count (log scale, capped — used as a tiebreaker)

Competitor brands are NOT filtered out. Off-category posts (skincare, fitness, body care) get low keyword scores and naturally lose out to in-category posts.

## When the bio fallback kicks in

If no qualifying video is found (account private, no recent posts, all posts off-category and old), the script generates the email from bio + niche only. These will land as "medium" or "low" confidence on the Review sheet.

## DO_NOT_SEND flag

If a creator's bio reveals they are the founder/CEO of a competing hair growth brand (e.g., @miraclegrowthwater), the email generation step flags them as `DO_NOT_SEND` and skips drafting. Filter these out before importing to your mail merge.

## What this does NOT do

- Send emails — output is XLSX only, plug into Gmail mail merge / Apollo / Streak / Mailshake
- Track replies — that lives in your mail tool
- Re-engage non-responders — separate script if/when needed

## Adding new creators

Just append rows to `creators_input.csv` and rerun. The script is idempotent on the same input (modulo Claude generation variance).
