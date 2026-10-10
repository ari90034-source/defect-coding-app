# Defect Coding App

A Python version of the Defect Coding Dashboard. Each month you upload the defect file, Claude suggests a defect code for every row, you decide on the red and yellow rows, and the month is saved to the history. The monthly summary updates automatically and downloads as Excel.

It comes with the defect code list and coding rules. The history starts empty: load your past months with **History → Import a history workbook** (one tab per month, code in column E), or code and save months as you go.

## What's in the folder

| File | What it is |
|---|---|
| `app.py` | The web interface (Streamlit) |
| `defect_core.py` | The logic: reading files, auto-coding, summary, Excel exports. No interface code, so it can also be used from a script. |
| `data/history.json` | Every saved month and its coded rows |
| `data/settings.json` | The defect code list and coding rules |
| `requirements.txt` | Python packages to install |
| `.streamlit/secrets.toml.example` | Template for your API key and optional editor password |

## Setup (one time)

1. Install Python 3.10 or newer.
2. Open a terminal in this folder and install the packages:
   ```
   pip install -r requirements.txt
   ```
3. Get an Anthropic API key from https://console.anthropic.com (API usage is billed separately from a Claude.ai subscription).
4. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and paste in your key. You can instead set `ANTHROPIC_API_KEY` as an environment variable.

Without an API key the app still works: rows that match a past description word for word are coded automatically, and you can upload a file that was coded elsewhere (for example, in the Defect Analysis project) to review and save it.

## Running it

```
streamlit run app.py
```

It opens in your browser at http://localhost:8501.

## Each month

1. **Code a month** tab: upload the month's Excel file (part number, date, program, description in columns A–D). Check the month and year it detected.
2. Select **Auto-code**. Rows matching a past description are coded instantly; Claude codes the rest in batches of 20.
3. Review the flagged rows: 🔴 no matching code, 🟡 needs review. In the **Decision** column choose Accepted, Excluded, or pick a different defect code. "Accept all yellow" and "Exclude all red" do it in one step.
4. **Save to history**. Excluded rows are left out.
5. **Monthly summary** tab: download the summary (codes in column A, one column per month, COUNTIF formulas) or the full history workbook.

To stop partway through, use **Download coded month**. Uploading that file later restores the codes and your decisions.

## Sharing it with others

To let coworkers use it, run it on a computer or server they can reach on your network:

```
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

They open `http://<that-computer's-name-or-IP>:8501`. Everyone shares the same `data` folder, so they see the same history and summary. Set `EDITOR_PASSWORD` in `secrets.toml` if only some people should be able to save months or change codes and rules; everyone else gets view-only access.

Back up the `data` folder regularly. It is the history.

## Before using it with real data

Auto-coding sends defect descriptions, along with past examples, to Anthropic's API. Check your company's rules on sending this kind of data to an outside service before turning auto-coding on. The rest of the app runs entirely on your own computer.

## Changing codes and rules

Use the **Codes & rules** tab, or edit `data/settings.json` directly. Rules are plain sentences; Claude follows them ahead of past examples.
