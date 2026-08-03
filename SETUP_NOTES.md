# What changed

1. **Density now carries forward** on new rows, exactly like Species Culture
   and Cycle Type already did — a new row starts pre-filled with the
   previous record's Density instead of 0.
2. **"Save Changes to Pond History" button is gone.** Each record is now its
   own card with its own **red "💾 Save"** button.
3. **Status column**: once you click Save on a card and it succeeds, that
   button is replaced by a green **"✅ Saved"** pill. If you edit any field
   on a saved card afterwards, it flips back to a Save button so you can
   re-save it.
4. **Issues is now a true multi-select** — pick as many issue tags as apply
   to that visit (stored in the sheet joined by `; `).
5. **Data now lives in Google Sheets**, not a local CSV. See setup below.
6. **Mobile layout**: only the history tables (`st.dataframe`) scroll
   horizontally now (that's native Streamlit behavior for wide tables).
   Every other block of fields — the customer/farm pickers and each pond
   record card — stacks into a single column, one field after another, on
   screens narrower than 700px.

# Connecting Google Sheets

1. In Google Cloud Console, create/select a project, then enable the
   **Google Sheets API** and **Google Drive API**.
2. Create a **Service Account**, then create and download a **JSON key**
   for it.
3. Create a Google Sheet for this app's data, and **Share** it with the
   service account's `client_email` (found in the JSON key) with
   **Editor** access.
4. Copy the Sheet ID out of its URL:
   `https://docs.google.com/spreadsheets/d/`**`THIS_PART`**`/edit`
5. Add a `.streamlit/secrets.toml` file (locally) or paste into the
   "Secrets" section of Streamlit Community Cloud:

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "...@....iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."

[gsheet]
sheet_id = "the-id-you-copied-from-the-url"
worksheet_name = "WaterQualityData"
```

Everything in the JSON key file maps directly to the fields above (just
paste each value in as a string). The app creates the `WaterQualityData`
worksheet and its header row automatically the first time it runs, if it
doesn't exist yet.

6. Keep `Customer List.xlsx` in the same folder as `app.py`, same as
   before — that part hasn't changed.

# Install & run

```bash
pip install -r requirements.txt
streamlit run app.py
```
