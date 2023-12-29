# Install

```bash
pip install skittle-sheets
```

To access Google Sheets (recommended), generate service file following
https://cloud.google.com/iam/docs/creating-managing-service-account-keys

Place it in the working directory, a parent directory, or your home directory. Make sure the filename matches `*service*.json`

# Example


# Usage (private Google Sheets)

Share your Google Sheets document with the service account generated above (read-only is fine). Then run:

```bash
skittle export "<spreadsheet>/<worksheet>"
```

If the Drive and Sheets APIs are not already enabled, there will be an error message with a link prompting you to enable them.

# Usage (public Google Sheets)

```bash
skittle export URL
```

# Usage (local)

```bash
skittle export test.csv
```
