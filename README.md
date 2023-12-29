# Install

```bash
pip install skittle-sheets
```

To access Google Sheets (recommended), generate service file following
https://cloud.google.com/iam/docs/creating-managing-service-account-keys

Place it in the working directory, a parent directory, or your home directory. Make sure the filename matches `*service*.json`

# Usage (Google Sheets)

Share your Google Sheets document with the service account generated above (read-only is fine). Then run:

```bash
skittle export "drive:<spreadsheet>/<worksheet>
``` 

# Usage (local)

```bash
skittle export test.csv
skittle export test.xlsx
```
