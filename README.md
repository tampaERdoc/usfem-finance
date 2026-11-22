# USF Emergency Medicine Revenue and Expenses

Password protected drill down of FY27 revenue and expenses by fund, program and person.

- `index.html` is the page. It decrypts `data.enc.json` in the browser after the password is entered.
- `data.enc.json` is the only data in this repo. It is AES-256-GCM encrypted. The password is not stored here.
- `build/build_data.py` rebuilds the data from the department Box folder and the iCloud ASF Monthly Pay Tracker.
- `build/encrypt.py` encrypts the rebuilt data.
- `build/revenue.csv` and `build/umsa_monthly.csv` hold the revenue lines and the UMSA 2600 monthly projection (no person level data).

## Weekly update

1. Stage the changed source files from Box and iCloud (paths are at the top of `build/build_data.py`).
2. `python3 -I build/build_data.py <uploads_root> build/revenue.csv build/umsa_monthly.csv data.json`
3. `FIN_PASSWORD=<password> python3 build/encrypt.py data.json data.enc.json`
4. Commit `data.enc.json` only. Never commit `data.json`.
