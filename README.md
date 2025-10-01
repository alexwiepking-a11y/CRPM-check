# CRPM Compliance Checker

This tool analyzes hotel rate plan compliance for VAT, subaccount, and city tax, helping you quickly identify issues and track accepted exceptions.

## Features

- Loads rate plan and standard data from Excel files
- Checks for deviations in VAT, subaccount, and city tax
- Compares deviations against accepted exceptions
- Suggests new exception rules based on patterns
- Generates prioritized Excel reports for issues and accepted deviations
- Creates an HTML dashboard for quick insights
- Uses command-line arguments for flexible file paths
- Logs all actions for easy troubleshooting

## How to Use

1. Place your Excel files in the `source` folder (or specify your own paths).
2. Run the script from the terminal:

   ```
   python check_crpm.py --input source/source_CRPM_check.xlsx --exceptions source/source_CRPM_exceptions.xlsx --output output
   ```

   - `--input`: Path to your rate plan Excel file
   - `--exceptions`: Path to your exceptions Excel file
   - `--output`: Output folder for results

   If you omit these arguments, the script uses the default paths.

3. Find results in the output folder, including:
   - Prioritized issues to fix
   - Accepted deviations
   - Suggested exception rules
   - Executive summary
   - HTML dashboard

## Configuration

- Edit the list of city tax hotels in `check_crpm.py` (`CITY_TAX_HOTELS`).
- To disable dashboard creation, set `CREATE_DASHBOARD = False` at the top of the script.
- To skip auto-opening the dashboard, set `AUTO_OPEN_DASHBOARD = False`.

## Troubleshooting

- If you see errors about missing columns or files, check your Excel file names and sheet names.
- For more details, set logging to `DEBUG` in the script.
- All errors and progress are logged in the terminal.

## Project Structure

- `check_crpm.py` — main script
- `exceptions.py` — exception logic
- `dashboard.py` — dashboard creation
- `test_exceptions.py` — unit tests
- `README.md` — this file

---

**Contact:**  
For questions or help, contact your project administrator.