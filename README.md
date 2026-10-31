Python Transaction Automation
A Python-based data-processing project created to reduce repetitive manual work. The tool processes transaction history exported as text, applies the project's calculation rules, and organizes the results into an Excel workbook for review.

Project focus: Python · Data processing · Excel automation · AI-assisted development · Result verification
Why I built this
I originally needed to copy transaction information, calculate totals, and enter the results into a spreadsheet by hand. The process was repetitive and time-consuming. I wanted to see whether I could automate it instead of continuing to do everything manually.

I used ChatGPT as a programming and troubleshooting assistant, while taking responsibility for reading the generated code, modifying it as needed, checking the output, and verifying the totals. The finished workflow saved me several hours of manual work.
What it does
- Reads transaction records from text-based exports.
- Extracts and organizes relevant purchase and redemption information.
- Applies the calculation and discount-handling rules implemented in the script.
- Exports structured information to Excel for review.
- Replaces repetitive copying and arithmetic with a reusable workflow.
How I approached the problem
1. Identify the repetitive work: I identified the fields and calculations I repeatedly handled by hand.
2. Develop a solution with AI: I used ChatGPT to help plan and write an initial Python script.
3. Review and refine: I read the code, experimented with changes, and troubleshot issues as they appeared.
4. Verify accuracy: I checked the calculated totals independently instead of assuming the generated code was correct.
5. Export to Excel: I organized the results into a format that was easier to review and reuse.
Example data
This repository includes fully fictional example records, not real financial activity:
- sample_purchases.txt — fabricated purchase and reward entries in the source text layout.
- sample_redemptions.txt — fabricated redemption entries and statuses in the source text layout.
Dates, financial values, IDs, and banking descriptions in these files were generated for demonstration. Do not add real personal or financial exports to this public repository.
Running the project
1. Install Python 3 and the dependencies required by the main script.
2. Clone or download this repository.
3. Use the two fictional sample files as inputs. If the script expects particular filenames or locations, adjust those settings or copy the sample files to the expected names locally.
4. Run the main Python script and examine the exported Excel file.
5. Compare a few records and totals with the sample inputs to check the calculations.
Setup note: Update this section with the actual script filename, dependencies, and exact run command once the Python source file is included. These details are intentionally not guessed here.

AI use and what I learned
ChatGPT helped me develop the script and work through unfamiliar Python problems. However, I did not treat generated code as automatically correct. I reviewed and modified the code, tested the calculations, and independently verified results.

The most useful lesson was that AI can make it easier to try new technical approaches, but the person building the tool still needs to understand the problem, validate the output, and take responsibility for the result.
Current limitations
- The script was designed around a specific exported text layout rather than arbitrary transaction formats.
- The supplied public sample data is synthetic; it cannot validate every edge case in real exports.
- Input layout changes may require adjustments to the parser.
Potential improvements
- Add automated tests for purchase totals, redemption statuses, discounts, and malformed entries.
- Add clear command-line arguments and error messages for invalid files.
- Support additional input formats and optional summary reports.
Privacy
The original data contained personal financial history and is not included in the repository. All example inputs provided here are fictional. Generated output workbooks should be checked for sensitive data before publishing.
