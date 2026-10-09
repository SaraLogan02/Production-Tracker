[README.md](https://github.com/user-attachments/files/33259433/README.md)

# Excel Production Tracker

An Excel tool I built to track claims workload, calculate production time from activity counts, and summarize meeting and downtime entries.

I created the original tracker in May 2021 to help monitor daily production, and it was adopted by my coworkers. This repository presents a demonstration version with sample entries.

## Purpose

Different claim types have different production rates. This tracker counts entries by type and status, converts selected counts into credited production time, and combines that time with meeting and other downtime entries. It provides a running summary without manually recounting each activity.

## Features

- Dropdown lists for consistent claim types and statuses.
- Automatic counts using `COUNTIF`, `SUM`, and grouped criteria.
- Production-time calculations based on the rate assigned to each claim category.
- Linked summaries across the DF Tracker and Downtime Tracker worksheets.
- Start and end time inputs for meeting and other downtime durations.
- Conditional formatting to distinguish claim types and statuses.
- Duplicate DCN highlighting to flag repeated identifiers for review.
- A VBA Clear button to reset claim inputs and restore time inputs to midnight while preserving the calculation cells.

## Files

| File | Purpose |
| --- | --- |
| [Production_Tracker_Demo.xlsm](Production_Tracker_Demo.xlsm) | Filled example showing counts, status colors, and time calculations. |
| [Production_Tracker_Blank.xlsm](Production_Tracker_Blank.xlsm) | Empty version for trying the workflow. |

## Screenshots

### Production tracking

![Production tracker with sample entries and calculated summaries](production-tracker.png)

### Time tracking

![Meeting and other downtime entries with duration totals](downtime-tracker.png)

## How to use

1. Download a workbook and open it in desktop Microsoft Excel.
2. Enter a sample DCN, select a claim type and status, and add an optional note on DF Tracker, starting in row 6.
3. Review the automatic counts and production-time summary. The entry and counting area runs through row 500.
4. On Downtime Tracker, enter start and end times for meetings or other downtime. Enter times such as `7:00 AM` or `1:30 PM`, rather than a whole number such as `7`.
5. Review the linked time totals.
6. Save a copy of any entries you want to retain before using Clear. The reset macro clears DF Tracker inputs in `A6:D500` and resets Downtime Tracker time inputs in `E2:F10` and `K2:L7` to zero, displayed as midnight.

The workbooks use the `.xlsm` format to retain VBA. The Clear button requires macros to be enabled in desktop Excel. Enable macros only after reviewing and trusting the file; Excel for the web does not run VBA macros.

## Calculation notes

- Credited production hours equal activity count divided by the applicable hourly production rate. Dividing by 24 stores that result as an Excel time value.
- Credited production time is rate-based, rather than a measurement of actual elapsed processing time.
- Meeting and other downtime durations equal end time minus start time. Enter same-day intervals with the end time at or after the start time.
- Valid DF/Other time is displayed separately and intentionally excluded from the combined total because it is entered separately in the production-reporting workflow.
- The status total includes Valid DF/Other, Returned CIW, and Created DF. Hold/Work Later is shown separately, so the status total may differ from the claim-type total.
- Duplicate highlighting flags repeated IDs for review. It does not remove entries or exclude them from counts.
- Internal labels such as DF, CIW, and DCN come from the original workflow; DCN serves as the record identifier in this demonstration.

## Skills demonstrated

Microsoft Excel formulas, data validation, conditional formatting, time calculations, cross-sheet references, VBA automation, and translating a recurring operational need into a reusable tool.

