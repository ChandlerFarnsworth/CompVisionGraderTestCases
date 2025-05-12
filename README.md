# Excel Worksheet Autograder

A simple autograder for Excel worksheets that evaluates both visible and hidden criteria.

## Files

This repository contains two sets of files - one for Coursera submission and one for local testing:

### For Coursera Submission

- `autograder.py` - Core grading script that integrates with Coursera's autograder
- `Dockerfile` - Container configuration for Coursera
- `solution.xlsx` - Reference solution file

### For Local Testing

- `grader.py` - Local version of the grading functionality
- `uploader.py` - Tool for grading individual submissions
- `batch.py` - Tool for batch grading multiple submissions

## How It Works

The autograder evaluates Excel worksheets based on two criteria:

1. **Y/N Values in Row 1** (80% of the score)
   - Compares cells in the first row (starting from column E)
   - Counts how many cells match between student submission and solution

2. **Hidden Test Cases** (20% of the score)
   - Checks specific cells (AD21, M62, AE187) without revealing them to students
   - Verifies that students have correctly completed those parts of the worksheet

The final score is calculated as:
```
Total Score = (Y/N Match % × 0.8) + (Hidden Test % × 0.2)
```

## Usage

### Local Testing

1. **Testing a Single Submission**:
   ```bash
   python uploader.py student_file.xlsx
   ```
   This will:
   - Copy the file to the `uploads` folder with a timestamp
   - Grade it against `solution.xlsx`
   - Show the score and basic feedback
   - Save detailed feedback to a text file

2. **Batch Testing**:
   ```bash
   # Grade all Excel files in a directory
   python batch.py submissions/
   
   # Grade specific files
   python batch.py file1.xlsx file2.xlsx
   
   # Grade all Excel files in current directory
   python batch.py
   ```
   This will:
   - Process multiple files at once
   - Generate summary reports in CSV and Excel formats in the `results` folder
   - Display statistics and a summary table

### Coursera Integration

To deploy this autograder on Coursera:

1. Create a ZIP file containing:
   - `autograder.py`
   - `Dockerfile`
   - `solution.xlsx`

2. Upload this ZIP file to Coursera's autograding system.

3. Set the assignment part ID in `autograder.py` to match your Coursera assignment:
   ```python
   COURSERA_PARTID = "Lg9eS"  # Update this with your assignment's part ID
   ```

## Requirements

- Python 3.6 or higher
- openpyxl (`pip install openpyxl`)
- pandas (`pip install pandas`) - for batch grading only

## Feedback Format

Students receive minimal feedback that includes their score and how many Y/N cells they matched correctly:

```
Your score: 82.50%
You correctly matched 64 out of 75 cells.
```

The feedback intentionally doesn't reveal the hidden tests to encourage students to complete the entire worksheet correctly.

## Troubleshooting

- If you see warnings about "Unknown extension" or "Conditional Formatting extension", these are just informational messages from openpyxl and don't affect grading.
- Ensure your solution file has the correct worksheets named "blank" and "solution".
- If batch grading requires too much memory, try processing files in smaller batches.