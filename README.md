# 🏆 Leaderboard Points Ranking Assessment

## Project Overview

This project provides a robust, scripted solution for accurately calculating and ranking a competitive leaderboard based on a complex, multi-level set of criteria. The solution handles non-numeric data, calculates total points, and implements a strict sequence of tiebreakers, including a recursive countback system and conditional alphabetical sorting with custom highlighting.

The primary goal was to take raw event data from an Excel spreadsheet and produce a finalized, ranked leaderboard that adheres precisely to all specified assessment rules.

## Problem Statement & Ranking Criteria

The leaderboard ranking is determined by the following criteria, applied strictly in order:

1.  **Primary Ranking:** **Total Points Scored** (Descending).
    * Scores marked as **"D$Q"** or **"-"** are treated as **zero (0)** points.

2.  **Tiebreaker 1:** **Total Spending** (Ascending).
    * Lower total spending ranks higher.

3.  **Tiebreaker 2 (Countback System):** A recursive system comparing the best scores and their frequencies.
    * **Sub-Criterion A:** Compare the **Highest Absolute Score**.
    * **Sub-Criterion B:** If scores are equal, compare the **Most Occurrences** of that score.
    * This process repeats sequentially for the 2nd highest score, 3rd highest score, and so on, until the tie is broken.

4.  **Final Tiebreaker:** **Player Name** (Alphabetical Order).
    * If all previous criteria fail to break the tie, the players are sorted alphabetically.
    * Players whose final rank was determined solely by this alphabetical tiebreaker are highlighted in **red**.

## Solution Methodology

The ranking logic is implemented using Python and the Pandas library for efficient data manipulation.

* **Data Cleaning:** Used a robust method to clean non-numeric values (e.g., `$m` from the spending column, `D$Q`, and `-` from scores) to ensure all data for calculation is in `float` format.
* **Total Calculation:** Standard summation was used for Total Points.
* **Countback Key:** The complex recursive countback system was consolidated into a single **`Countback_Key`** tuple. This tuple encodes the full sequence of (Score 1, Count 1, Score 2, Count 2, ...) for each player, allowing a single `sort_values` operation to correctly implement the entire recursive logic.
* **Final Sort:** A single, multi-column `sort_values` operation was executed using the priority order: `Total_Points` (Desc), `Spent` (Asc), `Countback_Key` (Desc), and `Player` (Asc).
* **Formatting:** The `openpyxl` engine was used to write the final DataFrame to an Excel file and apply the custom **red font highlighting** to the names of players who remained tied after all numerical criteria.

## Files and Dependencies

### Input
* `leaderboard.xlsx`: The raw data file containing player names, event scores (R01 to R22), and spending information.

### Output
* `final_leaderboard.xlsx`: The generated Excel file containing the complete data with the new **Rank** column. Names of players sorted only by the alphabetical rule are highlighted in **red**.

### Dependencies
* Python 3.x
* `pandas`
* `numpy`
* `openpyxl` (Required for Excel output and formatting)

To install dependencies:
```bash
pip install pandas numpy openpyxl

How to Run
Ensure Python and the required libraries (pandas, numpy, openpyxl) are installed.

Place the input file (leaderboard.xlsx) in the same directory as the script (leader.py).

Execute the script from your terminal:

Bash

python leader.py
The final ranked leaderboard will be saved as final_leaderboard.xlsx in the same directory.
