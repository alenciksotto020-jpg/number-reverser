# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a single-file React component (`number-reverser.jsx`) that processes Excel spreadsheet data containing measurement patterns. It reverses number sequences based on specific marker patterns used in construction/manufacturing contexts.

## Architecture

The application is a standalone React component with no build system or package manager. It uses:
- React with hooks (useState, useCallback)
- SheetJS (xlsx) library for Excel file parsing and export
- Inline CSS-in-JS styling

### Core Processing Logic

The component processes rows looking for specific patterns:
- **H1/V1 markers**: Rows starting with `h 1` or `v 1` (horizontal/vertical direction markers)
- **W2/S2 markers**: Output markers for width/spacing that follow H1/V1 rows
- **Combined patterns**: Format like `53.0+v1` which triggers row splitting
- **Imposts**: Section markers that get preserved in output

Processing rules:
1. Only odd-count number sequences are processed (even counts are skipped)
2. Combined patterns are split into two separate row pairs
3. Each H1/V1 row generates a corresponding W2/S2 row with reversed numbers
4. The marker type alternates (h→v, v→h) when splitting combined patterns

### Key Functions

- `process()`: Main processing loop that iterates through input rows
- `getNums()`: Extracts number sequences from a row starting after a marker
- `buildRow()`: Constructs output rows with proper marker placement
- `parse()`: Entry point for both file and pasted data

### Data Flow

```
Input (Excel/CSV/Paste) → parse() → process() → data state → Table/Export
```

## Usage

This component is designed to be imported and rendered in a React application. It handles:
- File upload (drag & drop or file picker for .xlsx, .xls, .csv)
- Tab-separated text paste
- Excel export with `reversed_` prefix
- Clipboard copy of processed data
