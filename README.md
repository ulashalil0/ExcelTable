# 📊 Console Based Mini Excel System

## 🚀 About The Project
This project is a console-based mini Excel simulation developed using C# programming language. It was created as part of the **Algorithm and Programming** course at **Dokuz Eylül University, Computer Engineering Department**.

The application simulates a simplified spreadsheet using **2-dimensional arrays**, allowing users to perform data entry, row/column management, and various polymorphic operations.

## ✨ Features & Operations

The system supports a grid of maximum **15 rows (1-15)** and **10 columns (A-J)**.

### 1. Basic Cell Operations
* **AssignValue:** Assigns `string` or `integer` values to cells. Strings longer than 5 characters are truncated with `_` for display.
* **ClearCell / ClearAll:** Clears content of a specific cell or the entire spreadsheet.
* **Display:** Users can view full cell content by typing the cell coordinate.

### 2. Table Management
* **AddRow:** Inserts a new row (shifts existing rows `up` or `down`).
* **AddColumn:** Inserts a new column (shifts existing columns `left` or `right`).
* **Constraints:** Handles boundary checks and strictly enforces the maximum size (15x10).

### 3. Clipboard Operations
* **Copy / CopyRow / CopyColumn:** Copies data from one location to another without deleting