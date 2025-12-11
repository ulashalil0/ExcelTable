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
* **Copy / CopyRow / CopyColumn:** Copies data from one location to another without deleting the source.
* **X / XRow / XColumn:** Cuts data (clears source) and pastes it to the destination.

### 4. Advanced Calculations (Polymorphic Operators)
The system uses special operators that behave differently based on the data types (String vs Integer):

* **Multiplication (*):**
  * `Int * Int`: Mathematical multiplication.
  * `String * Int`: Repeats the string `k` times. If `k` is negative, reverses the string then repeats.
* **Addition (+):**
  * `Int + Int`: Mathematical addition.
  * `String + String`: Concatenates strings (supports `low`/`up` case conversion).
* **Division (/):**
  * `Int / Int`: Integer division.
  * `String / Int`: Divides string length by `k`, returns the first segment (or last if negative).
* **Subtraction (-):**
  * `Int - Int`: Mathematical subtraction.
  * `String - Int`: Removes characters matching the ASCII value.
  * `String - String`: Removes occurrences of the shorter string from the longer one.
* **Encryption (#):**
  * Encrypts a string by shifting ASCII values by a given integer amount (Caesar Cipher logic).

### 5. File Operations
* The final state of the spreadsheet is automatically saved to `spreadsheet.txt` upon exit.

## 🛠️ Technical Details
* **Language:** C#
* **Data Structure:** Strictly uses **2D Arrays** for organization. `ArrayList` or `LinkedList` are NOT used as per assignment constraints.
* **Error Handling:** Robust checks for "Index Out of Bounds", "Illegal Operation", and Type Mismatches.

## 💻 Installation & Usage

1.  Download or clone the project repository.
2.  Navigate to the project directory and run:
    ```
    dotnet run
    ```
3.  Sample Command Usage:
    ```
    >> AssignValue(A1, integer, 10)
    >> AssignValue(B1, string, Apple)
    >> *(B1, A1, C1)
    ```

## 👨‍💻 Developer
* **Name:** Ulaş Halil Çetintaş
* **Department:** Computer Engineering
* **University:** Dokuz Eylül University