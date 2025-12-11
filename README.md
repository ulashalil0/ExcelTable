\# 📊 Console Based Mini Excel System



\## 🚀 About The Project

This project is a console-based mini Excel simulation developed using C# programming language. It was created as part of the \*\*Algorithm and Programming\*\* course at \*\*Dokuz Eylül University, Computer Engineering Department\*\*.



The application simulates a simplified spreadsheet using \*\*2-dimensional arrays\*\*, allowing users to perform data entry, row/column management, and various polymorphic operations.



\## ✨ Features \& Operations



\[cite\_start]The system supports a grid of maximum \*\*15 rows (1-15)\*\* and \*\*10 columns (A-J)\*\*\[cite: 11].



\### 1. Basic Cell Operations

\* \*\*AssignValue:\*\* Assigns `string` or `integer` values to cells. \[cite\_start]Strings longer than 5 characters are truncated with `\_` for display\[cite: 13, 21].

\* \[cite\_start]\*\*ClearCell / ClearAll:\*\* Clears content of a specific cell or the entire spreadsheet\[cite: 28, 33].

\* \[cite\_start]\*\*Display:\*\* Users can view full cell content by typing the cell coordinate\[cite: 22].



\### 2. Table Management

\* \[cite\_start]\*\*AddRow:\*\* Inserts a new row (shifts existing rows `up` or `down`)\[cite: 38].

\* \[cite\_start]\*\*AddColumn:\*\* Inserts a new column (shifts existing columns `left` or `right`)\[cite: 48].

\* \[cite\_start]\*\*Constraints:\*\* Handles boundary checks and strictly enforces the maximum size (15x10)\[cite: 44, 54].



\### 3. Clipboard Operations

\* \[cite\_start]\*\*Copy / CopyRow / CopyColumn:\*\* Copies data from one location to another without deleting the source\[cite: 58, 64, 72].

\* \[cite\_start]\*\*X / XRow / XColumn:\*\* Cuts data (clears source) and pastes it to the destination\[cite: 80, 86, 94].



\### 4. Advanced Calculations (Polymorphic Operators)

The system uses special operators that behave differently based on the data types (String vs Integer):



\* \*\*Multiplication (`\*`):\*\* \* `Int \* Int`: Mathematical multiplication.

&nbsp; \* `String \* Int`: Repeats the string `k` times. \[cite\_start]If `k` is negative, reverses the string then repeats\[cite: 102, 107].

\* \*\*Addition (`+`):\*\* \* `Int + Int`: Mathematical addition.

&nbsp; \* \[cite\_start]`String + String`: Concatenates strings (supports `low`/`up` case conversion)\[cite: 114, 119].

\* \*\*Division (`/`):\*\* \* `Int / Int`: Integer division.

&nbsp; \* `String / Int`: Divides string length by `k`, returns the first segment (or last if negative)\[cite: 124, 126].

\* \*\*Subtraction (`-`):\*\* \* `Int - Int`: Mathematical subtraction.

&nbsp; \* `String - Int`: Removes characters matching the ASCII value.

&nbsp; \* `String - String`: Removes occurrences of the shorter string from the longer one\[cite: 137, 140, 145].

\* \[cite\_start]\*\*Encryption (`#`):\*\* \* Encrypts a string by shifting ASCII values by a given integer amount (Caesar Cipher logic)\[cite: 151, 156].



\### 5. File Operations

\* \[cite\_start]The final state of the spreadsheet is automatically saved to `spreadsheet.txt` upon exit\[cite: 169].



\## 🛠️ Technical Details

\* \*\*Language:\*\* C#

\* \*\*Data Structure:\*\* Strictly uses \*\*2D Arrays\*\* for organization. \[cite\_start]`ArrayList` or `LinkedList` are NOT used as per assignment constraints\[cite: 219, 220].

\* \[cite\_start]\*\*Error Handling:\*\* robust checks for "Index Out of Bounds", "Illegal Operation", and Type Mismatches\[cite: 222].



\## 💻 Installation \& Usage



1\.  Clone the repository:

&nbsp;   ```bash

&nbsp;   git clone \[https://github.com/YOUR\_USERNAME/YOUR\_REPO\_NAME.git](https://github.com/YOUR\_USERNAME/YOUR\_REPO\_NAME.git)

&nbsp;   ```

2\.  Navigate to the project directory and run:

&nbsp;   ```bash

&nbsp;   dotnet run

&nbsp;   ```

3\.  Sample Command Usage:

&nbsp;   ```text

&nbsp;   >> AssignValue(A1, integer, 10)

&nbsp;   >> AssignValue(B1, string, Apple)

&nbsp;   >> \*(B1, A1, C1)

&nbsp;   ```



\## 👨‍💻 Developer

\* \*\*Name:\*\* Ulaş Halil Çetintaş

\* \*\*Department:\*\* Computer Engineering

\* \*\*University:\*\* Dokuz Eylül University

