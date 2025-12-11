using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelBoard
{
    internal class Program
    {
        static int currentRow = 8;
        static int currentCol = 5;
        static int[,] TableTypes = new int[15, 10]; // String --> 1 , Integer --> -1 , Empty --> 0 
        static string[,] Table = new string[15, 10]; // To save everything as a string
        static string[] letters = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j" };
        static int[] letterCordinates = { 5, 14, 23, 32, 41, 50, 59, 68, 77, 86 };
        static int[] numberCordinates = { 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };
        static void Main(string[] args)
        {
            LoadFromTxt();
            while (true)
            {
                Console.Clear();
                DrawTable();
                Console.Write(" >> ");
                string input = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(input))
                {
                    continue;
                }
                input = input.Trim();

                if (input.ToLower() == "exit")
                {
                    SaveToTxt(); 
                    Console.WriteLine("Exiting program...");
                    break; 
                }

                int openIndex = input.IndexOf('(');
                int closeIndex = input.LastIndexOf(')');

                if (openIndex == -1 || closeIndex == -1 || closeIndex < openIndex)
                {
                    if ((input.Length == 2 || input.Length == 3) && IsLetter(input.Substring(0, 1)) && IsInteger(input.Substring(1)))
                    {
                        if (CheckRange(Convert.ToInt32(input.Substring(1)), input.Substring(0, 1)))
                        {
                            ShowCell(input);
                            continue;
                        }
                        else
                        {
                            ErrorMessage("Out of the range!");
                            continue;
                        }
                    }
                    else
                    {
                        ErrorMessage("Invalid syntax!");
                        continue;
                    }
                }
                string command = input.Substring(0, openIndex).Trim().ToLower();
                string content = input.Substring(openIndex + 1, closeIndex - openIndex - 1).Trim();
                string[] parameters;
                if (string.IsNullOrWhiteSpace(content))
                {
                    parameters = new string[0]; 
                }
                else
                {
                    parameters = content.Split(',');
                    for (int i = 0; i < parameters.Length; i++)
                    {
                        parameters[i] = parameters[i].Trim();
                    }
                }
                for (int i = 0;  i < parameters.Length; i++)
                {
                    parameters[i] = parameters[i].Trim();
                }
                if (command == "assignvalue")
                {
                    if (parameters.Length != 3)
                    {
                        ErrorMessage("AssignValue requires 3 parameters!");
                        continue;
                    }
                    AssignValue(parameters[0], parameters[1], parameters[2]);
                }
                else if (command == "clearcell")
                {
                    if (parameters.Length != 1)
                    {
                        ErrorMessage("Invalid command expression!");
                        continue;
                    }
                    ClearCell(parameters[0]);
                }
                else if (command == "clearall")
                {
                    if (parameters.Length != 0)
                    {
                        ErrorMessage("ClearAll does not accept parameters!");
                        continue;
                    }
                    ClearAll();
                }
                else if (command == "addrow")
                {
                    if (parameters.Length != 2)
                    {
                        ErrorMessage("AddRow requires 2 parameters!");
                        continue;
                    }
                    AddRow(parameters);
                }
                else if (command == "addcolumn")
                {
                    if (parameters.Length != 2)
                    {
                        ErrorMessage("AddColumn requires 2 parameters!");
                        continue;
                    }
                    AddColumn(parameters);
                }
                else if (command == "copy")
                {
                    if (parameters.Length != 2)
                    {
                        ErrorMessage("Copy requires 2 parameters!");
                        continue;
                    }
                    Copy(parameters[0], parameters[1]);
                }
                else if (command == "copycolumn")
                {
                    if (parameters.Length != 2)
                    {
                        ErrorMessage("CopyColumn requires 2 parameters!");
                        continue;
                    }
                    CopyColumn(parameters);
                }
                else if (command == "copyrow")
                {
                    if (parameters.Length != 2)
                    {
                        ErrorMessage("CopyRow requires 2 parameters!");
                        continue;
                    }
                    CopyRow(parameters);
                }
                else if (command == "x")
                {
                    if (parameters.Length != 2)
                    {
                        ErrorMessage("X requires 2 parameters!");
                        continue;
                    }
                    X(parameters[0], parameters[1]);
                }
                else if (command == "xcolumn")
                {
                    if (parameters.Length != 2)
                    {
                        ErrorMessage("XColumn requires 2 parameters!");
                        continue;
                    }
                    XColumn(parameters);
                }
                else if (command == "xrow")
                {
                    if (parameters.Length != 2)
                    {
                        ErrorMessage("XRow requires 2 parameters!");
                        continue;
                    }
                    XRow(parameters);
                }
                else if (command == "*")
                {
                    if (parameters.Length != 3)
                    {
                        ErrorMessage("* operator requires 3 parameters!");
                        continue;
                    }
                    Multiply(parameters);
                }
                else if (command == "+")
                {
                    if (parameters.Length < 3 || parameters.Length > 4)
                    {
                        ErrorMessage("+ operator requires 3 or 4 parameters!");
                        continue;
                    }
                    SumOrJoin(parameters);
                }
                else if (command == "/")
                {
                    if (parameters.Length != 3)
                    {
                        ErrorMessage("/ operator requires 3 parameters!");
                        continue;
                    }
                    DivideOrSlice(parameters);
                }
                else if (command == "-")
                {
                    if (parameters.Length != 3)
                    {
                        ErrorMessage("- operator requires 3 parameters!");
                        continue;
                    }
                    SubtractOrRemove(parameters);
                }
                else if (command == "#")
                {
                    if (parameters.Length != 3)
                    {
                        ErrorMessage("# operator requires 3 parameters!");
                        continue;
                    }
                    ShiftCharacters(parameters);
                }
                else
                {
                    ErrorMessage("There is no such command.");
                    continue;
                }
            }  
        }
        static void ShowCell(string cordinate)
        {
            int row = GetNumber(cordinate);
            int col = GetLetter(cordinate);
            if (TableTypes[row, col] == 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\nCell is empty!");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("\nValue: " + Table[row, col]);
                string typeName = "";
                if (TableTypes[row, col] == -1)
                {
                    typeName = "Integer";
                }
                else if (TableTypes[row, col] == 1)
                {
                    typeName = "String";
                }
                Console.WriteLine("Type : " + typeName);
                Console.ResetColor();
            }
            Console.WriteLine("\nPress any key to continue");
            Console.ReadKey();
        }
        static bool IsLetter(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            for (int i = 0; i < text.Length; i++)
            {
                int ascii = (int)text[i];
                bool isUpper = (ascii >= 65 && ascii <= 90);
                bool isLower = (ascii >= 97 && ascii <= 122);

                if (!isUpper && !isLower)
                {
                    return false;
                }
            }
            return true;
        }
        static bool IsInteger(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            { 
                return false;
            }
            for (int i = 0; i < text.Length; i++)
            {
                int asciiCode = (int)text[i];
                if (i == 0 && asciiCode == 45)
                {
                    if (text.Length == 1) return false;
                    continue;
                }
                if (asciiCode < 48 || asciiCode > 57)
                {
                    return false;
                }
            }
            return true;
        }
        static void ShiftCharacters(string[] parameters)
        {
            for (int i = 0; i < parameters.Length; i++)
            {
                if (GetNumber(parameters[i]) == -1 || GetLetter(parameters[i]) == -1)
                {
                    ErrorMessage("Invalid coordinate format found in parameters!");
                    return;
                }

                if (i < parameters.Length - 1)
                {
                    if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 0)
                    {
                        ErrorMessage("One of the source cells is not assigned!");
                        return;
                    }
                }
            }
            int stringCounter = 0;
            int whichIsString = 0;
            string result = "";
            for (int i = 0; i < 2; i++)
            {
                if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 0)
                {
                    ErrorMessage("Cell is not assigned!");
                    return;
                }
                if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 1)
                {
                    stringCounter++;
                    whichIsString += i;
                }
            }
            if (stringCounter == 1)
            {
                int number = Convert.ToInt32(Table[GetNumber(parameters[1 - whichIsString]), GetLetter(parameters[1 - whichIsString])]);
                if (number < -20 || number > 30)
                {
                    ErrorMessage("Value out of range exception!");
                    return;
                }
                for (int i = 0; i < Table[GetNumber(parameters[whichIsString]), GetLetter(parameters[whichIsString])].Length; i++)
                {
                    char letter = Table[GetNumber(parameters[whichIsString]), GetLetter(parameters[whichIsString])][i];
                    int letterASCII = (int)letter + Convert.ToInt32(Table[GetNumber(parameters[1 - whichIsString]), GetLetter(parameters[1 - whichIsString])]);
                    char newLetter = (char)letterASCII;
                    result += newLetter;
                }
            }
            else
            {
                ErrorMessage("Illegal operation!");
                return;
            }
        }
        static void SubtractOrRemove(string[] parameters)
        {
            for (int i = 0; i < parameters.Length; i++)
            {
                if (GetNumber(parameters[i]) == -1 || GetLetter(parameters[i]) == -1)
                {
                    ErrorMessage("Invalid coordinate format found in parameters!");
                    return;
                }

                if (i < parameters.Length - 1)
                {
                    if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 0)
                    {
                        ErrorMessage("One of the source cells is not assigned!");
                        return;
                    }
                }
            }
            int stringCounter = 0;
            int whichIsString = 0;
            int letterASCII;
            for (int i = 0; i < 2; i++)
            {
                if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 0)
                {
                    ErrorMessage("Cell is not assigned!");
                    return;
                }
                if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 1) 
                {
                    stringCounter++;
                    whichIsString += i;
                }
            }
            if (stringCounter == 0)
            {
                int number1 = Convert.ToInt32(Table[GetNumber(parameters[0]), GetLetter(parameters[0])]);
                int number2 = Convert.ToInt32(Table[GetNumber(parameters[1]), GetLetter(parameters[1])]);
                int result = number1 - number2;
                Table[GetNumber(parameters[2]), GetLetter(parameters[2])] = Convert.ToString(result);
                TableTypes[GetNumber(parameters[2]), GetLetter(parameters[2])] = -1;
            }
            else if (stringCounter == 1)
            {
                letterASCII = Convert.ToInt32(Table[GetNumber(parameters[1 - whichIsString]), GetLetter(parameters[1 - whichIsString])]);
                if(letterASCII < 65 || letterASCII > 126)
                {
                    ErrorMessage("Value out of range exception!");
                    return;
                }
                string String = Table[GetNumber(parameters[whichIsString]), GetLetter(parameters[whichIsString])];
                char letter = (char)letterASCII;
                string result = "";
                for(int i = 0; i < String.Length; i++)
                {
                    if(String[i] != letter)
                    {
                        result += String[i];
                    }
                }
                Table[GetNumber(parameters[2]), GetLetter(parameters[2])] = result;
                TableTypes[GetNumber(parameters[2]), GetLetter(parameters[2])] = 1;
            }
            else
            {
                string longString = "";
                string shortString = "";
                string result;
                
                if (Table[GetNumber(parameters[0]), GetLetter(parameters[0])].Length > Table[GetNumber(parameters[1]), GetLetter(parameters[1])].Length)
                {
                    longString = Table[GetNumber(parameters[0]), GetLetter(parameters[0])];
                    shortString = Table[GetNumber(parameters[1]), GetLetter(parameters[1])];
                    result = longString.Replace(shortString, "");
                }
                else if (Table[GetNumber(parameters[0]), GetLetter(parameters[0])].Length < Table[GetNumber(parameters[1]), GetLetter(parameters[1])].Length)
                {
                    longString = Table[GetNumber(parameters[1]), GetLetter(parameters[1])];
                    shortString = Table[GetNumber(parameters[0]), GetLetter(parameters[0])];
                    result = longString.Replace(shortString, "");
                }
                else
                {
                    if (Table[GetNumber(parameters[0]), GetLetter(parameters[0])] == Table[GetNumber(parameters[1]), GetLetter(parameters[1])])
                    {
                        result = "";
                    }
                    else
                    {
                        result = Table[GetNumber(parameters[0]), GetLetter(parameters[0])];
                    }
                }
                Table[GetNumber(parameters[2]), GetLetter(parameters[2])] = Convert.ToString(result);
                TableTypes[GetNumber(parameters[2]), GetLetter(parameters[2])] = 1;
            }
        }
        static void DivideOrSlice(string[] parameters)
        {
            for (int i = 0; i < parameters.Length; i++)
            {
                if (GetNumber(parameters[i]) == -1 || GetLetter(parameters[i]) == -1)
                {
                    ErrorMessage("Invalid coordinate format found in parameters!");
                    return;
                }

                if (i < parameters.Length - 1)
                {
                    if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 0)
                    {
                        ErrorMessage("One of the source cells is not assigned!");
                        return;
                    }
                }
            }
            string cordinate1 = parameters[0];
            string cordinate2 = parameters[1];
            string cordinate3 = parameters[2];
            
            if (TableTypes[GetNumber(cordinate1), GetLetter(cordinate1)] == -1 && TableTypes[GetNumber(cordinate2), GetLetter(cordinate2)] == -1)
            {
                int number1 = Convert.ToInt32(Table[GetNumber(cordinate1), GetLetter(cordinate1)]);
                int number2 = Convert.ToInt32(Table[GetNumber(cordinate2), GetLetter(cordinate2)]);
                if (number2 == 0)
                {
                    ErrorMessage("Divide by zero exception!");
                    return;
                }
                int result = number1 / number2;
                AssignValue(letters[GetLetter(cordinate3)] + (GetNumber(cordinate3) + 1), "integer", Convert.ToString(result));
            }
            else if (TableTypes[GetNumber(cordinate1) , GetLetter(cordinate1)] == 1 && TableTypes[GetNumber(cordinate2), GetLetter(cordinate2)] == -1)
            {
                int stringLength = Table[GetNumber(cordinate1), GetLetter(cordinate1)].Length / Math.Abs(Convert.ToInt32(Table[GetNumber(cordinate2), GetLetter(cordinate2)]));
                if (Convert.ToInt32(Table[GetNumber(cordinate2), GetLetter(cordinate2)]) > 0)
                {
                    string result = "";
                    for (int i = 0; i < stringLength; i++)
                    {
                        result += Table[GetNumber(cordinate1), GetLetter(cordinate1)][i];
                    }
                    AssignValue(letters[GetLetter(cordinate3)] + (GetNumber(cordinate3) + 1), "string", result);
                }
                else
                {
                    string result = "";
                    for (int i = Table[GetNumber(cordinate1), GetLetter(cordinate1)].Length - 1; i >= Table[GetNumber(cordinate1), GetLetter(cordinate1)].Length - stringLength; i--)
                    {
                        result += Table[GetNumber(cordinate1), GetLetter(cordinate1)][i];
                    }
                    AssignValue(letters[GetLetter(cordinate3)] + (GetNumber(cordinate3) + 1), "string", result);
                }
            }
            else if (TableTypes[GetNumber(cordinate1) , GetLetter(cordinate1)] == -1 && TableTypes[GetNumber(cordinate2), GetLetter(cordinate2)] == 1)
            {
                int stringLength = Table[GetNumber(cordinate2), GetLetter(cordinate2)].Length / Math.Abs(Convert.ToInt32(Table[GetNumber(cordinate1), GetLetter(cordinate1)]));
                if (Convert.ToInt32(Table[GetNumber(cordinate1), GetLetter(cordinate1)]) > 0)
                {
                    string result = "";
                    for (int i = 0; i < stringLength; i++)
                    {
                        result += Table[GetNumber(cordinate2), GetLetter(cordinate2)][i];
                    }
                    AssignValue(letters[GetLetter(cordinate3)] + (GetNumber(cordinate3) + 1), "string", result);
                }
                else
                {
                    string result = "";
                    for (int i = Table[GetNumber(cordinate2), GetLetter(cordinate2)].Length - 1; i >= Table[GetNumber(cordinate2), GetLetter(cordinate2)].Length - stringLength; i--)
                    {
                        result += Table[GetNumber(cordinate2), GetLetter(cordinate2)][i];
                    }
                    AssignValue(letters[GetLetter(cordinate3)] + (GetNumber(cordinate3) + 1), "string", result);
                }
            }
            else
            {
                ErrorMessage("You cannot use this command for two strings.");
                return;
            }
        }
        static int GetNumber(string cordinate)
        {
            if (string.IsNullOrWhiteSpace(cordinate) || cordinate.Trim().Length < 2)
            {
                return -1;
            }
            string numberPart = cordinate.Trim().Substring(1);
            if (!IsInteger(numberPart))
            {
                return -1; 
            }
            int numberIndex = Convert.ToInt32(numberPart) - 1;
            return numberIndex;
        }
        static int GetLetter(string cordinate)
        {
            string firstChar = cordinate.Trim().Substring(0, 1).ToLower();
            int letterIndex = Array.IndexOf(letters ,firstChar);
            return letterIndex;
        }
        static void SumOrJoin(string[] parameters)
        {
            for (int i = 0; i < parameters.Length; i++)
            {
                if (GetNumber(parameters[i]) == -1 || GetLetter(parameters[i]) == -1)
                {
                    ErrorMessage("Invalid coordinate format found in parameters!");
                    return;
                }
                if (i < parameters.Length - 1)
                {
                    if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 0)
                    {
                        ErrorMessage("One of the cells is not assigned!");
                        return;
                    }
                }
            }
            if (parameters.Length != 3 && parameters.Length != 4)
            {
                ErrorMessage("Invalid command expression!");
                return;
            }
            bool isString = false;
            for (int i = 0; i < parameters.Length - 1; i++)
            {
                if (TableTypes[GetNumber(parameters[i]), GetLetter(parameters[i])] == 1)
                {
                    isString = true;
                    break;
                }
            }
            int targetRow = GetNumber(parameters[parameters.Length - 1]);
            int targetCol = GetLetter(parameters[parameters.Length - 1]);

            if (!isString)
            {
                int result = 0;
                for (int i = 0; i < parameters.Length - 1; i++)
                {
                    result += Convert.ToInt32(Table[GetNumber(parameters[i]), GetLetter(parameters[i])]);
                }
                Table[targetRow, targetCol] = Convert.ToString(result);
                TableTypes[targetRow, targetCol] = -1;
            }
            else
            {
                string result = "";
                for (int i = 0; i < parameters.Length - 1; i++)
                {
                    result += Table[GetNumber(parameters[i]), GetLetter(parameters[i])];
                }

                bool flag = false;
                while (!flag)
                {
                    flag = true; 
                    Console.Write("Please Select Letter Case (up, low): ");
                    string input = Console.ReadLine();
                    string letterCase = (input ?? "").ToLower().Trim();

                    if (letterCase == "up")
                    {
                        Table[targetRow, targetCol] = result.ToUpper();
                        TableTypes[targetRow, targetCol] = 1; 
                    }
                    else if (letterCase == "low")
                    {
                        Table[targetRow, targetCol] = result.ToLower();
                        TableTypes[targetRow, targetCol] = 1; 
                    }
                    else
                    {
                        ErrorMessage("Please enter a valid value (up or low)");
                        flag = false; 
                    }
                }
            }
        }
        static void Multiply(string[] parameters)
        {
            for (int i = 0; i < 3; i++) // 3 parametre var
            {
                if (GetNumber(parameters[i]) == -1 || GetLetter(parameters[i]) == -1)
                {
                    ErrorMessage("Invalid coordinate format in parameters!");
                    return;
                }
            }
            int number1 = GetNumber(parameters[0]);
            int number2 = GetNumber(parameters[1]);
            int number3 = GetNumber(parameters[2]);
            string letter1 = parameters[0].Trim().Substring(0, 1);
            string letter2 = parameters[1].Trim().Substring(0, 1);
            string letter3 = parameters[2].Trim().Substring(0, 1);
            int letterIndex1 = GetLetter(parameters[0]);
            int letterIndex2 = GetLetter(parameters[1]);
            if (TableTypes[number1, letterIndex1] == -1 && TableTypes[number2, letterIndex2] == -1)
            {
                int result = Convert.ToInt32(Table[number1, letterIndex1]) * Convert.ToInt32(Table[number2, letterIndex2]);
                AssignValue(letter3 + number3, "integer", Convert.ToString(result));
            }
            else if (TableTypes[number1, letterIndex1] == 1 && TableTypes[number2, letterIndex2] == -1)
            {
                int integer = Convert.ToInt32(Table[number2, letterIndex2]);
                string String = Table[number1, letterIndex1];
                string result = "";
                if (integer > 0)
                {
                    for(int i = 0; i < integer; i++)
                    {
                        result += String;
                    }
                }
                else
                {
                    string reverseString = "";
                    for(int i = String.Length - 1; i >= 0; i--)
                    {
                        reverseString += String[i];
                    }
                    for (int i = 0; i < integer * -1; i++)
                    {
                        result += reverseString;
                    }
                }
                AssignValue(letter3 + (number3 + 1), "string", result);
            }
            else if (TableTypes[number1, letterIndex1] == -1 && TableTypes[number2, letterIndex2] == 1)
            {
                int integer = Convert.ToInt32(Table[number1, letterIndex1]);
                string String = Table[number2, letterIndex2];
                string result = "";
                if (integer > 0)
                {
                    for(int i = 0; i < integer; i++)
                    {
                        result += String;
                    }
                }
                else
                {
                    string reverseString = "";
                    for(int i = String.Length - 1; i >= 0; i--)
                    {
                        reverseString += String[i];
                    }
                    for (int i = 0; i < integer * -1; i++)
                    {
                        result += reverseString;
                    }
                }
                AssignValue(letter3 + number3, "string", result);
            }
            else
            {
                ErrorMessage("String String operation is not allowed!");
                return;
            }
        }
        static void XRow(string[] parameters)
        {
            if (!IsInteger(parameters[0]) || !IsInteger(parameters[1]))
            {
                ErrorMessage("You must use an integer!");
                return;
            }
            int number1 = Convert.ToInt32(parameters[0]);
            int number2 = Convert.ToInt32(parameters[1]);

            for (int c = 0; c < currentCol; c++)
            {
                X(letters[c] + number1, letters[c] + number2);
            }
        }
        static void XColumn(string[] parameters)
        {
            string letter1 = parameters[0];
            string letter2 = parameters[1];

            for (int r = 1; r <= currentRow; r++)
            {
                X(letter1 + r, letter2 + r);
            }
        }
        static void X(string parameter1, string parameter2)
        {
            int number1 = GetNumber(parameter1);
            int number2 = GetNumber(parameter2);
            int letter1 = GetLetter(parameter1);
            int letter2 = GetLetter(parameter2);
            if (number1 == -1 || number2 == -1 || letter1 == -1 || letter2 == -1)
            {
                ErrorMessage("Invalid coordinates!");
                return;
            }
            Table[number2, letter2] = Table[number1, letter1];
            TableTypes[number2, letter2] = TableTypes[number1, letter1];
            Table[number1, letter1] = null;
            TableTypes[number1, letter1] = 0;
        }
        static void CopyRow(string[] parameters)
        {
            if (!IsInteger(parameters[0]) || !IsInteger(parameters[1]))
            {
                ErrorMessage("You must use an integer!");
                return;
            }

            int number1 = Convert.ToInt32(parameters[0]);
            int number2 = Convert.ToInt32(parameters[1]);

            for (int c = 0; c < currentCol; c++)
            {
                Copy(letters[c] + number1, letters[c] + number2);
            }
        }
        static void CopyColumn(string[] parameters)
        {
            string letter1 = parameters[0];
            string letter2 = parameters[1];

            for (int r = 1; r <= currentRow; r++)
            {
                Copy(letter1 + r, letter2 + r);
            }
        }
        static void Copy(string parameter1, string parameter2)
        {
            int number1 = GetNumber(parameter1);
            int number2 = GetNumber(parameter2);
            int letter1 = GetLetter(parameter1);
            int letter2 = GetLetter(parameter2);
            if (number1 == -1 || number2 == -1 || letter1 == -1 || letter2 == -1)
            {
                ErrorMessage("Invalid coordinates!");
                return;
            }
            Table[number2, letter2] = Table[number1, letter1];
            TableTypes[number2, letter2] = TableTypes[number1, letter1];
        }
        static void AddColumn(string[] parameters)
        {
            string[,] temp = new string[currentRow, currentCol + 1];
            int[,] tempTypes = new int[currentRow, currentCol + 1];
            int letterIndex = Array.IndexOf(letters, parameters[0].ToLower());
            string direction = parameters[1].Trim().ToLower();
            
            if (direction == "right")
            {
                currentCol++;
                for (int r = 0; r < currentRow; r++)
                {
                    for (int c = letterIndex + 1; c < currentCol - 1; c++)
                    {
                        temp[r, c + 1] = Table[r, c];
                        tempTypes[r, c + 1] = TableTypes[r, c];
                    }
                }
                for (int r = 0; r < currentRow; r++)
                {
                    for (int c = letterIndex + 1; c < currentCol; c++)
                    {
                        Table[r, c] = temp[r, c];
                        TableTypes[r, c] = tempTypes[r, c];
                    }
                }
            }
            else if (direction == "left")
            {
                currentCol++;
                for (int r = 0; r < currentRow; r++)
                {
                    for (int c = letterIndex ; c < currentCol - 1; c++)
                    {
                        temp[r, c + 1] = Table[r, c];
                        tempTypes[r, c + 1] = TableTypes[r, c];
                    }
                }
                for (int r = 0; r < currentRow; r++)
                {
                    for (int c = letterIndex ; c < currentCol; c++)
                    {
                        Table[r, c] = temp[r, c];
                        TableTypes[r, c] = tempTypes[r, c];
                    }
                }
            }
            else
            {
                ErrorMessage("Please enter a valid direction");
                return;
            }
        }
        static void AddRow(string[] parameters)
        {
            if (!IsInteger(parameters[0]))
            {
                ErrorMessage("You must use an integer for the row!");
                return;
            }
            string[,] temp = new string[currentRow + 1, currentCol];
            int[,] tempTypes = new int[currentRow + 1, currentCol];
            int numberIndex = Convert.ToInt32(parameters[0]) - 1;
            string direction = parameters[1].Trim().ToLower();
            if (direction == "up")
            {
                currentRow++; 
                for (int r = numberIndex; r < currentRow - 1; r++)
                {
                    for (int c = 0; c < currentCol; c++)
                    {
                        temp[r + 1, c] = Table[r, c];
                        tempTypes[r + 1, c] = TableTypes[r, c];
                    }
                }
                for (int r = numberIndex; r < currentRow; r++)
                {
                    for (int c = 0; c < currentCol; c++)
                    {
                        Table[r, c] = temp[r, c];
                        TableTypes[r, c] = tempTypes[r, c];
                    }
                }
            }
            else if (direction == "down")
            {
                currentRow++;
                for (int r = numberIndex + 1; r < currentRow - 1; r++)
                {
                    for (int c = 0; c < currentCol; c++)
                    {
                        temp[r + 1, c] = Table[r, c];  
                        tempTypes[r + 1, c] = TableTypes[r, c];  
                    }
                }
                for (int r = numberIndex + 1; r < currentRow ; r++)
                {
                    for (int c = 0; c < currentCol; c++)
                    {
                        Table[r, c] = temp[r, c];
                        TableTypes[r, c] = tempTypes[r, c];
                    }
                }
            }
            else
            {
                ErrorMessage("Please enter a valid direction");
                return;
            }
        }
        static void ClearAll()
        {
            for (int r = 0; r < currentRow; r++)
            {
                for (int c = 0; c < currentCol; c++)
                {
                    if (TableTypes[r, c] != 0)
                    {
                        ClearCell(letters[c] + (r + 1));
                    }
                }
            }
        }
        static void ClearCell(string cell)
        {
            if (cell.Trim().Length < 2)
            {
                ErrorMessage("Invalid command expression!");
                return;
            }
            string letter = cell.Trim().Substring(0, 1);
            string numberPart = cell.Trim().Substring(1);
            if (!IsInteger(numberPart))
            {
                ErrorMessage("Invalid coordinate format!");
                return;
            }
            int number = Convert.ToInt32(numberPart);
            int letterIndex = Array.IndexOf(letters, letter);
            if (CheckRange(number, letter))
            {
                if (TableTypes[number - 1, letterIndex] == 0)
                {
                    ErrorMessage("The cell is already empty!");
                    return;
                }
                TableTypes[number - 1, Array.IndexOf(letters, letter)] = 0;
                Table[number - 1, Array.IndexOf(letters, letter)] = null;
            }
            else
            {
                ErrorMessage("Please enter coordinates within the current boundaries!");
            }
        }
        static bool CheckRange(int number, string letter)
        {
            if (!letters.Contains(letter))
            {
                return false;
            }
            if(number <= currentRow && Array.IndexOf(letters, letter) <= currentCol - 1)
            {
                return true;
            }
            else 
            {
                return false;
            }
        }
        static void AssignValue(string cordinate, string type, string input)
        {
            int letterIndex = GetLetter(cordinate);
            int number = GetNumber(cordinate);

            if (number == -1 || letterIndex == -1)
            {
                ErrorMessage("Invalid coordinate format!");
                return;
            }

            if (CheckRange(number + 1, letters[letterIndex]))  
            {
                if (type.Trim().ToLower() == "integer")
                {
                    if (!IsInteger(input.Trim()))
                    {
                        ErrorMessage("You cannot store a string value as an integer!");
                        return;
                    }
                    TableTypes[number, letterIndex] = -1;
                    Table[number, letterIndex] = input;
                }
                else if (type.Trim().ToLower() == "string")
                {
                    TableTypes[number, letterIndex] = 1;
                    Table[number, letterIndex] = input.Trim();
                }
                else
                {
                    ErrorMessage("Invalid command expression!");
                    return;
                }
            }
            else
            {
                ErrorMessage("Please enter coordinates within the current boundaries");
                return;
            }
        }
        static void DrawTable()
        {
            Console.SetCursorPosition(8, 0);
            for (int ch = 0; ch < currentCol; ch++)
            {
                Console.Write($"{letters[ch].ToUpper()}        ");
            }
            Console.SetCursorPosition(8, 1);
            for (int ch = 0; ch < currentCol; ch++)
            {
                Console.Write("-        ");
            }
            Console.WriteLine();
            for(int r = 0; r < currentRow; r++)
            {
                if (r + 1 > 9)
                {
                    Console.Write($"{r + 1} |");
                }
                else
                {
                    Console.Write($" {r + 1} |");
                }
                for (int c = 0; c < currentCol; c++)
                {
                    Console.Write("        |");
                }
                Console.WriteLine();
            }
            for(int r = 0; r < currentRow; r++)
            {
                for(int c = 0; c < currentCol; c++)
                {
                    string cellValue = (Table[r, c] == null) ? "" : Table[r, c];
                    Console.SetCursorPosition(letterCordinates[c], numberCordinates[r]);
                    if (cellValue.Length > 5)
                    {
                        Console.Write(cellValue.Substring(0, 5) + "_");
                    }
                    else
                    {
                        Console.Write(cellValue);
                    }
                }
            }
            Console.WriteLine();
            Console.WriteLine();
        }
        static void ErrorMessage(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write(message);
            Thread.Sleep(1500);
            Console.ResetColor();
        }
        static void SaveToTxt()
        {
            string fileName = "spreadsheet.txt";
            System.IO.StreamWriter writer = new System.IO.StreamWriter(fileName);
            writer.WriteLine("DIMENSIONS;" + currentRow + ";" + currentCol);
            writer.WriteLine("Coordinate;Type;Value");

            for (int r = 0; r < currentRow; r++)
            {
                for (int c = 0; c < currentCol; c++)
                {
                    if (TableTypes[r, c] != 0)
                    {
                        string coord = letters[c].ToUpper() + (r + 1);
                        string type = (TableTypes[r, c] == -1) ? "integer" : "string";
                        string value = Table[r, c];
                        writer.WriteLine(coord + ";" + type + ";" + value);
                    }
                }
            }
            writer.Close();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Final content saved to " + fileName);
            Console.ResetColor();
            Thread.Sleep(1000);
        }
        static void LoadFromTxt()
        {
            string fileName = "spreadsheet.txt";
            if (!System.IO.File.Exists(fileName))
            {
                return;
            }

            System.IO.StreamReader reader = new System.IO.StreamReader(fileName);
            string line;

            while ((line = reader.ReadLine()) != null)
            {
                if (line.StartsWith("DIMENSIONS"))
                {
                    string[] dims = line.Split(';');
                    currentRow = Convert.ToInt32(dims[1]);
                    currentCol = Convert.ToInt32(dims[2]);
                    continue;
                }

                if (line.StartsWith("Coordinate"))
                {
                    continue;
                }

                string[] parts = line.Split(';');
                if (parts.Length == 3)
                {
                    string coord = parts[0];
                    string type = parts[1];
                    string value = parts[2];
                    int row = GetNumber(coord);
                    int col = GetLetter(coord);

                    if (row != -1 && col != -1 && CheckRange(row + 1, letters[col]))
                    {
                        Table[row, col] = value;
                        if (type == "integer")
                        {
                            TableTypes[row, col] = -1;
                        }
                        else
                        {
                            TableTypes[row, col] = 1;
                        }
                    }
                }
            }
            reader.Close();
        }
    }
}
