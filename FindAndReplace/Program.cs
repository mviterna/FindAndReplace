// ▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
//  Project: Find And Replace Multiple Word Documents
//  Date   : May, 21 2015
//  File   : Program.cs
//  Purpose: A quick .NET console app that allows users to find and replace a string across multiple Word documents at once
//  Usage:            
//  ──────────────────────────────────────────────────────────────────────────────────────────────
//Copyright 2015 Mark Viterna

//This code is licensed under the MIT License (MIT).

//Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal 
//in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or 
//sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// ▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀

using System;

namespace FindAndReplaceWord
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string _textToReplace;
            string _textToFind;
            ParsingEngine _pEngine;

            Console.WriteLine("My tool will traverse a given folder for Microsoft Office documents and do a find and replace for a given string.");

            Console.WriteLine("This is helpful when updating a certain value across many documents, that would otherwise be done by hand. ");

            Console.WriteLine("Example usage: Change all occurences of the year '2016' to '2017'");

            Console.WriteLine(Environment.NewLine);

            Console.WriteLine("Please enter the path of the directory containing the Word docuements (ex: C:\\MyWordDocs): ");

            _pEngine = new ParsingEngine(Console.ReadLine());

            Console.WriteLine("Please enter the text to find: ");
            _textToFind = Console.ReadLine();

            Console.WriteLine("Please enter the text to replace it with: ");
            _textToReplace = Console.ReadLine();

            _pEngine.TextToReplace = _textToReplace;
            _pEngine.TextToFind = _textToFind;
            _pEngine.Start();
            Console.ReadLine();
        }
    }

}

