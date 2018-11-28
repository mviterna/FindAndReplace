// ▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
//  Project: Find And Replace Multiple Word Documents
//  Date   : May, 21 2015
//  File   : FileCollector.cs
//  Purpose: Simple file class
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
// ▀▀▀

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindAndReplace
{
    class FileCollector
    {
        public FileCollector(string _pathToFiles)
        {
            DirectoryInfo _dir = new DirectoryInfo(_pathToFiles);
            if(!_dir.Exists)
            {
                Console.WriteLine("Directory does not exist. Please try again.");
                _pathToFiles = Console.ReadLine();
                _dir = new DirectoryInfo(_pathToFiles);
            }
            try
            {
                //We need to filter out the temp files Word generates
                this.Files = _dir.GetFiles("*.docx").Where(name => !name.Name.Contains("~"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Collection of files
        /// </summary>
        public IEnumerable<FileInfo> Files { get; set; }
    }
}
