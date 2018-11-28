// ▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
//  Project: Find And Replace Multiple Word Documents
//  Date   : May, 21 2015
//  File   : ParsingEngine.cs
//  Purpose: The main engine class that does the finding and replacing
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

using FindAndReplace;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace FindAndReplaceWord
{
    class ParsingEngine
    {
        #region Fields
        private FileCollector _fileCollector;
        #endregion

        #region Constructor
        public ParsingEngine(string pathToFiles)
        {
            this.PathToFiles = pathToFiles;
            this._fileCollector = new FileCollector(pathToFiles);
        }
        #endregion

        #region Properties
        /// <summary>
        /// The text to find in the documents
        /// </summary>
        public string TextToFind { get; set; }

        /// <summary>
        /// The text to replace in the documents
        /// </summary>
        public string TextToReplace { get; set; }

        /// <summary>
        /// File path of Word docs. Must be .docx extension. Does NOT include sub-directories
        /// </summary>
        public string PathToFiles { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// This is the main start method to begin the find and replace process. This must be manually called. It kicks off an individual thread for each file.
        /// </summary>
        public void Start()
        {           
            try
            {
                foreach (FileInfo file in _fileCollector.Files)
                {        
                    //Create a new thread for each Word document
                    FileInfo _tmp = file;
                    Thread _fileThread = new Thread(() => this.ReplaceInFile(_tmp));
                    _fileThread.Start();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
                        
            return;
        }

        /// <summary>
        /// The constructor loads the file on instantiation, but if you want to reload a different file path, call this method
        /// </summary>
        /// <param name="filePath"></param>
        public void SetFilePath(string pathToFiles)
        {
            this.PathToFiles = pathToFiles;
            this._fileCollector = new FileCollector(pathToFiles);
        }
        private void ReplaceInFile(FileInfo file)
        {
            try
            {
                ReplaceText(file);
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        /// <summary>
        /// The main find and replace method
        /// </summary>
        /// <param name="file"></param>
        private void ReplaceText(FileInfo file)
        {
            try
            {
                Console.WriteLine("Processing File: " + file.Name);

                Application _tmpApp = new Application();
                Document _tmpDoc = _tmpApp.Documents.Open(file.FullName, ReadOnly: false, Visible: false);
                Find findObject = _tmpDoc.Content.Find;
                findObject.ClearFormatting();
                findObject.Replacement.ClearFormatting();

                object replaceAll = WdReplace.wdReplaceAll;
                findObject.Execute(FindText: this.TextToFind, ReplaceWith: this.TextToReplace, Replace: 2, Wrap: 1);
                Console.WriteLine("File: " + file.Name + " is processed");

                //After the files are processed, we need to do some cleanup so there are not leftover MS Word instances running
                _tmpDoc.Close();
                _tmpApp.Quit(SaveChanges: true);
                if (_tmpDoc != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_tmpDoc);
                if (_tmpApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_tmpApp);
                _tmpDoc = null;
                _tmpApp = null;
                GC.Collect();

                //TODO
                //Add a callback to inform the user when all files are done processing
                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
        #endregion
    }
}
