﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

namespace Eminem
{
    /// <summary>
    /// Фасад для работы с документом. Открытие, запорление, сохранение.
    /// </summary>
    class MyDocument : IDisposable
    {
        private Word.Application application;
        private Word.Document document;
        private Object missingObj = System.Reflection.Missing.Value;
        private Object falseObj = false;
        
        private string FilePatch;
        private string SaveFilePatch;
        public bool Visible { get { return application.Visible; } set { application.Visible = value; } }

        public MyDocument(bool killAllProcesses, bool autoFind=true)
        {
            string pathToPatternFolder;
            if (!autoFind) {
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                dialog.SelectedPath = Directory.GetCurrentDirectory();
                dialog.Description = "Выберите папку Patterns, в которой содержится документ pattern.docx";
                while (dialog.ShowDialog() != DialogResult.OK);
                pathToPatternFolder = dialog.SelectedPath;
            }
            else
                pathToPatternFolder = Directory.GetParent(Directory.GetCurrentDirectory()).
                    Parent.Parent.Parent.FullName + "\\Patterns";
            
            FilePatch = pathToPatternFolder + "\\pattern.docx";
            SaveFilePatch = pathToPatternFolder + "\\result.doc";
            if (!File.Exists(FilePatch))
                throw new Exception($"Не найден шаблон заполнение отчета по пути: {FilePatch}");
            if (killAllProcesses)
                KillAllWordProcesses();
            Open_doc();
        }
        public void Dispose()
        {
            CloseDoc();
        }
        private void Open_doc()
        {
            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = FilePatch;
            try
            {
                document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine(
                    $"Ошибка доступа. Откройте шаблонный документ и разрешите редактирование. " +
                    $"Шаблонный документ нах-ся по адресу: {templatePathObj}"
                );
                throw ex;
            }
            catch (Exception ex)
            {
                if (document != null)
                    document.Close(ref falseObj, ref missingObj, ref missingObj);
                application.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                application = null;
                throw ex;
            }
        }
        public void Replase(string strToFind, string replaceStr)
        {
            // обьектные строки для Word
            object strToFindObj = strToFind;
            object replaceStrObj = replaceStr;
            // диапазон документа Word
            Word.Range wordRange;
            //тип поиска и замены
            object replaceTypeObj;
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            int z = 0;
            while (replaceStr.Length > 0)
            // обходим все разделы документа
            {
                bool ch = false;
                if (255 - strToFind.Length > replaceStr.Length)
                    z = replaceStr.Length;
                else
                { z = 255 - strToFind.Length; ch = true; }
                string buf = string.Copy(replaceStr);

                if (ch)
                { buf += strToFind; buf = buf.Remove(z + 1); }
                replaceStrObj = buf;
                for (int i = 1; i <= document.Sections.Count; i++)
                {

                    // берем всю секцию диапазоном
                    wordRange = document.Sections[i].Range;

                    Word.Find wordFindObj = wordRange.Find;


                    object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };


                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                replaceStr = replaceStr.Remove(0, z);
            }
            Object pathToSaveObj = SaveFilePatch;
            document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj);

        }
        public void ReplaseTable(string strToFind, string[,] replaceStr, string[] headers, int tab_num)
        {
            // обьектные строки для Word
            object strToFindObj = strToFind;
            object replaceStrObj = replaceStr;
            // диапазон документа Word
            Word.Range wordRange;
            //тип поиска и замены
            object replaceTypeObj;
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            // обходим все разделы документа
            //for (int i = 1; i <= document.Sections.Count; i++)
            {

                // берем всю секцию диапазоном
                wordRange = document.Range();
                Word.Range buf_range;
                Word.Table table = document.Tables[tab_num];
                for (int rows = 0; rows < replaceStr.GetLength(0); rows++)
                    table.Rows.Add();
                for (int col = 0; col < replaceStr.GetLength(1); col++)
                {
                    for (int rows = 0; rows < replaceStr.GetLength(0); rows++)
                    {
                        buf_range = table.Cell(rows + 2, col + 1).Range;
                        string buf = replaceStr[rows, col];
                        buf_range.Text = buf;
                    }
                }

                /*
                Word.Find wordFindObj = wordRange.Find;
               // Word.Table table = document.Tables[tables_count];
               // tables_count++;
                Word.Range buf_range;
                Word.Table table = new Word.Table;
                for(int col=0;i< replaceStr.GetLength(0); i++)
                    table.Columns.Add();
                for (int rows = 0; i < replaceStr.GetLength(1)+1; i++)
                    table.Rows.Add();
                for (int col = 0; i < replaceStr.GetLength(0); i++)
                {
                    buf_range = table.Cell(col, 1).Range;
                    buf_range.Text = headers[col];
                    for (int rows = 1; i < replaceStr.GetLength(1); i++)
                    {
                        buf_range = table.Cell(col, rows).Range;
                        buf_range.Text = replaceStr[rows,col];
                    }
                }
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, table, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
*/
            }
            Object pathToSaveObj = SaveFilePatch;
            document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj);

        }
        private void KillAllWordProcesses()
        {
            if (Process.GetProcessesByName("winword").Count() > 0)
            {
                System.Diagnostics.Process[] aProcWrd = System.Diagnostics.Process.GetProcessesByName("WINWORD");

                foreach (System.Diagnostics.Process oProc in aProcWrd)
                {
                    oProc.Kill();
                }
            }
        }
        private void CloseDoc()
        {
            Console.WriteLine("Закрываю документ");
            if (document != null)
                document.Close(ref falseObj, ref missingObj, ref missingObj);
            if (application != null)
                application.Quit(ref missingObj, ref missingObj, ref missingObj);
        }
    }
}
