using System;
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
    /// Фасад для работы с документом. Открытие, заполение, сохранение.
    /// </summary>
    class MyDocument : IDisposable
    {
        // блок объектов для работы с api word'а
        private Word.Application application;
        private Word.Document document;
        private Object missingObj = System.Reflection.Missing.Value;
        private Object falseObj = false;
        
        // пути к файлам (пределяются в конструкторе)
        private string PatternPath;
        private string ResultPath;

        // видимость документа
        public bool Visible { get { return application.Visible; } set { application.Visible = value; } }

        /// <summary>
        /// Пытаемся открыть шаблон
        /// </summary>
        /// <param name="killAllProcesses">убивать ли все процессы ворда</param>
        /// <param name="askByDialogWindow">спросить ли через диалоговое окно путь к папке, false = взять дефолтный путь</param>
        public MyDocument(bool killAllProcesses, bool askByDialogWindow=true)
        {
            // итоговый путь к папке
            string pathToPatternFolder;
            if (askByDialogWindow) {
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                dialog.SelectedPath = Directory.GetCurrentDirectory();
                dialog.Description = "Выберите папку Patterns, в которой содержится документ pattern.docx";
                while (dialog.ShowDialog() != DialogResult.OK);
                pathToPatternFolder = dialog.SelectedPath;
            }
            else
                pathToPatternFolder = Directory.GetParent(Directory.GetCurrentDirectory()).
                    Parent.Parent.Parent.FullName + "\\Patterns";
            
            PatternPath = pathToPatternFolder + "\\pattern.docx";
            ResultPath = pathToPatternFolder + "\\result.doc";
            if (!File.Exists(PatternPath))
                throw new Exception($"Не найден шаблон заполнения отчета по пути: {PatternPath}");
            if (killAllProcesses)
            {
                Console.Write("Осторожно! Убивает все открытые процессы ворда при запуске.\nДля продолжения нажмите ENTER:");
                Console.ReadLine();
                KillAllWordProcesses();
            }
            Open_doc();
        }
        public void Dispose()
        {
            CloseDoc();
        }
        private void Open_doc()
        {
            try
            {
                //создаем обьект приложения word
                application = new Word.Application();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine($"Ошибка создания приложения ворд, возможно на компьютере не установлен Office Word.");
                CloseDoc();
                throw ex;
            }
            
            // создаем путь к файлу
            Object templatePathObj = PatternPath;
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
                CloseDoc();
                throw ex;
            }
        }
        public void Replase(string strToFind, string replaceStr)
        {
            int max = 250;
            // обьектные строки для Word
            object strToFindObj = strToFind;
            object replaceStrObj = replaceStr;
            //тип поиска и замены
            object replaceTypeObj;
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            // обходим все разделы документа
            while (replaceStr.Length > 0)
            {
                string buf = string.Copy(replaceStr);
                int freePlace = max - strToFind.Length;
                // если строки слишком длинная, чтобы сразу заменять флаг на неё полнось (макс 255 символов на замену)
                if (freePlace <= replaceStr.Length)
                {
                    buf = buf.Remove(freePlace + 1); // в буфере обрезаем сторку до максиально возможной длины
                    buf += strToFind; // добавляем флаг
                    replaceStr = replaceStr.Remove(0, freePlace); // урезаем строку сначала, остается тоьлко то, что не попало в буфер
                }
                else
                    // иначе строка полностью попадает в буфер и мы её затираем
                    replaceStr = "";
                replaceStrObj = buf;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    // берем всю секцию диапазоном
                    //wordRange = document.Sections[i].Range;
                    Word.Find wordFindObj = document.Sections[i].Range.Find;
                    object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
            }
            Object pathToSaveObj = ResultPath;
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
            Object pathToSaveObj = ResultPath;
            document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj);

        }
        private void KillAllWordProcesses()
        {
            Console.WriteLine("Убиваю все процессы word ...");
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
