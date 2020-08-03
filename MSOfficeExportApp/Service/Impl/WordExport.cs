using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace MSOfficeExportApp.Service.Impl
{
    public class WordExport: IWordExport
    {
        public void Export(String dir, String text)
        {
            // Создание приложения
            // Получить объект приложения Word.
            Word.Application word = new Word.Application();
            word.Visible = true;


            // Создание документа
            Word.Document doc = word.Documents.Add();

            // Создание абзаца
            Word.Paragraph paragraph;
            paragraph = doc.Paragraphs.Add();
            paragraph.Range.Text = text;

            // Сохранение документа
            doc.SaveAs2(dir);
            doc.Close();
            word.Quit();

        }
    }
}
