using System;
using Word = Microsoft.Office.Interop.Word;

namespace MSOfficeExportApp.Service.Impl
{
    public class WordExport: IWordExport
    {
        public void Export(String dir, String text)
        {
            // Создание приложения
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

        public void Replace(string dir, string text)
        {
            object template = "F:\\CIT\\Others\\bookmarks.dotx";
            Word.Application word = new Word.Application();
            word.Visible = true;

            Word.Document doc = word.Documents.Open(template);

            String nameMark = "Name";
            Word.Bookmarks bookmarks = doc.Bookmarks;
           
            foreach (Word.Bookmark mark  in bookmarks)
            {
                if (mark.Name == nameMark) mark.Range.Text = text;
            }

            doc.Save();
            doc.Close();
            word.Quit();
        }
    }
}
