using System;
using Word = Microsoft.Office.Interop.Word;

namespace MSOfficeExportApp.Service.Impl
{
    public class WordExport: IWordExport
    {
        public void Export(String dir, String text)
        {
            // Создание приложения
            Word._Application word = new Word.Application() { Visible = true };
       
            // Создание документа
            Word._Document doc = word.Documents.Add();

            // Создание абзаца
            Word.Paragraph paragraph = doc.Paragraphs.Add();
            paragraph.Range.Text = text;

            // Сохранение документа
            doc.SaveAs(dir + "\\exampleDocx.docx");
            doc.Close();
            word.Quit();
        }

        public void ExportDotx(string dir, string text)
        {
            object template = dir +  "\\exampleDotx.dotx";
            Word._Application word = new Word.Application() { Visible = true };
            Word._Document doc = word.Documents.Add(template);

            // Поиск закладки Name и вставка введенного текста
            String nameMark = "Name";
            Word.Bookmarks bookmarks = doc.Bookmarks;

            foreach (Word.Bookmark mark  in bookmarks)
            {
                if (mark.Name == nameMark) mark.Range.Text = text;
            }

            doc.SaveAs(dir + "\\newTemplate.docx");
            doc.Close();
            word.Quit();
        }
    }
}
