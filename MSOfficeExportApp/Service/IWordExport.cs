using System;


namespace MSOfficeExportApp.Service
{
    public interface IWordExport
    {
        void Export(String dir, String text);

        // Работа с шаблоном  Word и Bookmarks
        void Replace(String dir, String text);

    }
}
