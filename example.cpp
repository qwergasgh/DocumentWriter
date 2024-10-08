#include <QCoreApplication>
#include <libDocumentWriter/DocumentWriter.h>
#include <QDebug>

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    DocumentWriter::FormatDocument formatDocument;
    formatDocument.landscape = true;
    DocumentWriter::Document document("path", formatDocument);

    DocumentWriter::FormatText fontText;
    fontText.align = Qt::AlignHCenter;

    document.setText("txt", fontText);
    document.setParagraph(3);

    DocumentWriter::FormatTable formatTable;
    formatTable.align = Qt::AlignLeft;
    formatTable.width = 100;
    DocumentWriter::Table table(formatTable, 3, 3);

    table.setValue(0, 0, "0-0", fontText);
    table.setValue(0, 1, "0-1", fontText);
    table.setValue(0, 2, "0-2", fontText);
    table.setValue(1, 0, "1-0", fontText);
    table.setValue(1, 1, "1-1", fontText);
    table.setValue(1, 2, "1-2", fontText);
    table.setValue(2, 0, "2-0", fontText);
    table.setValue(2, 1, "2-1", fontText);
    table.setValue(2, 2, "2-2", fontText);

    table.addRow();
    table.setValue(3, 0, "3-0", fontText);
    table.setValue(3, 1, "3-1", fontText);
    table.setValue(3, 2, "3-2", fontText);

    table.addColumn();
    table.setValue(0, 3, "0-3", fontText);
    table.setValue(1, 3, "1-3", fontText);
    table.setValue(2, 3, "2-3", fontText);
    table.setValue(3, 3, "3-3", fontText);
    //table.mergeRange(0, 3, 0, 0, "row span 0-3", fontText);
    //table.mergeRange(0, 0, 1, 3, "column span 1-3", fontText);
    //table.mergeRange(2, 3, 2, 3, "row column span 2-3 2-3", fontText);
    table.mergeRange(0, 3, 3, 3, "row span 0-3", fontText);
    document.setTable(table.getTable());

    //    DocumentWriter::FormatImage formatImage;
    //    document.setImage("path", formatImage);

    document.setRunningTitleHeader("Header", fontText);

    DocumentWriter::FormatText formatNumeration;
    formatNumeration.align = Qt::AlignCenter;
    formatNumeration.indent = 0;
    formatNumeration.bold = true;
    formatNumeration.underline = true;
    formatNumeration.size = 16;
    document.setNumerationPages(formatNumeration);
	
    document.closeDocument();

    return a.exec();
}
