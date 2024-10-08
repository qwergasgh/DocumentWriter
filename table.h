#ifndef TABLE_H
#define TABLE_H

#include "format.h"
#include "getalign.h"
#include "documentwriter_global.h"

#include <QDebug>

namespace DocumentWriter {
class DOCUMENTWRITERSHARED_EXPORT Table {
public:
    Table(const FormatTable &format, const int &rows = 0, const int &columns = 0);
    QString getTable();
    void addRow(const int &n = 1);
    void addColumn(const int &n = 1);
    void setValue(const int &row, const int &column,
                  const QString &value, const FormatText &format);
    void mergeRange(const int &rowSrc, const int &rowDst,
                    const int &columnSrc, const int &columnDst,
                    const QString &value, const FormatText &format);
    int rowCount();
    int columnCount();

private:
    QString _table;
    FormatTable _format;

    int _rows = 0;
    int _columns = 0;

    enum ARG {
        ROW = 0,
        COLUMN = 1
    };

    struct Cell {
        QString value;
        int row = 0;
        int column = 0;
        FormatText format;

        //merge
        bool rowSpan = false;
        bool columnSpan = false;
        int sizeRowSpan = 0;
        int sizeColumnSpan = 0;
        bool nullCell = false;
    };

    QMap<int, QList<Cell *>> _cells;

    void _createTable();
    void _addCells(const int &n, const int &argument);
    void _fillingTable();

};
}
#endif // TABLE_H
