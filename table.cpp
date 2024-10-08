#include "table.h"

DocumentWriter::Table::Table(const FormatTable &format,
                             const int &rows,
                             const int &columns) {
    if (rows != 0 && columns != 0) {
        _rows = rows;
        _columns = columns;
        _createTable();
    }
    _format = format;
}

void DocumentWriter::Table::_createTable() {
    for (int i = 0; i < _rows; i++) {
        QList<Cell *> listRowCells;
        for (int j = 0; j < _columns; j++) {
            Cell *cell = new Cell;
            cell->row = i;
            cell->column = j;
            listRowCells.push_back(cell);
        }
        _cells.insert(i, listRowCells);
    }
}

void DocumentWriter::Table::addRow(const int &n) {
    _addCells(n, ROW);
}

void DocumentWriter::Table::addColumn(const int &n) {
    _addCells(n, COLUMN);
}

void DocumentWriter::Table::_addCells(const int &n, const int &argument) {
    for (int i = 0; i < n; i++) {
        if (argument == ROW) {
            _rows++;
            QList<Cell *> listRowCells;
            for (int j = 0; j < _columns; j++) {
                Cell *cell = new Cell;
                cell->row = _rows - 1;
                cell->column = j;
                listRowCells.push_back(cell);
            }
            _cells.insert(_rows - 1, listRowCells);
            continue;
        }
        if (argument == COLUMN) {
            _columns++;
            for (int j = 0; j < _rows; j++) {
                Cell *cell = new Cell;
                cell->row = j;
                cell->column = _columns - 1;
                QList<Cell *> listRowCells = _cells.value(j);
                listRowCells.push_back(cell);
                _cells.insert(j, listRowCells);
            }
            continue;
        }
    }
}

void DocumentWriter::Table::setValue(const int &row,
                                     const int &column,
                                     const QString &value,
                                     const FormatText &format) {
    if (row >= _rows || column >= _columns) {
        qDebug() << "Error! - row > size || column > size";
        return;
    }
    foreach (Cell *cell, _cells.value(row)) {
        if (cell->column == column) {
            cell->value = value;
            cell->format = format;
        }
    }
}

void DocumentWriter::Table::mergeRange(const int &rowSrc,
                                       const int &rowDst,
                                       const int &columnSrc,
                                       const int &columnDst,
                                       const QString &value,
                                       const FormatText &format) {
    if (rowSrc > rowDst ||
        columnSrc > columnDst ||
        rowDst >= _rows ||
        columnDst >= _columns) {
        qDebug() << "Error! - row > size || column > size";
        return;
    }
    for (int i = rowSrc; i <= rowDst; i++) {
        foreach (Cell *cell, _cells.value(i)) {
            if (columnSrc != columnDst && rowSrc == rowDst) {
                if (cell->column == columnSrc) {
                    cell->columnSpan = true;
                    cell->sizeColumnSpan = columnDst - columnSrc + 1;
                    cell->value = value;
                    cell->format = format;
                    continue;
                }
                else if (cell->column <= columnDst && cell->column > columnSrc) {
                    cell->nullCell = true;
                    continue;
                }
            }
            if (rowSrc != rowDst && columnSrc == columnDst && columnSrc == cell->column) {
                if (cell->row == rowSrc) {
                    cell->rowSpan = true;
                    cell->sizeRowSpan = rowDst - rowSrc + 1;
                    cell->value = value;
                    cell->format = format;
                    continue;
                }
                else if (cell->row <= rowDst && cell->row > rowSrc) {
                    cell->nullCell = true;
                    continue;
                }
            }
            if (rowDst > rowSrc && columnDst > columnSrc) {
                if (cell->row == rowSrc && cell->column == columnSrc) {
                    cell->rowSpan = true;
                    cell->columnSpan = true;
                    cell->sizeRowSpan = rowDst - rowSrc + 1;
                    cell->sizeColumnSpan = columnDst - columnSrc + 1;
                    cell->value = value;
                    cell->format = format;
                    continue;
                }
                else if (cell->row <= rowDst && cell->row >= rowSrc &&
                           cell->column <= columnDst && cell->column >= columnSrc) {
                    cell->nullCell = true;
                    continue;
                }
            }
        }
    }
}

int DocumentWriter::Table::rowCount() {
    return _rows;
}

int DocumentWriter::Table::columnCount() {
    return _columns;
}

void DocumentWriter::Table::_fillingTable() {
    _table += "<table width='" + QString::number(_format.width) +
              "%' align='" + getAlign(_format.align) +
              "' cellpadding='2' cellspacing='0'>\n";
    QMapIterator<int, QList<Cell *>> count(_cells);
    while (count.hasNext()) {
        count.next();
        _table += "<tr>\n";
        for (int i = 0; i < count.value().size(); i++) {
            Cell *cell = count.value().at(i);
            if (cell->nullCell) continue;
            _table += "<td ";
            if (cell->rowSpan) {
                _table += "rowspan='" + QString::number(cell->sizeRowSpan) + "' ";
            }
            if (cell->columnSpan) {
                _table += "colspan='" + QString::number(cell->sizeColumnSpan) + "' ";
            }
            _table += "style='border: " + QString::number(_format.border) + "pt " +
                      _format.style + " " + _format.color.name();

            if (!cell->columnSpan && !cell->rowSpan) {
                if (i != count.value().size() - 1 && count.key() != _cells.size() - 1) {
                    _table += "; border-bottom: none; border-right: none";
                }
                if (i != count.value().size() - 1 && count.key() == _cells.size() - 1) {
                    _table += "; border-right: none";
                }
                if (i == count.value().size() - 1 && count.key() != _cells.size() - 1) {
                    _table += "; border-bottom: none";
                }
            }
            if (cell->columnSpan && !cell->rowSpan) {
                if (i + cell->sizeColumnSpan != count.value().size() &&
                    i != count.value().size() - 1 && count.key() != _cells.size() - 1) {
                    _table += "; border-bottom: none; border-right: none";
                }
                if (i + cell->sizeColumnSpan != count.value().size() &&
                    i != count.value().size() - 1 && count.key() == _cells.size() - 1) {
                    _table += "; border-right: none";
                }
                if (i + cell->sizeColumnSpan == count.value().size() &&
                    count.key() != _cells.size() - 1) {
                    _table += "; border-bottom: none";
                }
            }
            if (!cell->columnSpan && cell->rowSpan) {
                if (cell->row + cell->sizeRowSpan != _cells.size() &&
                    i != count.value().size() - 1) {
                    _table += "; border-bottom: none; border-right: none";
                }
                if (cell->row + cell->sizeRowSpan == _cells.size() &&
                    i != count.value().size() - 1) {
                    _table += "; border-right: none";
                }
                if (cell->row + cell->sizeRowSpan != _cells.size() &&
                    i == count.value().size() - 1) {
                    _table += "; border-bottom: none";
                }
            }
            if (cell->columnSpan && cell->rowSpan) {
                if (cell->row + cell->sizeRowSpan != _cells.size() &&
                    i != count.value().size() - 1) {
                    _table += "; border-bottom: none; border-right: none";
                }
                if (cell->row + cell->sizeRowSpan == _cells.size() &&
                    i != count.value().size() - 1) {
                    _table += "; border-right: none";
                }
                if (cell->row + cell->sizeRowSpan != _cells.size() &&
                    i == count.value().size() - 1) {
                    _table += "; border-bottom: none";
                }
            }

            _table +=  "; padding: 0cm'>\n<p align='" + getAlign(cell->format.align) +
                       "' style='font-size: " + QString::number(cell->format.size) +
                       "pt; font-family: " + cell->format.family + "; ";
            if (cell->format.bold) {
                _table += "font-weight: bold; ";
            }
            if (cell->format.underline) {
                _table += "text-decoration: underline; ";
            }
            if (cell->format.italic) {
                _table += "font-style: italic; ";
            }
            _table += "margin-top: 0px; margin-bottom: 0px;'>" +
                      cell->value + "</p>\n</td>\n";
        }
        _table += "</tr>\n";
    }
    _table += "</table>\n";
}

QString DocumentWriter::Table::getTable() {
    _fillingTable();
    return _table;
}
