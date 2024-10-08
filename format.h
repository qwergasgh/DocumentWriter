#ifndef FORMAT_H
#define FORMAT_H

#include <QString>
#include <QColor>
#include <QVariant>

namespace DocumentWriter {

struct FormatText {
    int size = 14;
    bool bold = false;
    QString family = "Times New Roman";
    int align = Qt::AlignJustify;
    bool underline = false;
    bool italic = false;
    QColor color = Qt::black;
    double indent = 1.5;
};

struct FormatDocument {
    bool landscape = false;
    double marginLeft = 1.5;
    double marginRight = 1;
    double marginTop = 1;
    double marginBottom = 1;
};

struct FormatTable {
    int align = Qt::AlignCenter;
    double border = 1;
    QColor color = Qt::black;
    QString style = "solid";
    int width = 100;
};

struct FormatImage {
    int align = Qt::AlignCenter;
    QColor borberColor = Qt::black;
    int border = 0;
    int width = 0;
    int height = 0;
};
}

#endif // FORMAT_H
