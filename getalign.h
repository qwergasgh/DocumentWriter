#ifndef GETALIGN_H
#define GETALIGN_H

#include <QString>

static QString getAlign(const int &align) {
    QString align_str;
    switch (align) {
    case Qt::AlignLeft:
        align_str = "left";
        break;
    case Qt::AlignRight:
        align_str = "right";
        break;
    case Qt::AlignHCenter:
        align_str = "center";
        break;
    case Qt::AlignCenter:
        align_str = "center";
        break;
    case Qt::AlignJustify:
        align_str = "justify";
        break;
    default:
        align_str = "justify";
        break;
    }
    return align_str;
}

#endif // GETALIGN_H
