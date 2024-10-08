#ifndef DOCUMENT_H
#define DOCUMENT_H

#include "documentwriter_global.h"
#include "format.h"
#include "getalign.h"
#include <QCoreApplication>
#include <QTextStream>
#include <QBuffer>
#include <QImage>
#include <QFile>
#include <QDebug>

#ifdef Q_OS_WIN32
#include "windows.h"
#include <QAxObject>
#endif

namespace DocumentWriter {
class DOCUMENTWRITERSHARED_EXPORT Document {
public:
    Document(const QString &fileName, FormatDocument &format);
    void setText(const QString &text, const struct FormatText &format);
    void setParagraph(const int &n = 1);
    void setTable(const QString &table);
    void setImage(const QString &image, const FormatImage &format);
    void setRunningTitleHeader(const QString &text, const struct FormatText &format);
    void setRunningTitleFooter(const QString &text, const struct FormatText &format);
    void setNumerationPages(const struct FormatText &format);
    void closeDocument();

private:
    const QString _suffixHTML= ".html";
    const QString _footer = "footer";
    const QString _header = "header";


    QString _suffix;
    QString _doc;
    QString _fileName;
    QString _name;
    QString _path;

    void _createHTML(FormatDocument &format);
    QString _getPropertiesImage(const QString &base, const QString &name,
                                const QString &align, const QString &width,
                                const QString &height, const QString &border);
    QString _getPropertiesDocument(const QString &marginLeft, const QString &marginRight,
                                   const QString &marginTop, const QString &marginBottom,
                                   bool landscape = false);
    QString _getPropertiesText(const QString &text, const struct FormatText &format);
    void _setPropertiesRunningTitle(const QString &disposition, const FormatText &format,
                                    const QString &text, const bool numeration = false);
    void _swapDoc(QString &a, QString &b);
    void _setLastDoc(QString &doc, QStringList &docList);
};
}

#endif // DOCUMENT_H
