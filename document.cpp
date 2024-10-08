#include "document.h"

DocumentWriter::Document::Document(const QString &fileName, FormatDocument &format) {
    QStringList listName = fileName.split(".");
    _fileName = listName.first();
    _suffix = listName.last();
    _name = _fileName.split("/").last();
    _createHTML(format);
}

void DocumentWriter::Document::setText(const QString &text, const FormatText &format) {
    if (text.isEmpty()) {
        qDebug() << "Error! - text is empty";
        return;
    }
    _doc+=_getPropertiesText(text, format);
}

void DocumentWriter::Document::setParagraph(const int &n) {
    if (n <= 0) {
        qDebug() << "Error! - n <= 0";
        return;
    }
    for (int i=0; i < n; i++) {
        _doc += "<p style='margin-top: 0px; margin-bottom: 0px;'><br></p>\n";
    }
}

void DocumentWriter::Document::setTable(const QString &table) {
    if (table.isEmpty() || !table.contains("<table")) {
        qDebug() << "Error! - no table";
        return;
    }
    _doc += table;
}

void DocumentWriter::Document::setImage(const QString &image, const FormatImage &format) {
    QImage img(image);

    QString width = QString::number(img.width());
    QString height = QString::number(img.height());

    QStringList listFileName = image.split("/");
    QStringList tmpName = listFileName.last().split(".");
    QString name = tmpName.first();

    QString tmpHtml;
    QString tmpHtmlResize;

    QBuffer buffer;
    if (buffer.open(QIODevice::WriteOnly)) {
        img.save(&buffer, "png");
        auto const base = buffer.data().toBase64();

        tmpHtml = _getPropertiesImage(base,
                                      name,
                                      getAlign(format.align),
                                      width,
                                      height,
                                      QString::number(format.border));
        tmpHtmlResize = _getPropertiesImage(base,
                                            name,
                                            getAlign(format.align),
                                            QString::number(format.width)+ "%",
                                            QString::number(format.height) + "%",
                                            QString::number(format.border));

    }
    buffer.close();
    if (format.width != 0 && format.height != 0) {
        if (format.border != 0) {
            _doc += "<font color='" + format.borberColor.name() + "'>" + tmpHtmlResize;
        } else {
            _doc += tmpHtmlResize;
        }
    } else {
        if (format.border != 0) {
            _doc += "<font color='" + format.borberColor.name() + "'>" + tmpHtml;
        } else {
            _doc += tmpHtml;
        }
    }
}

void DocumentWriter::Document::setRunningTitleHeader(const QString &text,
                                                     const FormatText &format) {
    _setPropertiesRunningTitle(_header, format, text);
}

void DocumentWriter::Document::setRunningTitleFooter(const QString &text,
                                                     const FormatText &format) {
    _setPropertiesRunningTitle(_footer, format, text);
}

void DocumentWriter::Document::setNumerationPages(const struct FormatText &format) {
    _setPropertiesRunningTitle(_footer, format, nullptr, true);
}

void DocumentWriter::Document::closeDocument() {
    _doc +="</div>\n</body>\n</html>";
    QString nameHTML = _fileName + _suffixHTML;
    QFile file(nameHTML);
    if (file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QTextStream out(&file);
        out.setCodec("utf-8");
        out << _doc;
        file.close();
    }
    _doc.clear();

#ifdef Q_OS_WIN32
    HRESULT result = OleInitialize(nullptr);
    if (result != S_OK && result != S_FALSE) {
        qDebug() << "Error! - HRESULT not ok or not false";
        return;
    }

    QAxObject *_wordApp = new QAxObject("Word.Application");
    if (_wordApp->isNull()) {
        qDebug() << "Error! - not open Word.Application";
        return;
    }

    QAxObject *_documents = _wordApp->querySubObject("Documents");
    if (_documents->isNull()) {
        qDebug() << "Error! - not open Documents";
        return;
    }

    const QVariant namehtml(nameHTML);
    const QVariant confirmconversions(false);
    const QVariant readonly(false);
    const QVariant addtorecentfiles(false);
    const QVariant passworddocument("");
    const QVariant passwordtemplate("");
    const QVariant revert(false);
    QAxObject *_document = _documents->querySubObject("Open(const QVariant&, "
                                                           "const QVariant&,"
                                                           "const QVariant&,"
                                                           "const QVariant&,"
                                                           "const QVariant&,"
                                                           "const QVariant&,"
                                                           "const QVariant&)",
                                                      namehtml,
                                                      confirmconversions,
                                                      readonly,
                                                      addtorecentfiles,
                                                      passworddocument,
                                                      passwordtemplate,
                                                      revert);
    if (_document->isNull()) {
        qDebug() << "Error! - not open Document";
        return;
    }

    QAxObject *_save = _wordApp->querySubObject("ActiveDocument");
    if (_save->isNull()) {
        qDebug() << "Error! - not open ActiveDocument";
        return;
    }

    const QVariant nameDOC(_fileName + ".docx");
    const QVariant formatDoc(16);
    const QVariant LockComments(false);
    const QVariant Password("");
    const QVariant recent(true);
    const QVariant writePassword("");
    const QVariant ReadOnlyRecommended(false);
    _save->querySubObject("SaveAs(const QVariant&, "
                                 "const QVariant&,"
                                 "const QVariant&,"
                                 "const QVariant&,"
                                 "const QVariant&,"
                                 "const QVariant&,"
                                 "const QVariant&)",
                          nameDOC,
                          formatDoc,
                          LockComments,
                          Password,
                          recent,
                          writePassword,
                          ReadOnlyRecommended);

    _document->dynamicCall("Close(const QVariant&)", false);
    _wordApp->dynamicCall("Quit(void)");

    _wordApp = nullptr;
    _documents = nullptr;
    _document = nullptr;
    _save = nullptr;

    QFile::remove(nameHTML);
#else
    system(QString("loffice --convert-to " + _suffix +" " + nameHTML).toStdString().c_str());
    QFile::copy(QCoreApplication::applicationDirPath() + "/" + _name + "." + _suffix,
                _fileName + "." + _suffix);
    QFile::remove(QCoreApplication::applicationDirPath() + "/" + _name + "." + _suffix);
#endif
}

void DocumentWriter::Document::_createHTML(FormatDocument &format) {
    _doc += "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//RU'>\n"
            "<html>\n<head>\n<meta http-equiv='content-type' content='text/html;"
            " charset=utf-8'/>\n<style type='text/css'>\n";
    if (format.landscape) {
        _doc += _getPropertiesDocument(QString::number(format.marginLeft),
                                       QString::number(format.marginRight),
                                       QString::number(format.marginTop),
                                       QString::number(format.marginBottom),
                                       true);
    }
    else {
        _doc += _getPropertiesDocument(QString::number(format.marginLeft),
                                       QString::number(format.marginRight),
                                       QString::number(format.marginTop),
                                       QString::number(format.marginBottom));
    }
    _doc += "p { font-family:'Times New Roman'; so-language: ru-RU; }\n"
            "td p { font-family:'Times New Roman'; so-language: ru-RU; }\n"
            "strong { font-weight: bold; font-size: 14pt; }\n"
            "div.Word {page: Word;}\n</style>\n</head>\n<body lang='ru-RU'>\n"
            "<div class=Word>\n";
}

QString DocumentWriter::Document::_getPropertiesDocument(const QString &marginLeft,
                                                         const QString &marginRight,
                                                         const QString &marginTop,
                                                         const QString &marginBottom,
                                                         bool landscape) {
    QString html;
    if (landscape) {
        html += "@page Word {size: 29.7cm 21cm; mso-page-orientation:landscape; "
                "mso-header-margin:0.5in; mso-header: h1; mso-footer-margin:0.5in;"
                " mso-footer: f1;";
    } else {
        html += "@page { size: 21cm 29.7cm; ";
    }
    html += "margin-left: " + marginLeft +
            "cm; margin-right: " + marginRight +
            "cm; margin-top: " + marginTop +
            "cm; margin-bottom: " + marginBottom + "cm }\n";
    return html;
}

QString DocumentWriter::Document::_getPropertiesText(const QString &text,
                                                     const FormatText &format) {
    QString tmp;
    tmp += "<p align='" + getAlign(format.align) +
           "', style='font-size: " + QString::number(format.size) +
           "pt; font-family: " + format.family +
           "; text-indent: " + QString::number(format.indent) + "cm; ";
    if (format.bold) {
        tmp += "font-weight: bold; ";
    }
    if (format.underline) {
        tmp += "text-decoration: underline; ";
    }
    if (format.italic) {
        tmp += "font-style: italic; ";
    }
    tmp += "color: " + format.color.name() +
           "margin-top: 0px; margin-bottom: 0px;'>" + text + "</p>\n";
    return tmp;
}

void DocumentWriter::Document::_setPropertiesRunningTitle(const QString &disposition,
                                                          const FormatText &format,
                                                          const QString &text,
                                                          const bool numeration) {
    QString body = "<div class=Word>";
    QStringList doc_list = _doc.split(body);
    QString doc_f= doc_list.first();
    QString doc = doc_f +
            body +
            "\n<table style='margin-left:50in;'><tr style='height: 1pt;"
            "mso-height-rule:exactly'><td><div style='mso-element:" +
            disposition + "' id=";
    if (disposition == _header)  {
        doc+= "h1";
    }
    else {
        doc+= "f1";
    }
    doc += ">\n";
    if (numeration) {
        doc += "<p align='" + getAlign(format.align) + "'>\n" +
               "<span style='mso-field-code:PAGE'>\n</span>\n";
    }
    else doc += _getPropertiesText(text, format);
    doc += "</p>\n</div>&nbsp;\n</td>\n</tr></table>";
    _setLastDoc(doc, doc_list);
    _swapDoc(_doc, doc);
}

void DocumentWriter::Document::_swapDoc(QString &a, QString &b) {
    a.clear();
    a = b;
    b.clear();
}

void DocumentWriter::Document::_setLastDoc(QString &doc, QStringList &docList) {
    if (docList.count() > 1) {
        QString doc_l = docList.last();
        doc += doc_l;
    }
}

QString DocumentWriter::Document::_getPropertiesImage(const QString &base,
                                                      const QString &name,
                                                      const QString &align,
                                                      const QString &width,
                                                      const QString &height,
                                                      const QString &border) {
    return "<p style='margin-bottom: 0cm'>"
           "<img src='data:image/png;base64," + base +
           "' name='" + name +
           "' align='" + align +
           "' width='" + width +
           "' height='" + height +
           "' border='" + border +
           "'/></br></p>\n";
}
