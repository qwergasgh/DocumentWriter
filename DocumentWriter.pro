TARGET = DocumentWriter
TEMPLATE = lib

QT += core

win32 {
    QT += axcontainer
}

DEFINES += DOCUMENTWRITER_LIBRARY

DEFINES += QT_DEPRECATED_WARNINGS


SOURCES += \
        document.cpp \
        table.cpp

HEADERS += \
        document.h \
        documentwriter_global.h \ 
        format.h \
        getalign.h \
        table.h

unix {
    target.path = /usr/lib
    INSTALLS += target
}
