#ifndef DOCUMENTWRITER_GLOBAL_H
#define DOCUMENTWRITER_GLOBAL_H

#include <QtCore/qglobal.h>

#if defined(DOCUMENTWRITER_LIBRARY)
#  define DOCUMENTWRITERSHARED_EXPORT Q_DECL_EXPORT
#else
#  define DOCUMENTWRITERSHARED_EXPORT Q_DECL_IMPORT
#endif

#endif // DOCUMENTWRITER_GLOBAL_H
