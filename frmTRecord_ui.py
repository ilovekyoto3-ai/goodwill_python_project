# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'frmTRecord.ui'
##
## Created by: Qt User Interface Compiler version 6.7.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QHeaderView, QSizePolicy, QTableView,
    QWidget)

class Ui_frmTRecord(object):
    def setupUi(self, frmTRecord):
        if not frmTRecord.objectName():
            frmTRecord.setObjectName(u"frmTRecord")
        frmTRecord.resize(1280, 720)
        self.tbview = QTableView(frmTRecord)
        self.tbview.setObjectName(u"tbview")
        self.tbview.setGeometry(QRect(12, 30, 1261, 679))
        self.tbview.setAutoScroll(False)

        self.retranslateUi(frmTRecord)

        QMetaObject.connectSlotsByName(frmTRecord)
    # setupUi

    def retranslateUi(self, frmTRecord):
        frmTRecord.setWindowTitle(QCoreApplication.translate("frmTRecord", u"Form", None))
    # retranslateUi

