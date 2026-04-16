# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'frmMaintainDB.ui'
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
from PySide6.QtWidgets import (QApplication, QComboBox, QFormLayout, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QPushButton, QSizePolicy, QTableWidget, QTableWidgetItem,
    QTextEdit, QVBoxLayout, QWidget)

class Ui_frmMaintainDB(object):
    def setupUi(self, frmMaintainDB):
        if not frmMaintainDB.objectName():
            frmMaintainDB.setObjectName(u"frmMaintainDB")
        frmMaintainDB.resize(1080, 755)
        self.table1 = QTableWidget(frmMaintainDB)
        self.table1.setObjectName(u"table1")
        self.table1.setGeometry(QRect(6, 33, 847, 637))
        font = QFont()
        font.setPointSize(10)
        self.table1.setFont(font)
        self.groupBox_2 = QGroupBox(frmMaintainDB)
        self.groupBox_2.setObjectName(u"groupBox_2")
        self.groupBox_2.setGeometry(QRect(870, 170, 205, 337))
        font1 = QFont()
        font1.setFamilies([u"\u65b0\u7d30\u660e\u9ad4"])
        font1.setPointSize(11)
        self.groupBox_2.setFont(font1)
        self.horizontalLayout = QHBoxLayout(self.groupBox_2)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.lblTableName = QLabel(self.groupBox_2)
        self.lblTableName.setObjectName(u"lblTableName")
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lblTableName.sizePolicy().hasHeightForWidth())
        self.lblTableName.setSizePolicy(sizePolicy)
        self.lblTableName.setFont(font1)

        self.verticalLayout.addWidget(self.lblTableName)

        self.cboTable = QComboBox(self.groupBox_2)
        self.cboTable.setObjectName(u"cboTable")
        sizePolicy.setHeightForWidth(self.cboTable.sizePolicy().hasHeightForWidth())
        self.cboTable.setSizePolicy(sizePolicy)
        self.cboTable.setFont(font1)

        self.verticalLayout.addWidget(self.cboTable)

        self.lblTableName_2 = QLabel(self.groupBox_2)
        self.lblTableName_2.setObjectName(u"lblTableName_2")
        sizePolicy.setHeightForWidth(self.lblTableName_2.sizePolicy().hasHeightForWidth())
        self.lblTableName_2.setSizePolicy(sizePolicy)
        self.lblTableName_2.setFont(font1)

        self.verticalLayout.addWidget(self.lblTableName_2)

        self.ltxtkw_script = QLineEdit(self.groupBox_2)
        self.ltxtkw_script.setObjectName(u"ltxtkw_script")
        sizePolicy.setHeightForWidth(self.ltxtkw_script.sizePolicy().hasHeightForWidth())
        self.ltxtkw_script.setSizePolicy(sizePolicy)
        self.ltxtkw_script.setFont(font1)

        self.verticalLayout.addWidget(self.ltxtkw_script)

        self.btnQuery = QPushButton(self.groupBox_2)
        self.btnQuery.setObjectName(u"btnQuery")
        sizePolicy1 = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        sizePolicy1.setHorizontalStretch(0)
        sizePolicy1.setVerticalStretch(0)
        sizePolicy1.setHeightForWidth(self.btnQuery.sizePolicy().hasHeightForWidth())
        self.btnQuery.setSizePolicy(sizePolicy1)
        self.btnQuery.setFont(font1)

        self.verticalLayout.addWidget(self.btnQuery)

        self.btnInsert = QPushButton(self.groupBox_2)
        self.btnInsert.setObjectName(u"btnInsert")
        sizePolicy2 = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        sizePolicy2.setHorizontalStretch(0)
        sizePolicy2.setVerticalStretch(0)
        sizePolicy2.setHeightForWidth(self.btnInsert.sizePolicy().hasHeightForWidth())
        self.btnInsert.setSizePolicy(sizePolicy2)
        self.btnInsert.setFont(font1)

        self.verticalLayout.addWidget(self.btnInsert)

        self.btnDelete = QPushButton(self.groupBox_2)
        self.btnDelete.setObjectName(u"btnDelete")
        sizePolicy1.setHeightForWidth(self.btnDelete.sizePolicy().hasHeightForWidth())
        self.btnDelete.setSizePolicy(sizePolicy1)
        self.btnDelete.setFont(font1)

        self.verticalLayout.addWidget(self.btnDelete)

        self.btnUpdate = QPushButton(self.groupBox_2)
        self.btnUpdate.setObjectName(u"btnUpdate")
        sizePolicy1.setHeightForWidth(self.btnUpdate.sizePolicy().hasHeightForWidth())
        self.btnUpdate.setSizePolicy(sizePolicy1)
        self.btnUpdate.setFont(font1)

        self.verticalLayout.addWidget(self.btnUpdate)

        self.btnImport = QPushButton(self.groupBox_2)
        self.btnImport.setObjectName(u"btnImport")
        sizePolicy1.setHeightForWidth(self.btnImport.sizePolicy().hasHeightForWidth())
        self.btnImport.setSizePolicy(sizePolicy1)
        self.btnImport.setFont(font1)

        self.verticalLayout.addWidget(self.btnImport)

        self.btnExport = QPushButton(self.groupBox_2)
        self.btnExport.setObjectName(u"btnExport")
        sizePolicy1.setHeightForWidth(self.btnExport.sizePolicy().hasHeightForWidth())
        self.btnExport.setSizePolicy(sizePolicy1)
        self.btnExport.setFont(font1)

        self.verticalLayout.addWidget(self.btnExport)


        self.horizontalLayout.addLayout(self.verticalLayout)

        self.groupBox_1 = QGroupBox(frmMaintainDB)
        self.groupBox_1.setObjectName(u"groupBox_1")
        self.groupBox_1.setGeometry(QRect(870, 40, 205, 121))
        self.groupBox_1.setFont(font1)
        self.verticalLayout_4 = QVBoxLayout(self.groupBox_1)
        self.verticalLayout_4.setObjectName(u"verticalLayout_4")
        self.txtMsg = QTextEdit(self.groupBox_1)
        self.txtMsg.setObjectName(u"txtMsg")
        sizePolicy2.setHeightForWidth(self.txtMsg.sizePolicy().hasHeightForWidth())
        self.txtMsg.setSizePolicy(sizePolicy2)
        self.txtMsg.setMinimumSize(QSize(0, 0))
        self.txtMsg.setMaximumSize(QSize(16777215, 16777215))
        self.txtMsg.setFont(font1)

        self.verticalLayout_4.addWidget(self.txtMsg)

        self.groupBox_3 = QGroupBox(frmMaintainDB)
        self.groupBox_3.setObjectName(u"groupBox_3")
        self.groupBox_3.setGeometry(QRect(870, 515, 205, 187))
        sizePolicy3 = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Preferred)
        sizePolicy3.setHorizontalStretch(0)
        sizePolicy3.setVerticalStretch(0)
        sizePolicy3.setHeightForWidth(self.groupBox_3.sizePolicy().hasHeightForWidth())
        self.groupBox_3.setSizePolicy(sizePolicy3)
        self.groupBox_3.setFont(font1)
        self.verticalLayout_3 = QVBoxLayout(self.groupBox_3)
        self.verticalLayout_3.setObjectName(u"verticalLayout_3")
        self.verticalLayout_2 = QVBoxLayout()
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.formLayout = QFormLayout()
        self.formLayout.setObjectName(u"formLayout")
        self.lblUserName = QLabel(self.groupBox_3)
        self.lblUserName.setObjectName(u"lblUserName")
        sizePolicy3.setHeightForWidth(self.lblUserName.sizePolicy().hasHeightForWidth())
        self.lblUserName.setSizePolicy(sizePolicy3)
        self.lblUserName.setFont(font1)

        self.formLayout.setWidget(0, QFormLayout.LabelRole, self.lblUserName)

        self.ltxtUserName = QLineEdit(self.groupBox_3)
        self.ltxtUserName.setObjectName(u"ltxtUserName")
        sizePolicy4 = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        sizePolicy4.setHorizontalStretch(0)
        sizePolicy4.setVerticalStretch(0)
        sizePolicy4.setHeightForWidth(self.ltxtUserName.sizePolicy().hasHeightForWidth())
        self.ltxtUserName.setSizePolicy(sizePolicy4)
        self.ltxtUserName.setFont(font1)

        self.formLayout.setWidget(0, QFormLayout.FieldRole, self.ltxtUserName)

        self.lblPassword = QLabel(self.groupBox_3)
        self.lblPassword.setObjectName(u"lblPassword")
        sizePolicy3.setHeightForWidth(self.lblPassword.sizePolicy().hasHeightForWidth())
        self.lblPassword.setSizePolicy(sizePolicy3)
        self.lblPassword.setFont(font1)

        self.formLayout.setWidget(1, QFormLayout.LabelRole, self.lblPassword)

        self.ltxtPassword = QLineEdit(self.groupBox_3)
        self.ltxtPassword.setObjectName(u"ltxtPassword")
        sizePolicy4.setHeightForWidth(self.ltxtPassword.sizePolicy().hasHeightForWidth())
        self.ltxtPassword.setSizePolicy(sizePolicy4)
        self.ltxtPassword.setFont(font1)

        self.formLayout.setWidget(1, QFormLayout.FieldRole, self.ltxtPassword)

        self.lblUpScriptName = QLabel(self.groupBox_3)
        self.lblUpScriptName.setObjectName(u"lblUpScriptName")
        sizePolicy3.setHeightForWidth(self.lblUpScriptName.sizePolicy().hasHeightForWidth())
        self.lblUpScriptName.setSizePolicy(sizePolicy3)
        self.lblUpScriptName.setFont(font1)

        self.formLayout.setWidget(2, QFormLayout.LabelRole, self.lblUpScriptName)

        self.cboUpScript = QComboBox(self.groupBox_3)
        self.cboUpScript.setObjectName(u"cboUpScript")
        sizePolicy4.setHeightForWidth(self.cboUpScript.sizePolicy().hasHeightForWidth())
        self.cboUpScript.setSizePolicy(sizePolicy4)

        self.formLayout.setWidget(2, QFormLayout.FieldRole, self.cboUpScript)

        self.lblUpVersion = QLabel(self.groupBox_3)
        self.lblUpVersion.setObjectName(u"lblUpVersion")
        sizePolicy.setHeightForWidth(self.lblUpVersion.sizePolicy().hasHeightForWidth())
        self.lblUpVersion.setSizePolicy(sizePolicy)

        self.formLayout.setWidget(3, QFormLayout.LabelRole, self.lblUpVersion)

        self.ltxtUpVersion = QLineEdit(self.groupBox_3)
        self.ltxtUpVersion.setObjectName(u"ltxtUpVersion")
        sizePolicy4.setHeightForWidth(self.ltxtUpVersion.sizePolicy().hasHeightForWidth())
        self.ltxtUpVersion.setSizePolicy(sizePolicy4)

        self.formLayout.setWidget(3, QFormLayout.FieldRole, self.ltxtUpVersion)


        self.verticalLayout_2.addLayout(self.formLayout)

        self.btnUploadToDB = QPushButton(self.groupBox_3)
        self.btnUploadToDB.setObjectName(u"btnUploadToDB")
        sizePolicy4.setHeightForWidth(self.btnUploadToDB.sizePolicy().hasHeightForWidth())
        self.btnUploadToDB.setSizePolicy(sizePolicy4)
        self.btnUploadToDB.setFont(font1)

        self.verticalLayout_2.addWidget(self.btnUploadToDB)


        self.verticalLayout_3.addLayout(self.verticalLayout_2)


        self.retranslateUi(frmMaintainDB)

        QMetaObject.connectSlotsByName(frmMaintainDB)
    # setupUi

    def retranslateUi(self, frmMaintainDB):
        frmMaintainDB.setWindowTitle(QCoreApplication.translate("frmMaintainDB", u"\u8cc7\u6599\u7dad\u8b77", None))
        self.groupBox_2.setTitle(QCoreApplication.translate("frmMaintainDB", u"\u672c\u6a5f\u8cc7\u6599\u7dad\u8b77", None))
        self.lblTableName.setText(QCoreApplication.translate("frmMaintainDB", u"\u8cc7\u6599\u7dad\u8b77\u9805\u76ee:", None))
        self.lblTableName_2.setText(QCoreApplication.translate("frmMaintainDB", u"\u811a\u672c\u540d\u7a31:", None))
#if QT_CONFIG(tooltip)
        self.ltxtkw_script.setToolTip(QCoreApplication.translate("frmMaintainDB", u"\u67e5\u8a62\u6216\u66f4\u65b0\u8cc7\u6793\u5eab\u689d\u4ef6\u52a0\u5165\u811a\u672c\u540d\u7a31\u95dc\u9375\u5b57", None))
#endif // QT_CONFIG(tooltip)
        self.btnQuery.setText(QCoreApplication.translate("frmMaintainDB", u"\u67e5\u8a62", None))
        self.btnInsert.setText(QCoreApplication.translate("frmMaintainDB", u"\u65b0\u589e\u5217", None))
        self.btnDelete.setText(QCoreApplication.translate("frmMaintainDB", u"\u522a\u9664\u5217", None))
        self.btnUpdate.setText(QCoreApplication.translate("frmMaintainDB", u"\u66f4\u65b0\u8cc7\u6599\u5eab", None))
        self.btnImport.setText(QCoreApplication.translate("frmMaintainDB", u"\u8f09\u5165\u6a94\u6848", None))
        self.btnExport.setText(QCoreApplication.translate("frmMaintainDB", u"\u532f\u51fa\u6a94\u6848", None))
        self.groupBox_1.setTitle(QCoreApplication.translate("frmMaintainDB", u"\u8a0a\u606f\u986f\u793a", None))
        self.groupBox_3.setTitle(QCoreApplication.translate("frmMaintainDB", u"\u6e2c\u8a66\u811a\u672c\u4e0a\u50b3\u96f2\u7aef", None))
        self.lblUserName.setText(QCoreApplication.translate("frmMaintainDB", u"\u5e33\u865f:", None))
#if QT_CONFIG(tooltip)
        self.ltxtUserName.setToolTip(QCoreApplication.translate("frmMaintainDB", u"\u67e5\u8a62\u6216\u66f4\u65b0\u8cc7\u6793\u5eab\u689d\u4ef6\u52a0\u5165\u811a\u672c\u540d\u7a31\u95dc\u9375\u5b57", None))
#endif // QT_CONFIG(tooltip)
        self.lblPassword.setText(QCoreApplication.translate("frmMaintainDB", u"\u5bc6\u78bc:", None))
#if QT_CONFIG(tooltip)
        self.ltxtPassword.setToolTip(QCoreApplication.translate("frmMaintainDB", u"\u67e5\u8a62\u6216\u66f4\u65b0\u8cc7\u6793\u5eab\u689d\u4ef6\u52a0\u5165\u811a\u672c\u540d\u7a31\u95dc\u9375\u5b57", None))
#endif // QT_CONFIG(tooltip)
        self.lblUpScriptName.setText(QCoreApplication.translate("frmMaintainDB", u"\u811a\u672c:", None))
        self.lblUpVersion.setText(QCoreApplication.translate("frmMaintainDB", u"\u7248\u672c:", None))
        self.btnUploadToDB.setText(QCoreApplication.translate("frmMaintainDB", u"\u66f4\u65b0\u96f2\u7aef\u811a\u672c", None))
    # retranslateUi

