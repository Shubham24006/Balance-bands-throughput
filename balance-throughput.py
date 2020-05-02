import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox


class UIDialog(object):
    """
    class to generate the dialog.
    """

    def __init__(self):
        self.label = QtWidgets.QLabel(Dialog)
        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_2 = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_3 = QtWidgets.QLineEdit(Dialog)
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton_2 = QtWidgets.QPushButton(Dialog)
        self.pushButton_3 = QtWidgets.QPushButton(Dialog)

    def message_box(self, title, text, icon):
        message_box = QMessageBox()
        message_box.setWindowTitle(title)
        message_box.setText(text)
        message_box.setIcon(icon)
        message_box.exec_()

    def browse_file(self):
        filename, _ = QtWidgets.QFileDialog.getOpenFileName()
        self.lineEdit.setText(filename)

    def choose_folder(self):
        directory = QtWidgets.QFileDialog.getExistingDirectory()
        self.lineEdit_2.setText(directory)

    def on_submit(self):
        filename = self.lineEdit_3.text()
        source_file_path = self.lineEdit.text()
        destination_file_path = self.lineEdit_2.text() + "/" + filename + ".xlsx"

        if not filename:
            self.message_box(title="ERROR", text="Enter the File Name", icon=QMessageBox.Warning)
            return

        try:

            df = pd.read_excel(source_file_path)

            df_rule = pd.read_excel(source_file_path, sheet_name="Sheet2")

            source_rule = list(map(int, df_rule.iloc[0, 0].split(",")))

            bands_list = df_rule.iloc[0, 1:6]
            no_of_bands = len(bands_list)
            bands = df.iloc[0, 0:no_of_bands]

            rules_dict = {band: [] for band in bands}

            for key, i in zip(rules_dict, bands_list):
                rules_dict[key] = list(map(int, i.split(",")))

            bands_dict = {band: [] for band in bands}
            sources_dict = {band: [] for band in bands}

            # drop the first row in the dataframe
            df = df.drop([0])

            sources = {}
            targets = {}

            sources_id = []
            data_frames = []

            # first we find all our sources and targets in each row
            for row in df.itertuples():
                for index, key in enumerate(sources_dict, 1):

                    if row[index + no_of_bands] < source_rule[0] and row[index + 2 * no_of_bands] > source_rule[1]:
                        sources_dict[key].append(row[index])

                        if f's{row[0]}' in sources:
                            sources[f's{row[0]}'].append(row[index])
                        else:
                            sources[f's{row[0]}'] = [row[index]]

                if sources.get(f's{row[0]}'):
                    _ = []
                    for index, key in enumerate(sources_dict, 1):
                        if row[index + no_of_bands] > rules_dict[key][0] and row[index + 2 * no_of_bands] < \
                                rules_dict[key][1]:
                            _.append({key: row[index]})
                        targets[f't{row[0]}'] = _

            # seperate targets with particular source
            for (source, target) in zip(sources.items(), targets.items()):

                d = {k: v for d in target[1] for k, v in d.items()}

                for val in source[1]:
                    sources_id.append(val)
                    for band, band_list in bands_dict.items():
                        band_list.append(d.get(band, np.nan))

            # create data_frames for each cell
            for key, value in sources_dict.items():
                lists = [[] for _ in range(no_of_bands)]

                for val in value:
                    index = sources_id.index(val)
                    for i, ke in enumerate(rules_dict):
                        lists[i].append(bands_dict[ke][index])

                rdf = {str(key) + " CELL_ID": value}
                for e, band in enumerate(bands):
                    if band != key:
                        rdf[band] = lists[e]

                data_frames.append(pd.DataFrame(rdf))

            # To write in the sheets
            writer = pd.ExcelWriter(destination_file_path, engine='xlsxwriter')
            c = 1
            for data_frame in data_frames:
                data_frame.to_excel(writer, sheet_name=f'Sheet{c}', index=False)
                c += 1
            writer.save()

        except Exception as e:
            self.message_box(title="ERROR", text=str(e), icon=QMessageBox.Critical)

        else:
            self.message_box(title="SUCCESS", text="Done Sucessfully", icon=QMessageBox.Information)

    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(595, 416)

        self.pushButton.setGeometry(QtCore.QRect(410, 70, 89, 25))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.browse_file)

        self.lineEdit.setGeometry(QtCore.QRect(42, 70, 341, 25))
        self.lineEdit.setObjectName("lineEdit")

        self.pushButton_2.setGeometry(QtCore.QRect(410, 160, 100, 25))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.choose_folder)

        self.lineEdit_2.setGeometry(QtCore.QRect(40, 160, 341, 25))
        self.lineEdit_2.setObjectName("lineEdit_2")

        self.pushButton_3.setGeometry(QtCore.QRect(40, 310, 89, 25))
        self.pushButton_3.setObjectName("pushButton_3")

        self.pushButton_3.clicked.connect(self.on_submit)

        self.label.setGeometry(QtCore.QRect(40, 226, 201, 31))
        self.label.setObjectName("label")

        self.lineEdit_3.setGeometry(QtCore.QRect(190, 230, 211, 25))
        self.lineEdit_3.setObjectName("lineEdit_3")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Balance-Throughput"))
        self.pushButton.setText(_translate("Dialog", "Browse file"))
        self.pushButton_2.setText(_translate("Dialog", "select Folder"))
        self.pushButton_3.setText(_translate("Dialog", "submit"))
        self.label.setText(_translate("Dialog", "Enter File Name"))


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = UIDialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
