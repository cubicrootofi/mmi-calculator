from PyQt5.QtWidgets import (
    QMainWindow, QLabel, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem,
    QFileDialog, QApplication, QDialog, QTextEdit
)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import uic
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from typing import Optional
import xlsxwriter
import os
import sys
import qdarkstyle
import pandas as pd
import joblib


def load_model(file_path):
    return joblib.load(file_path)


def predict(model, x, R, q, lambda_):
    input_data = pd.DataFrame([[x, R, q, lambda_]], columns=['x', 'R', 'q', 'lambda'])
    y_pred = model.predict(input_data)
    return y_pred[0]


class MainApp(QMainWindow):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super(MainApp, self).__init__(parent)
        uic.loadUi(os.path.join(os.path.dirname(__file__), "inertia_calculator.ui"), self)
        self.setWindowTitle('Modified Moment of Inertia Calculator')
        self.setWindowIcon(QIcon(r"D:\PycharmProjects\pythonProject1\Eng Waleed\alpha.png"))

        # Connect actions
        self.actionExport_to_PDF.triggered.connect(self.export_to_pdf)
        self.actionExport_to_Excel.triggered.connect(self.export_to_excel)
        self.actionExit.triggered.connect(self.close)
        self.actionLicense.triggered.connect(self.show_popup)
        self.actionExport_to_HTML.triggered.connect(self.export_to_html)

        # Connect buttons
        self.pushButton.clicked.connect(self.inertia_calculator)
        self.pushButton_2.clicked.connect(self.clear_table)
        self.initialize_summary_table()

        # Create a permanent label
        self.permanent_label = QLabel("Yousef A. Sedik ©")
        font = QFont()
        font.setPointSize(10)
        self.permanent_label.setFont(font)
        self.statusBar().addPermanentWidget(self.permanent_label)

    def initialize_summary_table(self) -> None:
        """Initialize the summary table."""
        self.summary_table_frame = self.findChild(QWidget, "frame")
        self.summary_table_layout = QVBoxLayout(self.summary_table_frame)
        self.table = QTableWidget(self.summary_table_frame)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Alpha α", "L/dg", "q", "R", "λ"])
        self.summary_table_layout.addWidget(self.table)
        self.statusBar().setStyleSheet("color:#bfbfbf")

    def inertia_calculator(self) -> None:
        """Perform the inertia calculation and update the table."""
        try:
            model_file_path = r'_internal\gradient_boosting_model.pkl'
            model = load_model(model_file_path)

            # Extract and calculate the necessary values
            ipe_sections = {
                "W12X14": {
                    "area": 26.838,
                    "h": 302.26,
                    "b": 100.8,
                    "s": 5.08,
                    "t": 5.715
                },
                "IPE80": {
                    "area": 7.64,
                    "weight": 6,
                    "h": 80,
                    "b": 46,
                    "s": 3.8,
                    "r": 5,
                    "t": 5.2,
                    "c": 10.2,
                    "h-2c": 59,
                    "Ix": 80.1,
                    "Sx": 20,
                    "rx": 3.24,
                    "Iy": 8.49,
                    "Sy": 3.69,
                    "ry": 1.05
                },
                "IPE100": {
                    "area": 10.3,
                    "weight": 8.1,
                    "h": 100,
                    "b": 55,
                    "s": 4.1,
                    "r": 7,
                    "t": 5.7,
                    "c": 12.7,
                    "h-2c": 74,
                    "Ix": 171,
                    "Sx": 34.2,
                    "rx": 4.07,
                    "Iy": 15.9,
                    "Sy": 5.79,
                    "ry": 1.24
                },
                "IPE120": {
                    "area": 13.2,
                    "weight": 10.4,
                    "h": 120,
                    "b": 64,
                    "s": 4.4,
                    "r": 7,
                    "t": 6.3,
                    "c": 13.3,
                    "h-2c": 93,
                    "Ix": 318,
                    "Sx": 53,
                    "rx": 4.9,
                    "Iy": 27.7,
                    "Sy": 8.65,
                    "ry": 1.45
                },
                "IPE140": {
                    "area": 16.4,
                    "weight": 12.9,
                    "h": 140,
                    "b": 73,
                    "s": 4.7,
                    "r": 7,
                    "t": 6.9,
                    "c": 13.9,
                    "h-2c": 112,
                    "Ix": 541,
                    "Sx": 77.3,
                    "rx": 5.74,
                    "Iy": 44.9,
                    "Sy": 12.3,
                    "ry": 1.65
                },
                "IPE160": {
                    "area": 20.1,
                    "weight": 15.8,
                    "h": 160,
                    "b": 82,
                    "s": 5,
                    "r": 9,
                    "t": 7.4,
                    "c": 16.4,
                    "h-2c": 127,
                    "Ix": 869,
                    "Sx": 109,
                    "rx": 6.58,
                    "Iy": 68.3,
                    "Sy": 16.7,
                    "ry": 1.84
                },
                "IPE180": {
                    "area": 23.9,
                    "weight": 18.8,
                    "h": 180,
                    "b": 91,
                    "s": 5.3,
                    "r": 9,
                    "t": 8,
                    "c": 17,
                    "h-2c": 146,
                    "Ix": 1320,
                    "Sx": 146,
                    "rx": 7.42,
                    "Iy": 101,
                    "Sy": 22.2,
                    "ry": 2.05
                },
                "IPE200": {
                    "area": 28.5,
                    "weight": 22.4,
                    "h": 200,
                    "b": 100,
                    "s": 5.6,
                    "r": 12,
                    "t": 8.5,
                    "c": 20.5,
                    "h-2c": 159,
                    "Ix": 1940,
                    "Sx": 194,
                    "rx": 8.26,
                    "Iy": 142,
                    "Sy": 28.5,
                    "ry": 2.24
                },
                "IPE220": {
                    "area": 33.4,
                    "weight": 26.2,
                    "h": 220,
                    "b": 110,
                    "s": 5.9,
                    "r": 12,
                    "t": 9.2,
                    "c": 21.2,
                    "h-2c": 177,
                    "Ix": 2770,
                    "Sx": 252,
                    "rx": 9.11,
                    "Iy": 205,
                    "Sy": 37.3,
                    "ry": 2.48
                },
                "IPE240": {
                    "area": 39.1,
                    "weight": 30.7,
                    "h": 240,
                    "b": 120,
                    "s": 6.2,
                    "r": 15,
                    "t": 9.8,
                    "c": 24.8,
                    "h-2c": 190,
                    "Ix": 3890,
                    "Sx": 324,
                    "rx": 9.97,
                    "Iy": 284,
                    "Sy": 47.3,
                    "ry": 2.69
                },
                "IPE270": {
                    "area": 45.9,
                    "weight": 36.1,
                    "h": 270,
                    "b": 135,
                    "s": 6.6,
                    "r": 15,
                    "t": 10.2,
                    "c": 25.2,
                    "h-2c": 219,
                    "Ix": 5790,
                    "Sx": 429,
                    "rx": 11.2,
                    "Iy": 420,
                    "Sy": 62.2,
                    "ry": 3.02
                },
                "IPE300": {
                    "area": 53.8,
                    "weight": 42.2,
                    "h": 300,
                    "b": 150,
                    "s": 7.1,
                    "r": 15,
                    "t": 10.7,
                    "c": 25.7,
                    "h-2c": 248,
                    "Ix": 8360,
                    "Sx": 557,
                    "rx": 12.5,
                    "Iy": 604,
                    "Sy": 80.5,
                    "ry": 3.35
                },
                "IPE330": {
                    "area": 62.6,
                    "weight": 49.1,
                    "h": 330,
                    "b": 160,
                    "s": 7.5,
                    "r": 18,
                    "t": 11.5,
                    "c": 29.5,
                    "h-2c": 271,
                    "Ix": 11770,
                    "Sx": 713,
                    "rx": 13.7,
                    "Iy": 788,
                    "Sy": 98.5,
                    "ry": 3.55
                },
                "IPE360": {
                    "area": 72.7,
                    "weight": 57.1,
                    "h": 360,
                    "b": 170,
                    "s": 8,
                    "r": 18,
                    "t": 12.7,
                    "c": 30.7,
                    "h-2c": 298,
                    "Ix": 16270,
                    "Sx": 904,
                    "rx": 15,
                    "Iy": 1040,
                    "Sy": 123,
                    "ry": 3.79
                },
                "IPE400": {
                    "area": 84.5,
                    "weight": 66.3,
                    "h": 400,
                    "b": 180,
                    "s": 8.6,
                    "r": 21,
                    "t": 13.5,
                    "c": 34.5,
                    "h-2c": 331,
                    "Ix": 23130,
                    "Sx": 1160,
                    "rx": 16.5,
                    "Iy": 1320,
                    "Sy": 146,
                    "ry": 3.95
                },
                "IPE450": {
                    "area": 98.8,
                    "weight": 77.6,
                    "h": 450,
                    "b": 190,
                    "s": 9.4,
                    "r": 21,
                    "t": 14.6,
                    "c": 35.6,
                    "h-2c": 378,
                    "Ix": 33740,
                    "Sx": 1500,
                    "rx": 18.5,
                    "Iy": 1680,
                    "Sy": 176,
                    "ry": 4.12
                },
                "IPE500": {
                    "area": 116,
                    "weight": 90.7,
                    "h": 500,
                    "b": 200,
                    "s": 10.2,
                    "r": 21,
                    "t": 16,
                    "c": 37,
                    "h-2c": 426,
                    "Ix": 48200,
                    "Sx": 1930,
                    "rx": 20.4,
                    "Iy": 2140,
                    "Sy": 214,
                    "ry": 4.31
                },
                "IPE550": {
                    "area": 134,
                    "weight": 106,
                    "h": 550,
                    "b": 210,
                    "s": 11.1,
                    "r": 24,
                    "t": 17.2,
                    "c": 41.2,
                    "h-2c": 467,
                    "Ix": 67120,
                    "Sx": 2440,
                    "rx": 22.3,
                    "Iy": 2670,
                    "Sy": 254,
                    "ry": 4.45
                },
                "IPE600": {
                    "area": 156,
                    "weight": 122,
                    "h": 600,
                    "b": 220,
                    "s": 12,
                    "r": 24,
                    "t": 19,
                    "c": 43,
                    "h-2c": 514,
                    "Ix": 92080,
                    "Sx": 3070,
                    "rx": 24.3,
                    "Iy": 3390,
                    "Sy": 308,
                    "ry": 4.66
                }
            }
            length = float(self.lineEdit.text())
            opening_diameter = float(self.lineEdit_2.text())
            parent_section = self.comboBox_3.currentText()
            dimensions = ipe_sections[parent_section]
            R = float(self.comboBox_2.currentText())

            beam_height = dimensions["h"]
            dg = R * beam_height
            x = length / dg
            q = dg / opening_diameter

            if not (1.25 <= q <= 1.75):
                self.statusBar().showMessage(f"Warning! q is out of range: 1.25 ≤ q ≤ 1.75, q = {q:.3f}")
                return

            q = self.round_q(q)
            lambda_ = self.get_lambda(parent_section)
            y_predicted = predict(model, x, R, q, lambda_)

            # Insert values into the table
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(f"{y_predicted:.3f}"))
            self.table.setItem(row_position, 1, QTableWidgetItem(f"{x:.3f}"))
            self.table.setItem(row_position, 2, QTableWidgetItem(f"{q:.3f}"))
            self.table.setItem(row_position, 3, QTableWidgetItem(f"{R:.3f}"))
            self.table.setItem(row_position, 4, QTableWidgetItem(f"{lambda_:.4f}"))

        except Exception as e:
            self.statusBar().showMessage(f"Error: {e}")

    def round_q(self, q: float) -> float:
        """Round the value of q to the nearest acceptable value."""
        if 0 < q <= 1.25:
            return 1.25
        elif 1.25 < q <= 1.35:
            return 1.35
        elif 1.35 < q <= 1.45:
            return 1.45
        elif 1.45 < q <= 1.55:
            return 1.55
        elif 1.55 < q <= 1.65:
            return 1.65
        elif 1.65 < q <= 1.75:
            return 1.75
        return q

    def get_lambda(self, parent_section: str) -> float:
        """Get lambda value based on the parent section."""
        section_lambdas = {
            "IPE80": 0.4878, "IPE100": 0.4878,
            "IPE120": 0.5238, "IPE140": 0.5238, "IPE160": 0.5238, "IPE180": 0.5238, "IPE200": 0.5238,
            "IPE220": 0.5499, "IPE240": 0.5499, "IPE270": 0.5499, "IPE300": 0.5499,
            "IPE330": 0.5857, "IPE360": 0.5857, "IPE400": 0.5857,
            "IPE450": 0.6789, "IPE500": 0.6789,
            "IPE550": 0.7378, "IPE600": 0.7378
        }
        return section_lambdas.get(parent_section, 0.0)

    def clear_table(self) -> None:
        """Clear the summary table."""
        self.table.setRowCount(0)

    def export_to_pdf(self) -> None:
        """Export the table to a PDF file."""
        try:
            file_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF files (*.pdf)")
            if file_path:
                pdf = SimpleDocTemplate(file_path, pagesize=letter)
                table_data = [["L/dg", "q", "R", "λ", "Alpha α"]]

                for row in range(self.table.rowCount()):
                    row_data = [
                        self.table.item(row, column).text() if self.table.item(row, column) else ""
                        for column in range(self.table.columnCount())
                    ]
                    table_data.append(row_data)

                table = Table(table_data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                pdf.build([table])
                self.statusBar().showMessage(f"Exported to PDF: {file_path}")
        except Exception as e:
            self.statusBar().showMessage(f"Error exporting to PDF: {e}")

    def export_to_excel(self) -> None:
        """Export the table to an Excel file with styling."""
        try:
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel files (*.xlsx)")
            if file_path:
                workbook = xlsxwriter.Workbook(file_path)
                worksheet = workbook.add_worksheet()

                # Define some formats
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#f4cccc', 'border': 1, 'align': 'center'
                })
                center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

                # Write headers with styling
                headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
                for col, header in enumerate(headers):
                    worksheet.write(0, col, header, header_format)

                # Write data with styling
                for row in range(self.table.rowCount()):
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        value = item.text() if item else ''
                        worksheet.write(row + 1, col, value, center_format)

                # Set column width
                for col in range(self.table.columnCount()):
                    worksheet.set_column(col, col, 15)

                workbook.close()
                self.statusBar().showMessage(f"Exported to Excel: {file_path}")
        except Exception as e:
            self.statusBar().showMessage(f"Error exporting to Excel: {e}")

    def export_to_html(self) -> None:
        """Export the table to an HTML file with styling."""
        try:
            file_path, _ = QFileDialog.getSaveFileName(self, "Save HTML", "", "HTML files (*.html)")
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write("<html>\n<head>\n")
                    file.write("<style>\n")
                    file.write("table { border-collapse: collapse; width: 100%; }\n")
                    file.write("th, td { border: 1px solid #ddd; padding: 8px; }\n")
                    file.write("th { background-color: #f2f2f2; color: #333; text-align: center; }\n")
                    file.write("tr:nth-child(even) { background-color: #f9f9f9; }\n")
                    file.write("tr:hover { background-color: #f1f1f1; }\n")
                    file.write("</style>\n</head>\n<body>\n")
                    file.write("<h1>Table Export</h1>\n")
                    file.write("<table>\n")

                    # Write table headers
                    file.write("<thead>\n<tr>\n")
                    headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
                    for header in headers:
                        file.write(f"<th>{header}</th>\n")
                    file.write("</tr>\n</thead>\n")

                    # Write table data
                    file.write("<tbody>\n")
                    for row in range(self.table.rowCount()):
                        file.write("<tr>\n")
                        for col in range(self.table.columnCount()):
                            item = self.table.item(row, col)
                            value = item.text() if item else ''
                            file.write(f"<td>{value}</td>\n")
                        file.write("</tr>\n")
                    file.write("</tbody>\n</table>\n")
                    file.write("</body>\n</html>\n")

                self.statusBar().showMessage(f"Exported to HTML: {file_path}")
        except Exception as e:
            self.statusBar().showMessage(f"Error exporting to HTML: {e}")

    def show_popup(self) -> None:
        """Show a license popup."""
        dialog = QDialog(self)
        dialog.setWindowTitle("License")
        layout = QVBoxLayout(dialog)
        text_edit = QTextEdit(dialog)
        text_edit.setText("""Software License Agreement

1. License Grant

This License Agreement (the "Agreement") is made and entered into by and between Yousef Sedik ("Author") and Eng. Waleed Mohamed Ahmed ("Collaborator") (collectively referred to as "Licensors") and the licensee ("Licensee") who obtains the software application ("Software"). The Software is a proprietary application designed to calculate the modified moment of inertia.

2. Use of Software

The Licensee is granted a non-exclusive, non-transferable license to use the Software solely for the Licensee's internal engineering purposes. The Software may not be used for any commercial, governmental, or other purposes without the explicit prior written consent of the Licensors.

3. License Fees

The Licensee agrees to pay the Licensors the required fee for using the Software as determined by the Licensors. Use of the Software is strictly prohibited without payment of the applicable fee.

4. Restrictions

The Licensee may not:

Copy, modify, or distribute the Software without prior written consent from the Licensors.
Reverse engineer, decompile, disassemble, or otherwise attempt to discover the source code of the Software.
Use the Software in any manner that violates the terms of this Agreement.

5. Ownership

The Software and any associated documentation, including any updates, are the intellectual property of the Licensors. The Licensee does not acquire any ownership rights in the Software by virtue of this Agreement.

6. Termination

The Licensors reserve the right to terminate this Agreement and the License granted hereunder if the Licensee breaches any terms of this Agreement. Upon termination, the Licensee must cease all use of the Software and destroy all copies in their possession.

7. Warranty Disclaimer

The Software is provided "as-is" without any warranties, express or implied. The Licensors make no representations or warranties regarding the Software's performance or its fitness for a particular purpose.

8. Limitation of Liability

In no event shall the Licensors be liable for any damages arising from the use or inability to use the Software, including but not limited to incidental or consequential damages.

9. Governing Law

This Agreement shall be governed by and construed in accordance with the laws of Egypt, without regard to its conflict of laws principles.

10. Entire Agreement

This Agreement constitutes the entire agreement between the Licensors and the Licensee concerning the Software and supersedes all prior agreements or understandings, whether written or oral, relating to the Software.

By using the Software, the Licensee acknowledges that they have read, understood, and agree to be bound by the terms and conditions of this Agreement.

Licensors:

Yousef Sedik
Email: yousefsedik.bus@gmail.com

Eng. Waleed Mohamed Ahmed
Email: waleed.sayed@eng.asu.edu.eg""")
        text_edit.setReadOnly(True)
        layout.addWidget(text_edit)
        dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    mainWindow = MainApp()
    mainWindow.show()
    sys.exit(app.exec_())
