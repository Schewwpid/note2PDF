from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QPainter, QPageSize
from PyQt6.QtPrintSupport import QPrinter
from PyQt6.QtCore import QSize
from PyQt6.QtSvg import QSvgRenderer
from zipfile import ZipFile, BadZipFile
from lxml import etree
import plistlib
import sys, os
import nv_core
from nv_core import log

# 初始化 QApplication
app = QApplication(sys.argv)


class XmlWriterWithUID(plistlib._PlistWriter):
    def write_value(self, value):
        if isinstance(value, plistlib.UID):
            self.simple_element("string", f'UID:{value.data}')
        else:
            super().write_value(value)


# Override XML writer
plistlib._FORMATS[plistlib.FMT_XML]['writer'] = XmlWriterWithUID


def convertPlistToXml(plistData: bytes):
    """Convert binary plist data to XML string."""
    plist = plistlib.loads(plistData)
    return plistlib.dumps(plist, fmt=plistlib.FMT_XML).decode('utf-8')


def openFile(filePath: str):
    """Extract .note file, convert to SVG and return SVG content, width, and height."""
    log(f'Opening file: {filePath}')
    try:
        noteZipFile = ZipFile(filePath)
    except BadZipFile:
        log('Error: File is not a valid zip file')
        return None, None, None

    contentNames = noteZipFile.namelist()
    if not contentNames:
        log('Error: Zip file is empty')
        return None, None, None

    noteFolderName = contentNames[0].split('/')[0]
    log(f'Note folder name: {noteFolderName}')

    try:
        sessionList = noteZipFile.read(noteFolderName + '/Session.plist')
    except KeyError:
        log('Error: Session.plist not found')
        return None, None, None

    if not sessionList:
        log('Error: Session.plist is empty')
        return None, None, None

    svg, width, height = nv_core.convertPlistToSVG(plistlib.loads(sessionList))
    noteZipFile.close()
    return svg, width, height


def convertSvgToPdf(svg, width, height, outputPath):
    """Convert SVG to PDF using QPrinter and QSvgRenderer."""
    printer = QPrinter()
    printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
    printer.setOutputFileName(outputPath)
    printer.setPageSize(QPageSize(QSize(int(width), int(height))))
    printer.setResolution(300)

    # 使用 QPainter 和 QSvgRenderer 渲染 SVG 到 PDF
    svgData = etree.tostring(svg, xml_declaration=True, encoding="utf-8")
    svgRenderer = QSvgRenderer(svgData)

    # 确保在渲染之前初始化 QPainter，并在完成后正确关闭
    painter = QPainter()
    if not painter.begin(printer):
        log(f"Failed to open file {outputPath} for writing")
        return

    if svgRenderer.isValid():
        svgRenderer.render(painter)

    painter.end()
    log(f'PDF saved at {outputPath}')


def processDirectory(directory: str):
    """Convert all .note files in a directory to PDFs."""
    if not os.path.isdir(directory):
        log(f'Error: {directory} is not a valid directory')
        return

    for fileName in os.listdir(directory):
        if fileName.endswith('.note'):
            filePath = os.path.join(directory, fileName)
            log(f'Processing file: {filePath}')

            svg, width, height = openFile(filePath)
            if svg is None:
                log(f'Failed to process: {filePath}')
                continue

            outputPdfPath = os.path.splitext(filePath)[0] + '.pdf'
            convertSvgToPdf(svg, width, height, outputPdfPath)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python convert.py <directory_path>")
        sys.exit(1)

    directoryPath = sys.argv[1]
    processDirectory(directoryPath)
    print("Conversion complete.")
