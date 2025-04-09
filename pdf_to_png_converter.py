import os
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QListWidget, QLabel, QFileDialog, QProgressBar,
                             QProgressDialog, QMessageBox, QFrame)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from pdf2image import convert_from_path


class DropListWidget(QListWidget):
    def __init__(self, converter):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.DragDrop)
        self.setSelectionMode(QListWidget.ExtendedSelection)
        self.converter = converter  # Referência direta ao conversor

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            links = []
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith('.pdf'):
                    links.append(path)

            self.converter.add_files(links)  # Usar a referência direta
        else:
            event.ignore()


class PDFtoPNGConverter(QMainWindow):
    def __init__(self):
        super().__init__()

        self.files_to_convert = []

        self.setWindowTitle("Conversor PDF para PNG")
        self.setGeometry(100, 100, 700, 500)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Layout principal
        main_layout = QVBoxLayout(central_widget)

        # Área de arrastar e soltar
        drop_area_label = QLabel("Arraste arquivos PDF aqui ou use o botão 'Selecionar PDFs'")
        drop_area_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(drop_area_label)

        # Lista de arquivos com suporte para arrastar e soltar
        self.files_list = DropListWidget(self)  # Passar self como referência
        main_layout.addWidget(self.files_list)

        # Frame de botões
        button_frame = QWidget()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setContentsMargins(0, 0, 0, 0)

        # Botões
        self.select_button = QPushButton("Selecionar PDFs")
        self.select_button.clicked.connect(self.select_files)
        button_layout.addWidget(self.select_button)

        self.convert_button = QPushButton("Converter para PNG")
        self.convert_button.clicked.connect(self.convert_files)
        button_layout.addWidget(self.convert_button)

        self.remove_button = QPushButton("Remover Selecionados")
        self.remove_button.clicked.connect(self.remove_selected_files)
        button_layout.addWidget(self.remove_button)

        self.clear_button = QPushButton("Limpar Tudo")
        self.clear_button.clicked.connect(self.clear_all_files)
        button_layout.addWidget(self.clear_button)

        main_layout.addWidget(button_frame)

        # Barra de status
        self.statusBar().showMessage("Pronto para selecionar arquivos PDF")

    def add_files(self, files):
        for file in files:
            if file not in self.files_to_convert:
                self.files_to_convert.append(file)
                self.files_list.addItem(os.path.basename(file))

        self.statusBar().showMessage(f"{len(self.files_to_convert)} arquivos prontos para conversão")

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecione os arquivos PDF",
            "",
            "Arquivos PDF (*.pdf)"
        )

        self.add_files(files)

    def convert_files(self):
        if not self.files_to_convert:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo selecionado para conversão")
            return

        progress = QProgressDialog("Convertendo arquivos...", "Cancelar", 0, len(self.files_to_convert), self)
        progress.setWindowTitle("Progresso da Conversão")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)

        converted_count = 0
        error_count = 0

        for i, pdf_file in enumerate(self.files_to_convert):
            if progress.wasCanceled():
                break

            progress.setLabelText(f"Convertendo: {os.path.basename(pdf_file)}")
            progress.setValue(i)

            try:
                # Converter PDF para imagens
                images = convert_from_path(pdf_file)

                # Salvar cada página como PNG
                pdf_dir = os.path.dirname(pdf_file)
                pdf_name = os.path.splitext(os.path.basename(pdf_file))[0]

                for j, image in enumerate(images):
                    image.save(os.path.join(pdf_dir, f"{pdf_name}_page_{j + 1}.png"), "PNG")

                converted_count += 1
                self.statusBar().showMessage(f"Convertido: {os.path.basename(pdf_file)}")
            except Exception as e:
                error_count += 1
                self.statusBar().showMessage(f"Erro ao converter {os.path.basename(pdf_file)}: {str(e)}")
                QMessageBox.critical(self, "Erro", f"Erro ao converter {os.path.basename(pdf_file)}:\n{str(e)}")

        progress.setValue(len(self.files_to_convert))

        QMessageBox.information(
            self,
            "Conversão Concluída",
            f"Conversão finalizada!\n\nArquivos convertidos: {converted_count}\nErros: {error_count}"
        )

        self.statusBar().showMessage(
            f"Conversão concluída! {converted_count} arquivos convertidos, {error_count} erros.")

    def remove_selected_files(self):
        selected_items = self.files_list.selectedItems()
        if not selected_items:
            return

        # Criar uma lista dos itens selecionados
        items_to_remove = []
        for item in selected_items:
            items_to_remove.append(item)

        # Remover os itens da lista
        for item in items_to_remove:
            row = self.files_list.row(item)
            self.files_list.takeItem(row)
            del self.files_to_convert[row]

        self.statusBar().showMessage(f"{len(self.files_to_convert)} arquivos prontos para conversão")

    def clear_all_files(self):
        self.files_to_convert.clear()
        self.files_list.clear()
        self.statusBar().showMessage("Lista de arquivos limpa")


# Iniciar a aplicação
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFtoPNGConverter()
    window.show()
    sys.exit(app.exec_())
