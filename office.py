import sys
import os
import urllib.request
import subprocess
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton

class InstallationGUI(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        self.setStyleSheet("background-color: #222;")

        self.license_text = QLabel(
            """
            Custom Office Suite Installation
            This software is provided for free. If you are fine with this, click 'Agree and Proceed' to continue.
            """
        )
        self.license_text.setStyleSheet("font-size: 18px; color: #FFF; margin: 20px;")
        layout.addWidget(self.license_text)

        self.agree_button = QPushButton("Agree and Proceed")
        self.agree_button.setStyleSheet("background-color: #3498DB; color: #FFF; padding: 10px;")
        self.agree_button.clicked.connect(self.show_select_components)
        layout.addWidget(self.agree_button)

        self.setLayout(layout)

    def show_select_components(self):
        self.agree_button.setEnabled(False)
        self.agree_button.setText("License Agreement Accepted")
        self.agree_button.setStyleSheet("background-color: #27AE60; color: black; padding: 10px;")

        self.components_label = QLabel("Do you want Project and Visio components installed?")
        self.components_label.setStyleSheet("font-size: 16px; color: #FFF; margin: 20px;")
        self.layout().addWidget(self.components_label)

        self.yes_button = QPushButton("Yes")
        self.yes_button.setStyleSheet("background-color: #3498DB; color: #FFF; padding: 10px;")
        self.yes_button.clicked.connect(self.start_installation)
        self.layout().addWidget(self.yes_button)

        self.no_button = QPushButton("No")
        self.no_button.setStyleSheet("background-color: #E74C3C; color: #FFF; padding: 10px;")
        self.no_button.clicked.connect(self.start_installation_no)
        self.layout().addWidget(self.no_button)

    def start_installation(self):
        self.clear_layout()
        self.installation_label = QLabel("Installing Office...")
        self.installation_label.setStyleSheet("font-size: 16px; color: #FFF; margin: 20px;")
        self.layout().addWidget(self.installation_label)

        components = []
        if self.sender() == self.yes_button:
            components = ["Project", "Visio"]
            self.download_files("configuration-Office2021Enterprise.xml", "setup.exe")
            self.download_and_run_script("install.bat")
        elif self.sender() == self.no_button:
            self.download_files("configuration-Office2021novisio.xml", "setup.exe")

        self.complete_label = QLabel("Installation complete! Enjoy your Custom Office Suite.")
        self.complete_label.setStyleSheet("font-size: 16px; color: #FFF; margin: 20px;")
        self.layout().addWidget(self.complete_label)

    def start_installation_no(self):
        self.clear_layout()
        self.installation_label = QLabel("Installing Office...")
        self.installation_label.setStyleSheet("font-size: 16px; color: #FFF; margin: 20px;")
        self.layout().addWidget(self.installation_label)

        self.download_files("configuration-Office2021novisio.xml", "setup.exe")

        self.complete_label = QLabel("Now installing! Enjoy your Custom Office Suite.")
        self.complete_label.setStyleSheet("font-size: 16px; color: #FFF; margin: 20px;")
        self.layout().addWidget(self.complete_label)

    def clear_layout(self):
        for i in reversed(range(self.layout().count())):
            widget = self.layout().itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

    def download_files(self, config_file, setup_file):
        file_urls = {
            config_file: f"https://raw.githubusercontent.com/axeltechtips/office21/main/{config_file}",
            setup_file: f"https://raw.githubusercontent.com/axeltechtips/office21/main/{setup_file}"
        }

        for filename, url in file_urls.items():
            urllib.request.urlretrieve(url, filename)
            print(f"{filename} downloaded.")

    def download_and_run_script(self, script_file):
        script_url = f"https://raw.githubusercontent.com/axeltechtips/office21/main/{script_file}"
        script_path = os.path.join(os.getcwd(), script_file)
        
        urllib.request.urlretrieve(script_url, script_path)
        print(f"{script_file} downloaded.")
        
        subprocess.Popen([script_path], shell=True)

def main():
    app = QApplication(sys.argv)
    gui = InstallationGUI()
    gui.setWindowTitle("Custom Office Suite Installation")
    gui.setGeometry(100, 100, 600, 300)
    gui.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
