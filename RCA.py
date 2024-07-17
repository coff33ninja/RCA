import subprocess
import json
import threading
import time
from wakeonlan import send_magic_packet
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
)
from winrm.protocol import Protocol

ADDRESS_BOOK_FILE = "device_address_book.json"


class Device:
    def __init__(
        self,
        nickname,
        ip_address,
        username="",
        password="",
        mac_address="",
        status="offline",
    ):
        self.nickname = nickname
        self.ip_address = ip_address
        self.username = username
        self.password = password
        self.mac_address = mac_address
        self.status = status


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Remote Control")
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.nickname_input = QLineEdit(self)
        self.nickname_input.setPlaceholderText("Enter Nickname")

        self.ip_input = QLineEdit(self)
        self.ip_input.setPlaceholderText("Enter IP Address")
        self.ip_input.returnPressed.connect(self.add_device)

        self.username_input = QLineEdit(self)
        self.username_input.setPlaceholderText("Username (optional)")

        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("Password (optional)")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)

        self.mac_input = QLineEdit(self)
        self.mac_input.setPlaceholderText("MAC Address (optional)")

        self.add_button = QPushButton("Add Device", self)
        self.add_button.clicked.connect(self.add_device)

        self.edit_button = QPushButton("Edit Device", self)
        self.edit_button.clicked.connect(self.edit_device)

        self.remove_button = QPushButton("Remove Device", self)
        self.remove_button.clicked.connect(self.remove_device)

        self.devices_list = QListWidget(self)
        self.devices_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.nickname_input)
        main_layout.addWidget(self.ip_input)

        credentials_layout = QHBoxLayout()
        credentials_layout.addWidget(self.username_input)
        credentials_layout.addWidget(self.password_input)
        credentials_layout.addWidget(self.mac_input)
        main_layout.addLayout(credentials_layout)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.remove_button)
        main_layout.addLayout(button_layout)

        main_layout.addWidget(self.devices_list)

        self.reboot_button = QPushButton("Reboot Device", self)
        self.reboot_button.clicked.connect(
            lambda: self.execute_command("Restart-Computer -Force")
        )
        self.shutdown_button = QPushButton("Shutdown Device", self)
        self.shutdown_button.clicked.connect(
            lambda: self.execute_command("Stop-Computer -Force")
        )
        self.wol_button = QPushButton("Wake-on-LAN Device", self)
        self.wol_button.clicked.connect(self.send_wol)

        actions_layout = QHBoxLayout()
        actions_layout.addWidget(self.reboot_button)
        actions_layout.addWidget(self.shutdown_button)
        actions_layout.addWidget(self.wol_button)
        main_layout.addLayout(actions_layout)

        self.central_widget.setLayout(main_layout)

        self.devices = []
        self.load_devices()

        self.ping_thread = threading.Thread(target=self.ping_devices, daemon=True)
        self.ping_thread.start()

    def add_device(self):
        nickname = self.nickname_input.text().strip()
        ip_address = self.ip_input.text().strip()
        username = self.username_input.text().strip()
        password = self.password_input.text().strip()
        mac_address = self.mac_input.text().strip()

        if nickname and ip_address:
            device = Device(nickname, ip_address, username, password, mac_address)
            self.devices.append(device)
            self.nickname_input.clear()
            self.ip_input.clear()
            self.username_input.clear()
            self.password_input.clear()
            self.mac_input.clear()
            self.update_device_list()
            self.save_devices()
        else:
            QMessageBox.warning(
                self, "Input Error", "Nickname and IP Address are required"
            )

    def edit_device(self):
        selected_items = self.devices_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "No device selected")
            return

        selected_index = self.devices_list.row(selected_items[0])
        device = self.devices[selected_index]

        self.nickname_input.setText(device.nickname)
        self.ip_input.setText(device.ip_address)
        self.username_input.setText(device.username)
        self.password_input.setText(device.password)
        self.mac_input.setText(device.mac_address)

        self.add_button.setText("Update Device")
        self.add_button.clicked.disconnect()
        self.add_button.clicked.connect(lambda: self.update_device(selected_index))

    def update_device(self, index):
        nickname = self.nickname_input.text().strip()
        ip_address = self.ip_input.text().strip()
        username = self.username_input.text().strip()
        password = self.password_input.text().strip()
        mac_address = self.mac_input.text().strip()

        if nickname and ip_address:
            self.devices[index].nickname = nickname
            self.devices[index].ip_address = ip_address
            self.devices[index].username = username
            self.devices[index].password = password
            self.devices[index].mac_address = mac_address
            self.update_device_list()
            self.save_devices()

            self.nickname_input.clear()
            self.ip_input.clear()
            self.username_input.clear()
            self.password_input.clear()
            self.mac_input.clear()
            self.add_button.setText("Add Device")
            self.add_button.clicked.disconnect()
            self.add_button.clicked.connect(self.add_device)
        else:
            QMessageBox.warning(
                self, "Input Error", "Nickname and IP Address are required"
            )

    def remove_device(self):
        selected_items = self.devices_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "No device selected")
            return

        selected_index = self.devices_list.row(selected_items[0])
        del self.devices[selected_index]
        self.update_device_list()
        self.save_devices()

    def update_device_list(self):
        self.devices_list.clear()
        for device in self.devices:
            item = QListWidgetItem(
                f"{device.nickname} ({device.ip_address}) - {device.status}"
            )
            self.devices_list.addItem(item)

    def execute_command(self, action):
        selected_items = self.devices_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "No device selected")
            return

        selected_index = self.devices_list.row(selected_items[0])
        device = self.devices[selected_index]

        try:
            session = Protocol(
                endpoint=f"http://{device.ip_address}:5985/wsman",
                transport="ntlm",
                username=device.username,
                password=device.password,
            )
            shell_id = session.open_shell()
            command_id = session.run_command(shell_id, f"powershell -Command {action}")
            std_out, std_err, status_code = session.get_command_output(
                shell_id, command_id
            )
            print(f"Output: {std_out}")
            print(f"Error: {std_err}")
            session.cleanup_command(shell_id, command_id)
            session.close_shell(shell_id)
            if status_code == 0:
                QMessageBox.information(
                    self,
                    "Success",
                    f"{action.split('-')[0]} command executed successfully on {device.nickname}",
                )
            else:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Failed to execute {action} on {device.nickname}: {std_err.decode()}",
                )
        except Exception as e:
            print(f"Failed to execute {action} on {device.nickname}: {e}")
            QMessageBox.critical(
                self, "Error", f"Failed to execute {action} on {device.nickname}: {e}"
            )

    def send_wol(self):
        selected_items = self.devices_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "No device selected")
            return

        selected_index = self.devices_list.row(selected_items[0])
        device = self.devices[selected_index]

        try:
            mac_formatted = device.mac_address.replace(":", "").replace("-", "")
            send_magic_packet(mac_formatted)
            QMessageBox.information(
                self, "WOL Sent", f"Wake-on-LAN packet sent to {device.nickname}"
            )
        except Exception as e:
            print(f"Failed to send WOL to {device.nickname}: {e}")
            QMessageBox.critical(
                self, "Error", f"Failed to send WOL to {device.nickname}: {e}"
            )

    def save_devices(self):
        with open(ADDRESS_BOOK_FILE, "w") as f:
            json.dump([device.__dict__ for device in self.devices], f)

    def load_devices(self):
        try:
            with open(ADDRESS_BOOK_FILE, "r") as f:
                device_data_list = json.load(f)
                for device_data in device_data_list:
                    device = Device(
                        nickname=device_data.get("nickname", "Unnamed Device"),
                        ip_address=device_data["ip_address"],
                        username=device_data.get("username", ""),
                        password=device_data.get("password", ""),
                        mac_address=device_data.get("mac_address", ""),
                        status=device_data.get("status", "offline"),
                    )
                    self.devices.append(device)
                self.update_device_list()
        except FileNotFoundError:
            pass

    def ping_devices(self):
        while True:
            for device in self.devices:
                response = subprocess.run(
                    ["ping", "-n", "1", device.ip_address], stdout=subprocess.PIPE
                )
                device.status = "online" if response.returncode == 0 else "offline"
            self.update_device_list()
            time.sleep(5)

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
