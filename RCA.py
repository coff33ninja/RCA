import subprocess
import json
import threading
import time
from wakeonlan import send_magic_packet
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.popup import Popup
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.recycleview import RecycleView
from kivy.uix.recycleview.views import RecycleDataViewBehavior
from kivy.uix.recycleboxlayout import RecycleBoxLayout
from kivy.uix.recycleview.layout import LayoutSelectionBehavior
from kivy.uix.behaviors import FocusBehavior
from kivy.properties import BooleanProperty, ObjectProperty
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


class SelectableRecycleBoxLayout(
    FocusBehavior, LayoutSelectionBehavior, RecycleBoxLayout
):
    """Adds selection and focus behaviour to the view."""


class SelectableLabel(RecycleDataViewBehavior, Label):
    """Add selection support to the Label"""

    index = None
    selected = BooleanProperty(False)
    selectable = BooleanProperty(True)

    def refresh_view_attrs(self, rv, index, data):
        """Catch and handle the view changes"""
        self.index = index
        return super(SelectableLabel, self).refresh_view_attrs(rv, index, data)

    def on_touch_down(self, touch):
        """Add selection on touch down"""
        if super(SelectableLabel, self).on_touch_down(touch):
            return True
        if self.collide_point(*touch.pos) and self.selectable:
            return self.parent.select_with_touch(self.index, touch)

    def apply_selection(self, rv, index, is_selected):
        """Respond to the selection of items in the view."""
        self.selected = is_selected
        if is_selected:
            rv.selected_index = index
        else:
            rv.selected_index = None


class RV(RecycleView):
    selected_index = ObjectProperty(None)

    def __init__(self, **kwargs):
        super(RV, self).__init__(**kwargs)
        self.data = []

        # Set the layout manager to SelectableRecycleBoxLayout
        self.layout_manager = SelectableRecycleBoxLayout(
            default_size=(None, 48), size_hint=(1, None), orientation="vertical"
        )


class MainWindow(FloatLayout):
    def __init__(self, **kwargs):
        super(MainWindow, self).__init__(**kwargs)

        self.nickname_input = TextInput(
            hint_text="Enter Nickname",
            size_hint=(0.4, 0.05),
            pos_hint={"x": 0.05, "top": 0.95},
            font_size=16,
        )
        self.add_widget(self.nickname_input)

        self.ip_input = TextInput(
            hint_text="Enter IP Address",
            size_hint=(0.4, 0.05),
            pos_hint={"x": 0.05, "top": 0.88},
            font_size=16,
        )
        self.add_widget(self.ip_input)

        self.username_input = TextInput(
            hint_text="Username (optional)",
            size_hint=(0.4, 0.05),
            pos_hint={"x": 0.05, "top": 0.81},
            font_size=16,
        )
        self.add_widget(self.username_input)

        self.password_input = TextInput(
            hint_text="Password (optional)",
            size_hint=(0.4, 0.05),
            pos_hint={"x": 0.05, "top": 0.74},
            font_size=16,
            password=True,
        )
        self.add_widget(self.password_input)

        self.mac_input = TextInput(
            hint_text="MAC Address (optional)",
            size_hint=(0.4, 0.05),
            pos_hint={"x": 0.05, "top": 0.67},
            font_size=16,
        )
        self.add_widget(self.mac_input)

        self.add_button = Button(
            text="Add Device",
            size_hint=(0.4, 0.05),
            pos_hint={"x": 0.05, "top": 0.6},
            font_size=16,
        )
        self.add_button.bind(on_press=self.add_device)
        self.add_widget(self.add_button)

        self.edit_button = Button(
            text="Edit Device",
            size_hint=(0.4, 0.05),
            pos_hint={"x": 0.05, "top": 0.53},
            font_size=16,
        )
        self.edit_button.bind(on_press=self.edit_device)
        self.add_widget(self.edit_button)

        self.remove_button = Button(
            text="Remove Device",
            size_hint=(0.4, 0.05),
            pos_hint={"x": 0.05, "top": 0.46},
            font_size=16,
        )
        self.remove_button.bind(on_press=self.remove_device)
        self.add_widget(self.remove_button)

        self.rv = RV(size_hint=(0.4, 0.35), pos_hint={"x": 0.55, "top": 0.95})
        self.rv.viewclass = "SelectableLabel"
        self.add_widget(self.rv)

        self.reboot_button = Button(
            text="Reboot Device",
            size_hint=(0.25, 0.05),
            pos_hint={"x": 0.05, "top": 0.39},
            font_size=16,
        )
        self.reboot_button.bind(
            on_press=lambda x: self.execute_command("Restart-Computer -Force")
        )
        self.add_widget(self.reboot_button)

        self.shutdown_button = Button(
            text="Shutdown Device",
            size_hint=(0.25, 0.05),
            pos_hint={"x": 0.35, "top": 0.39},
            font_size=16,
        )
        self.shutdown_button.bind(
            on_press=lambda x: self.execute_command("Stop-Computer -Force")
        )
        self.add_widget(self.shutdown_button)

        self.wol_button = Button(
            text="Wake-on-LAN Device",
            size_hint=(0.25, 0.05),
            pos_hint={"x": 0.65, "top": 0.39},
            font_size=16,
        )
        self.wol_button.bind(on_press=self.send_wol)
        self.add_widget(self.wol_button)

        self.devices = []
        self.load_devices()

        self.ping_thread = threading.Thread(target=self.ping_devices, daemon=True)
        self.ping_thread.start()

    def add_device(self, instance):
        nickname = self.nickname_input.text.strip()
        ip_address = self.ip_input.text.strip()
        username = self.username_input.text.strip()
        password = self.password_input.text.strip()
        mac_address = self.mac_input.text.strip()

        if nickname and ip_address:
            device = Device(nickname, ip_address, username, password, mac_address)
            self.devices.append(device)
            self.nickname_input.text = ""
            self.ip_input.text = ""
            self.username_input.text = ""
            self.password_input.text = ""
            self.mac_input.text = ""
            self.update_device_list()
            self.save_devices()
        else:
            self.show_popup("Input Error", "Nickname and IP Address are required")

    def edit_device(self, instance):
        selected_index = self.rv.selected_index
        if selected_index is None:
            self.show_popup("Selection Error", "No device selected")
            return

        device = self.devices[selected_index]

        self.nickname_input.text = device.nickname
        self.ip_input.text = device.ip_address
        self.username_input.text = device.username
        self.password_input.text = device.password
        self.mac_input.text = device.mac_address

        self.add_button.text = "Update Device"
        self.add_button.unbind(on_press=self.add_device)
        self.add_button.bind(on_press=lambda x: self.update_device(selected_index))

    def update_device(self, index):
        nickname = self.nickname_input.text.strip()
        ip_address = self.ip_input.text.strip()
        username = self.username_input.text.strip()
        password = self.password_input.text.strip()
        mac_address = self.mac_input.text.strip()

        if nickname and ip_address:
            self.devices[index].nickname = nickname
            self.devices[index].ip_address = ip_address
            self.devices[index].username = username
            self.devices[index].password = password
            self.devices[index].mac_address = mac_address
            self.update_device_list()
            self.save_devices()

            self.nickname_input.text = ""
            self.ip_input.text = ""
            self.username_input.text = ""
            self.password_input.text = ""
            self.mac_input.text = ""
            self.add_button.text = "Add Device"
            self.add_button.unbind(on_press=self.update_device)
            self.add_button.bind(on_press=self.add_device)
        else:
            self.show_popup("Input Error", "Nickname and IP Address are required")

    def remove_device(self, instance):
        selected_index = self.rv.selected_index
        if selected_index is None:
            self.show_popup("Selection Error", "No device selected")
            return

        del self.devices[selected_index]
        self.update_device_list()
        self.save_devices()

    def update_device_list(self):
        self.rv.data = [{"text": device.nickname} for device in self.devices]

    def save_devices(self):
        with open(ADDRESS_BOOK_FILE, "w") as f:
            json.dump(
                [
                    {
                        "nickname": device.nickname,
                        "ip_address": device.ip_address,
                        "username": device.username,
                        "password": device.password,
                        "mac_address": device.mac_address,
                    }
                    for device in self.devices
                ],
                f,
                indent=4,
            )

    def load_devices(self):
        try:
            with open(ADDRESS_BOOK_FILE, "r") as f:
                devices_data = json.load(f)
                self.devices = [
                    Device(
                        d["nickname"],
                        d["ip_address"],
                        d["username"],
                        d["password"],
                        d["mac_address"],
                    )
                    for d in devices_data
                ]
                self.update_device_list()
        except FileNotFoundError:
            pass

    def ping_devices(self):
        while True:
            for device in self.devices:
                result = subprocess.run(
                    ["ping", "-n", "1", "-w", "1000", device.ip_address],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                )
                if result.returncode == 0:
                    device.status = "online"
                else:
                    device.status = "offline"
                time.sleep(5)
            self.update_device_list()

    def send_wol(self, instance):
        selected_index = self.rv.selected_index
        if selected_index is None:
            self.show_popup("Selection Error", "No device selected")
            return

        device = self.devices[selected_index]
        if device.mac_address:
            send_magic_packet(device.mac_address)
        else:
            self.show_popup("WOL Error", "MAC Address is not provided")

    def execute_command(self, command):
        selected_index = self.rv.selected_index
        if selected_index is None:
            self.show_popup("Selection Error", "No device selected")
            return

        device = self.devices[selected_index]
        if device.ip_address:
            try:
                protocol = Protocol(
                    endpoint=f"http://{device.ip_address}:5985/wsman",
                    username=device.username,
                    password=device.password,
                )
                shell_id = protocol.open_shell()
                protocol.run_command(shell_id, command)
                protocol.close_shell(shell_id)
            except Exception as e:
                self.show_popup("Execution Error", str(e))
        else:
            self.show_popup("Execution Error", "IP Address is not provided")

    def show_popup(self, title, message):
        popup = Popup(
            title=title,
            content=Label(text=message),
            size_hint=(None, None),
            size=(400, 200),
        )
        popup.open()


class MyApp(App):
    def build(self):
        return MainWindow()


if __name__ == "__main__":
    MyApp().run()
