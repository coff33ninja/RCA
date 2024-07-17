# RCA/Remote Control application
  (Name to be decided)

This Python application allows you to remotely control devices via WinRM (Windows Remote Management) and Wake-on-LAN (WOL). It provides a graphical user interface (GUI) built with PyQt6 for managing devices and executing remote commands.

## Features

- **Add, Edit, and Remove Devices**: Easily add, edit, and remove devices by entering their IP address, credentials, and optional MAC address.
- **Remote Actions**: Execute remote actions such as rebooting and shutting down devices using WinRM.
- **Wake-on-LAN (WOL)**: Send Wake-on-LAN packets to wake devices from sleep or powered-off state.
- **Continuous Ping**: Automatically ping devices to check their online status.
- **Device Address Book**: Stores device details in a JSON file for easy retrieval on application startup.

## Prerequisites

- Python 3.6+
- Required Python packages: PyQt6, requests-ntlm, wakeonlan

Install the required packages using pip:

```bash
pip install requests-ntlm wakeonlan kivy kivy-garden pywinrm
```
### Usage

1. **Run the Application**
2. **Add Devices**:
   - Enter the IP address of the device.
   - Optionally, enter the username, password, and MAC address.
   - Click on "Add Device" to add the device to the list.
3. **Execute Remote Actions**:
   - Select a device from the list.
   - Click on "Reboot Device" or "Shutdown Device" to execute the corresponding action using WinRM.
   - Click on "Wake-on-LAN Device" to send a Wake-on-LAN packet to the selected device.
4. **Continuous Ping**:
   - Devices in the list will be continuously pinged to check their online status.
5. **Manage Devices**:
   - Use the "Edit Device" button to modify the details of a selected device.
   - Use the "Remove Device" button to delete a device from the list.

### Notes

- Ensure devices are reachable over the network, and WinRM is properly configured on Windows devices for remote management.
- Save device details automatically to `device_address_book.json`.
- On application startup, previously saved devices are loaded from the JSON file.
