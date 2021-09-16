import argparse
import json

from winreg import EnumKey, OpenKey, CloseKey, EnumValue, HKEY_LOCAL_MACHINE
from oui import get_oui_manufacturer


class NetworkHistory:
    def __init__(self, outfile):
        self.outfile = outfile

    def print_stored_networks(self):
        # Set registry path as raw string
        network_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\Unmanaged"

        # Open registry path with specified HKEY
        key = OpenKey(HKEY_LOCAL_MACHINE, network_path)

        print("Networks stored in registry: ")

        json_string = '['

        i = 0

        while True:
            try:
                # Attempt to open guid entries under \Unmanaged
                guid = EnumKey(key, i)
                # Open the entry for the guid
                guid_key = OpenKey(key, str(guid))
                # Get store name at position 4 and MAC address at position 5
                (n, name, t) = EnumValue(guid_key, 4)
                (n, address, t) = EnumValue(guid_key, 5)
                # Convert MAC address to hex from binary
                mac_address = str(address.hex(":")).upper()
                network_name = str(name)
                # Get first connected and last used date from the Windows registry for the network name
                first_connected, last_used = self.get_network_profile_dates(network_name)

                # Create OUI octet
                mac_octet = mac_address[0:2] + mac_address[3:5] + mac_address[6:8]
                manufacturer = get_oui_manufacturer(mac_octet)

                print('[+] ' + network_name + ' ' + mac_address + ' ' + manufacturer)

                # Build JSON contents
                if i == 0:
                    json_string += '"' + network_name + '"'
                else:
                    json_string += ',"' + network_name + '"'

                json_string += ', { "MAC": "' + mac_address + '"'
                json_string += ', "MAC manufacturer": "' + manufacturer.strip('()') + '"'
                json_string += ', "first_connected": "' + first_connected + '"'
                json_string += ', "last_used": "' + last_used + '"}'

                print("[+] First connected: " + first_connected)
                print("[+] Last used: " + last_used)

                CloseKey(guid_key)

                i += 1
            except Exception as e:
                CloseKey(key)
                break

        json_string += ']'

        # Format JSON
        parsed_json = json.loads(json_string)
        serialized_json = json.dumps(parsed_json, indent=4, sort_keys=True)

        # If specified as an argument, output JSON to outfile
        if self.outfile != "":
            with open(self.outfile, "w") as f:
                f.write(serialized_json)

    def get_network_profile_dates(self, network_name):
        profile_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles"

        key = OpenKey(HKEY_LOCAL_MACHINE, profile_path)

        first_connected = ""
        last_used = ""

        i = 0

        while True:
            try:
                # Attempt to open guid entries under \Profiles
                guid = EnumKey(key, i)
                # Open the entry for the guid
                guid_key = OpenKey(key, str(guid))
                (n, profile_name, t) = EnumValue(guid_key, 0)

                # Check if network name matches name in profiles entry
                if network_name == profile_name:
                    (n, first_connected, t) = EnumValue(guid_key, 4)
                    (n, last_used, t) = EnumValue(guid_key, 6)

                    first_connected = self.reg_binary_date_to_string(first_connected)
                    last_used = self.reg_binary_date_to_string(last_used)

                CloseKey(guid_key)

                i += 1

            except Exception as e:
                CloseKey(key)
                break

        return first_connected, last_used

    @staticmethod
    def reg_binary_date_to_string(binary):
        # Convert REG_BINARY date to big endian and then to integer
        year   = int((binary[1]  << 8) + binary[0])
        month  = int((binary[3]  << 8) + binary[2])
        day    = int((binary[7]  << 8) + binary[6])
        hour   = int((binary[9]  << 8) + binary[8])
        minute = int((binary[11] << 8) + binary[10])

        # Pad with leading zero
        day = str(day).zfill(2)
        month = str(month).zfill(2)
        hour = str(hour).zfill(2)
        minute = str(minute).zfill(2)

        return f"{day}-{month}-{year} {hour}:{minute}"


parser = argparse.ArgumentParser(description='Simple application for reading Windows registers for stored wireless network info.')
parser.add_argument("-o", "--outfile", type=str,
                                       help="Optional output file path if you want to store results in JSON format.",
                                       default="")

args = parser.parse_args()

history = NetworkHistory(args.outfile)
history.print_stored_networks()
