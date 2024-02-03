"""This Script is used to get the Detail of the system or System
Configuration. It is written to specifically work on Windows
Operating System."""

# Importing necessary module and libraries.
import subprocess
import sys


def check_dependency(package):
    try:
        # Attempt to import the module
        __import__(package)
        print(f"{package} is already installed.")
        return True
    except ImportError:
        print(f"{package} is not installed.")
        return False


def install_dependency(package):
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        print(f"{package} has been successfully installed.")
    except subprocess.CalledProcessError:
        print(f"Failed to install {package}.")


def get_installed_software():
    """This Function will return the list of
    software installed on your system."""

    # Checking the OS, if it is windows or any other OS.
    if platform.system() == "Windows":
        # Run the command to fetch the list and read the
        # output and store it in the program variable.
        programs = os.popen('wmic product get name').read()
        return programs.split('\n')[1:-1]
    else:
        return "This Operating System is not Supported."


def get_internet_speed():
    test = speedtest.Speedtest()
    download_speed_info = test.download() / 10 ** 6  # Convert to Mbps
    upload_speed_info = test.upload() / 10 ** 6  # Convert to Mbps
    return download_speed_info, upload_speed_info


def get_screen_resolution():
    try:
        import screeninfo
        screen = screeninfo.get_monitors()[0]
        return f"{screen.width}x{screen.height} pixels"
    except ImportError:
        return ("Screen resolution information not available (screeninfo "
                "library is not installed)")


def get_cpu_info():
    cpu_details = {'model': platform.processor(), 'cores': psutil.cpu_count(logical=False),
                   'threads': psutil.cpu_count(logical=True)}
    return cpu_details


def get_gpu_info():
    try:
        import GPUtil
        gpu_info = GPUtil.getGPUs()[0]
        return gpu_info.name if gpu_info else "No GPU detected"
    except ImportError:
        return "GPU information not available (GPUtil library is not installed)"


def get_ram_size():
    return round(psutil.virtual_memory().total / (1024 ** 3), 2)  # Convert to GigaByte


def get_screen_size():
    try:
        import screeninfo
        import math
        monitor_info = screeninfo.get_monitors()[0]
        screen_size = round(math.sqrt((monitor_info.height_mm ** 2 + monitor_info.width_mm ** 2)) / 25.4, 1)
        return f'{screen_size} inches'
    except ImportError:
        return "Screen size information not available"


def get_mac_address(interface="Ethernet"):
    try:
        mac_address = ':'.join(
            ['{:02x}'.format((int(uuid.getnode()) >> elements) & 0xff) for elements in range(0, 2 * 6, 2)][::-1])
        return mac_address
    except Exception as e:
        return f"Error fetching MAC address: {str(e)}"


def get_public_ip():
    try:
        public_ip = socket.gethostbyname(socket.gethostname())
        return public_ip
    except Exception as e:
        return f"Error fetching public IP address: {str(e)}"


def get_windows_version():
    return platform.version()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    dependencies = ['speedtest-cli', 'screeninfo', 'psutil', 'GPUtil', 'platform', 'socket', 'os']
    for dependency in dependencies:
        if not check_dependency(dependency):
            install_dependency(dependency)

    import psutil
    import platform
    import speedtest
    import socket
    import os
    import uuid

    print("Installed Software:")
    for software in get_installed_software():
        print(f"- {software}")

    download_speed, upload_speed = get_internet_speed()
    print(f"\nInternet Speed:")
    print(f"- Download Speed: {download_speed:.2f} Mbps")
    print(f"- Upload Speed: {upload_speed:.2f} Mbps")

    print(f"\nScreen Resolution: {get_screen_resolution()}")

    print("\nCPU Information:")
    cpu_info = get_cpu_info()
    print(f"- Model: {cpu_info['model']}")
    print(f"- Number of Cores: {cpu_info['cores']}")
    print(f"- Number of Threads: {cpu_info['threads']}")

    print(f"\nGPU Information: {get_gpu_info()}")

    print(f"\nRAM Size: {get_ram_size()} GB")

    print(f"\nScreen Size: {get_screen_size()}")

    print(f"\nMac Address (Ethernet): {get_mac_address('Ethernet')}")

    print(f"\nPublic IP Address: {get_public_ip()}")

    print(f"\nWindows Version: {get_windows_version()}")
