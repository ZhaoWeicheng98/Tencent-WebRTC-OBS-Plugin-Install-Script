import ctypes
import errno
import hashlib
import json
import os
import platform
import shutil
import sys
import winreg
from os import path
import traceback
import psutil

import colorama
from colorama import Fore, Style
from pyfiglet import Figlet
from win32api import HIWORD, LOWORD, GetFileVersionInfo

PACKAGE_FILES = {
    "obs-plugins/64bit/libcrypto-1_1-x64.dll": "2aeb5ce32a1d60a588054e599bb14736",
    "obs-plugins/64bit/libssl-1_1-x64.dll": "29d9a6d1e9ba7c691d778a4c0810b116",
    "obs-plugins/64bit/obs-webrtc.dll": "cca4f3eb9cd76e9bbaea393c6bbf332e",
    "obs-plugins/64bit/websocketclient.dll": "aad616ebeeddb30be4222ad2dff7287c"
}


def is_user_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() == 1
    except AttributeError:
        return False


def calculate_check_sum(file_path):
    # Open,close, read file and calculate MD5 on its contents
    with open(file_path, 'rb') as file_to_check:
        # read contents of the file
        data = file_to_check.read()
        # pipe contents of the file through
        md5_returned = hashlib.md5(data).hexdigest()
    return md5_returned


def get_obs_path():
    # connecting to key in registry
    access_registry = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)

    access_key = winreg.OpenKey(access_registry, r"SOFTWARE\OBS Studio")
    # accessing the key to open the registry directories under
    try:
        x = winreg.QueryValue(access_key, "")
        return x
    except WindowsError:
        return ""


def get_version_number(filename):
    try:
        info = GetFileVersionInfo(filename, "\\")
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        return HIWORD(ms), LOWORD(ms), HIWORD(ls), LOWORD(ls)
    except:
        return 0, 0, 0, 0


def install_file(source_file_path, dest_dir_path):
    file_name = path.basename(source_file_path)
    if path.exists(path.join(dest_dir_path, file_name)):
        backup_file_name = file_name + ".backup"
        os.rename(path.join(dest_dir_path, file_name),
                  path.join(dest_dir_path, backup_file_name))
    dest_file_path = shutil.copy2(source_file_path, dest_dir_path)
    return dest_file_path


def uninstall_file(file_path):
    if path.exists(file_path):
        os.remove(file_path)
    backup_file_path = file_path + ".backup"
    if path.exists(backup_file_path):
        os.rename(backup_file_path, file_path)


def install_service(service_file):
    TENCENT_WEBRTC_SERVICE = \
        {
            "name": "Tencent webrtc",
            "common": True,
            "servers":
            [
                {
                    "name": "Default",
                    "url": "https://webrtcpush.myqcloud.com/webrtc/v1/pushstream"
                }
            ],
            "recommended":
            {
                "keyint": 2,
                "output": "tencentcloud_output"
            }
        }

    with open(service_file, "r", encoding="utf-8") as f:
        services_config = json.load(f)
    services = services_config.get("services", [])
    for service in services:
        service_name = service.get("name", "")
        if service_name == "Tencent webrtc":
            services.remove(service)
            break
    services.append(TENCENT_WEBRTC_SERVICE)
    services_config['services'] = services
    with open(service_file, "w", encoding="utf-8") as f:
        json.dump(services_config, f, ensure_ascii=False, indent=4)


def uninstall_service(service_file):
    with open(service_file, "r", encoding="utf-8") as f:
        services_config = json.load(f)
    services = services_config.get("services", [])
    for service in services:
        service_name = service.get("name", "")
        if service_name == "Tencent webrtc":
            services.remove(service)
            break
    services_config['services'] = services
    with open(service_file, "w", encoding="utf-8") as f:
        json.dump(services_config, f, ensure_ascii=False, indent=4)


if __name__ == "__main__":
    colorama.init()
    print(f"??????????????????????????????...", end='\r')
    if platform.system().lower() != "windows":
        print(
            f"??????????????????????????????...{platform.system()} {Fore.RED}??????????????????????????? Windows{Style.RESET_ALL}")
        os.system("pause")
        sys.exit(errno.EPERM)
    else:
        print(
            f"??????????????????????????????...{platform.system()} {Fore.GREEN}OK{Style.RESET_ALL}")

    print(f"????????????????????????...", end='\r')
    if platform.machine().lower() != "amd64":
        print(
            f"????????????????????????...{platform.machine()} {Fore.RED}??????????????????????????? AMD64 ??????{Style.RESET_ALL}")
        os.system("pause")
        sys.exit(errno.EPERM)
    else:
        print(
            f"????????????????????????...{platform.machine()} {Fore.GREEN}OK{Style.RESET_ALL}")

    print(f"???????????????????????????...", end='\r')
    if platform.release() not in ['10', '11']:
        print(
            f"???????????????????????????...{platform.release()} {Fore.RED}??????????????????????????? Windows 10 ??? Windows 11{Style.RESET_ALL}")
        os.system("pause")
        sys.exit(errno.EPERM)
    else:
        print(
            f"???????????????????????????...{platform.release()} {Fore.GREEN}OK{Style.RESET_ALL}")

    print(f"????????????????????????...", end='\r')
    if not is_user_admin():
        print(f"????????????????????????... {Fore.RED}??????????????????????????????????????????{Style.RESET_ALL}")
        os.system("pause")
        sys.exit(errno.EACCES)
    else:
        print(f"????????????????????????... {Fore.GREEN}OK{Style.RESET_ALL}")

    print(f"?????????????????????????????????...", end='\r')
    for index, (file_path, checksum) in enumerate(PACKAGE_FILES.items()):
        print(f"?????????????????????????????????... ({index+1}/{len(PACKAGE_FILES)})", end='\r')
        if not (path.exists(file_path) and path.isfile(file_path)):
            print(
                f"?????????????????????????????????... ({index+1}/{len(PACKAGE_FILES)}) {Fore.RED}???????????????{file_path}?????????{Style.RESET_ALL}")
            os.system("pause")
            sys.exit(errno.ENOENT)

        if calculate_check_sum(file_path) != checksum:
            print(
                f"?????????????????????????????????... ({index+1}/{len(PACKAGE_FILES)}) {Fore.RED}???????????????{file_path}???????????????{Style.RESET_ALL}")
            os.system("pause")
            sys.exit(errno.ENOENT)
        else:
            print(
                f"?????????????????????????????????... ({index+1}/{len(PACKAGE_FILES)}) {file_path} {Fore.GREEN}OK{Style.RESET_ALL}")

    print(f"???????????? OBS Studio ????????????...", end='\r')
    obs_install_path = get_obs_path()
    if not (obs_install_path and path.exists(obs_install_path) and path.isdir(obs_install_path)):
        print(
            f"???????????? OBS Studio ????????????...{obs_install_path} {Fore.RED}????????????????????? OBS Studio ????????????{Style.RESET_ALL}")
        os.system("pause")
        sys.exit(errno.ENOENT)
    elif not path.exists(path.join(obs_install_path, r"bin\64bit\obs64.exe")):
        print(
            f"???????????? OBS Studio ????????????...{obs_install_path} {Fore.RED}?????????????????? OBS Studio ?????????{Style.RESET_ALL}")
        os.system("pause")
        sys.exit(errno.ENOENT)
    else:
        print(
            f"???????????? OBS Studio ????????????...{obs_install_path} {Fore.GREEN}OK{Style.RESET_ALL}")

    print(f"???????????? OBS Studio ??????...", end='\r')
    obs_version, obs_subversion, obs_patch, _ = get_version_number(
        path.join(obs_install_path, r"bin\64bit\obs64.exe"))
    if obs_version < 26:
        print(
            f"???????????? OBS Studio ??????... {obs_version}.{obs_subversion}.{obs_patch} {Fore.RED}?????????OBS Studio ????????????{Style.RESET_ALL}")
        os.system("pause")
        sys.exit(errno.ENOENT)
    else:
        print(
            f"???????????? OBS Studio ??????... {obs_version}.{obs_subversion}.{obs_patch} {Fore.GREEN}OK{Style.RESET_ALL}")

    print(f"???????????? OBS Studio ??????...", end='\r')
    if "obs64.exe" in (p.name() for p in psutil.process_iter()):
        print(
            f"???????????? OBS Studio ??????... {Fore.RED}?????????OBS Studio ?????????????????????????????????{Style.RESET_ALL}")
        os.system("pause")
        sys.exit(errno.EAGAIN)
    else:
        print(
            f"???????????? OBS Studio ??????... {Fore.GREEN}OK{Style.RESET_ALL}")

    os.system("cls")
    width = os.get_terminal_size().columns

    print("".join(["="]*width))

    f = Figlet(font='slant')
    print(f.renderText('AsaChiri.cn'))

    print("".join(["="]*width))
    print(f"{Fore.YELLOW}???????????? Tencent WebRTC OBS ?????????????????????{Style.RESET_ALL}")
    print(f"?????????????????????????????????")
    print(f"{Fore.YELLOW} 1 {Style.RESET_ALL} ?????? Tencent WebRTC OBS ??????")
    print(f"{Fore.YELLOW} 2 {Style.RESET_ALL} ?????? Tencent WebRTC OBS ??????")
    print(f"{Fore.YELLOW} 0 {Style.RESET_ALL} ???????????????")
    action = input("????????? [0,1,2]???")
    if action not in ["0", "1", "2"]:
        print(f"{Fore.RED} ?????????????????????{Style.RESET_ALL} ???????????????...")
        os.system("pause")
        sys.exit(errno.ENOSYS)
    elif action == "0":
        print(f"{Fore.BLUE} ?????????{Style.RESET_ALL} ???????????????...")
        os.system("pause")
        sys.exit(0)
    elif action == "1":
        plugin_dir_path = path.join(obs_install_path, r"obs-plugins\64bit")
        if not (path.exists(plugin_dir_path) and path.isdir(plugin_dir_path)):
            print(f"{Fore.RED} OBS Studio ????????????????????????{Style.RESET_ALL} ???????????????...")
            os.system("pause")
            sys.exit(errno.ENOENT)

        print(f"??????????????????...", end='\r')
        for index, (file_path, checksum) in enumerate(PACKAGE_FILES.items()):
            print(f"??????????????????... ({index+1}/{len(PACKAGE_FILES)})", end='\r')
            try:
                install_file(file_path, plugin_dir_path)
                print(
                    f"??????????????????... ({index+1}/{len(PACKAGE_FILES)}) {file_path} {Fore.GREEN}OK{Style.RESET_ALL}")
            except:
                print(
                    f"??????????????????... ({index+1}/{len(PACKAGE_FILES)}) {Fore.RED}???????????????{file_path}??????????????????{Style.RESET_ALL}")
                traceback.print_exc()
                os.system("pause")
                sys.exit(1)

        service_config_path_local = path.join(
            obs_install_path, r"data\obs-plugins\rtmp-services\services.json")
        if not (path.exists(service_config_path_local) and path.isfile(service_config_path_local)):
            print(f"{Fore.RED} OBS Studio RTMP ??????????????????????????????{Style.RESET_ALL} ???????????????...")
            os.system("pause")
            sys.exit(errno.ENOENT)

        print(f"????????????????????????...", end='\r')
        try:
            install_service(service_config_path_local)
            print(
                f"????????????????????????... {Fore.GREEN}OK{Style.RESET_ALL}")
        except:
            print(
                f"????????????????????????... {Fore.RED}?????????????????????????????????????????????{Style.RESET_ALL}")
            traceback.print_exc()
            os.system("pause")
            sys.exit(1)

        service_config_path_roaming = path.join(path.expandvars(
            "%APPDATA%"), r"obs-studio\plugin_config\rtmp-services\services.json")
        if (path.exists(service_config_path_local) and path.isfile(service_config_path_local)):
            print(f"????????????????????????...", end='\r')
            try:
                install_service(service_config_path_roaming)
                print(
                    f"????????????????????????... {Fore.GREEN}OK{Style.RESET_ALL}")
            except:
                print(
                    f"????????????????????????... {Fore.RED}?????????????????????????????????????????????{Style.RESET_ALL}")
                traceback.print_exc()
                os.system("pause")
                sys.exit(1)

        print(f"{Fore.BLUE}??????????????????????????????????????? OBS Studio????????????{Style.RESET_ALL} ???????????????...")
        os.system("pause")
        sys.exit(0)

    elif action == "2":
        plugin_dir_path = path.join(obs_install_path, r"obs-plugins\64bit")
        if (path.exists(plugin_dir_path) and path.isdir(plugin_dir_path)):
            print(f"??????????????????...", end='\r')
            for index, (file_path, checksum) in enumerate(PACKAGE_FILES.items()):
                print(f"??????????????????... ({index+1}/{len(PACKAGE_FILES)})", end='\r')
                try:
                    file_name = path.basename(file_path)
                    uninstall_file(path.join(plugin_dir_path, file_name))
                    print(
                        f"??????????????????... ({index+1}/{len(PACKAGE_FILES)}) {file_path} {Fore.GREEN}OK{Style.RESET_ALL}")
                except:
                    print(
                        f"??????????????????... ({index+1}/{len(PACKAGE_FILES)}) {Fore.RED}???????????????{file_path}??????????????????{Style.RESET_ALL}")
                    traceback.print_exc()
                    os.system("pause")
                    sys.exit(1)

        service_config_path_local = path.join(
            obs_install_path, r"data\obs-plugins\rtmp-services\services.json")
        if (path.exists(service_config_path_local) and path.isfile(service_config_path_local)):
            print(f"????????????????????????...", end='\r')
            try:
                uninstall_service(service_config_path_local)
                print(
                    f"????????????????????????... {Fore.GREEN}OK{Style.RESET_ALL}")
            except:
                print(
                    f"????????????????????????... {Fore.RED}?????????????????????????????????????????????{Style.RESET_ALL}")
                traceback.print_exc()
                os.system("pause")
                sys.exit(1)

        service_config_path_roaming = path.join(path.expandvars(
            "%APPDATA%"), r"obs-studio\plugin_config\rtmp-services\services.json")
        if (path.exists(service_config_path_local) and path.isfile(service_config_path_local)):
            print(f"????????????????????????...", end='\r')
            try:
                uninstall_service(service_config_path_roaming)
                print(
                    f"????????????????????????... {Fore.GREEN}OK{Style.RESET_ALL}")
            except:
                print(
                    f"????????????????????????... {Fore.RED}?????????????????????????????????????????????{Style.RESET_ALL}")
                traceback.print_exc()
                os.system("pause")
                sys.exit(1)

        print(f"{Fore.BLUE}??????????????????????????????????????? OBS Studio????????????{Style.RESET_ALL} ???????????????...")
        os.system("pause")
        sys.exit(0)
