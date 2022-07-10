import win32com.client as win32
import os
import winreg
import logging
import sys


if __name__ == '__main__':
    root_dir = os.getcwd()

    try:
        assert 'hwpx' in os.listdir(root_dir), 'There is no hwpx output directory.'
        assert 'hwp' in os.listdir(root_dir), 'There is no hwp input directory.'
    except Exception:
        logging.exception('Error has occurred')
        os.system("pause")
        sys.exit()

    try:
        hwpx_dir = os.path.join(root_dir, 'hwpx')
        hwp_dir = os.path.join(root_dir, 'hwp')
        dll_path = os.path.join(root_dir, 'AutomationSecurity.dll')

        handle_key = winreg.OpenKeyEx(winreg.HKEY_CURRENT_USER, "SOFTWARE\\HNC\\HwpAutomation\\Modules",
                                      0, winreg.KEY_ALL_ACCESS)

        print(winreg.QueryValueEx(handle_key, "AutomationSecurity"))
        winreg.CloseKey(handle_key)
    except Exception:
        logging.exception("Error has occurred. You have to update registry.")
        os.system("pause")
        sys.exit()

    try:
        print('accessing hwpframe control...')
        hwp = win32.gencache.EnsureDispatch('hwpframe.hwpobject')
        hwp.RegisterModule("FilePathCheckDLL", "AutomationSecurity")
    except Exception:
        logging.exception("Error has occurred. While loading hwpframe.")
        os.system("pause")
        sys.exit()

    arg = "suspendpassword:True;versionworning:False"

    for file in os.listdir(hwp_dir):
        if file.endswith('.hwp'):
            path = os.path.join(hwp_dir, file)
            print('[Opening] {}'.format(file))
            hwp.Open(path, arg=arg)
            path = os.path.join(hwpx_dir, file + 'x')
            hwp.SaveAs(path, 'HWPX')
            print('[Saved] {}'.format(file + 'x'))

    hwp.Clear(1)
    hwp.Quit()
    print('done!')
    os.system("pause")
