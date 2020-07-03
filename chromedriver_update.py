# importing libraires
import sys

def update_check():
    '''
    Checks current version of Chromedriver and ensures it is the correct version for the version of Chrome installed on computer.

    If version is not correct, downloads correct version of Chrome into current directory and cleans up.

    Creates a version.txt document to assist in this process.

    Import and run function at start of programme.

    Designed to be put in a python executable run by average user (e.g. a Selenium based application for searching websites automatically).

    Does not use beautiful soup in order to reduce size of executable.

    Print functions not needed in final distributed version.

    '''

    # checking is system is Windows, if not, skipping chromedriver check
    platform = sys.platform

    if platform != 'windows':
        print('Not windows system, skipping...')

    # function to open version.txt document and update it, and create if one does not exist

    else:

        import os
        from win32com.client import Dispatch
        import urllib.request
        from zipfile import ZipFile

        def version_control():
            with open('version.txt', 'w+') as f:
                f.write(driver_version)

        # function to get Chrome browser version
        def get_version_via_com(filename):
            parser = Dispatch("Scripting.FileSystemObject")
            try:
                version = parser.GetFileVersion(filename)
            except Exception:
                return None
            return version

        # store Chrome version
        version = ''

        # get Chrome browser version (assumes Chrome installed in default location, paths can be added as needed)
        paths = [r'C:/Program Files/Google/Chrome/Application/chrome.exe',
                 r'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe']
        version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
        print(version)

        # format version number for URL to check for chromedriver version
        short_version = version.split('.')[:3]
        short_version = '.'.join(short_version)
        print(short_version)

        # checking required chromedriver version
        driver_version = urllib.request.urlopen('https://chromedriver.storage.googleapis.com/LATEST_RELEASE_' + short_version).read().decode('utf-8')
        print(driver_version)

        # opening text file to check current chromedriver version (avoids checking via chromedriver function, since if it is out of date, it will fail)
        f = open('version.txt', 'w+')
        current_version = f.read()

        # checking if current chromedriver version = required chromedriver version and if not, updating and cleaning up
        if driver_version == current_version:
            f.close()

        else:
            url = 'https://chromedriver.storage.googleapis.com/' + driver_version + '/chromedriver_win32.zip'

            remote = urllib.request.urlopen(url)  # read remote file
            data = remote.read()  # read from remote file
            remote.close()  # close urllib request
            local = open('chromedriver_win32.zip', 'wb')  # write binary to local file
            local.write(data)
            local.close()  # close file

            # Create a ZipFile Object and loads updated zip file
            with ZipFile('chromedriver_win32.zip', 'r') as zipObj:
                # Extract all the contents of zip file in current directory (assumes chromedriver is in same directory as exe file for progamme)
                zipObj.extractall()

            os.remove('./chromedriver_win32.zip')
            f.close()
            print('updated')
            # calls function to update text file with chromedriver version
            version_control()


update_check()
