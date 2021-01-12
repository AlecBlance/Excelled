""" Removes password from excel files """

import zipfile
import re
import os
import sys
import shutil
from os import walk
from zipfile import ZipFile as zipper


def main():
    _, _, filenames = next(walk("original/"))
    for x_file in filenames:
        to_convert = copy(x_file)
        converted = convert(to_convert, to_zip=True)
        convert(remove_password(converted))
    print("Done! Your unprotected file/s can be found in output/ folder.")


def copy(x_file):
    cwd = os.getcwd()
    output_loc = f'{cwd}/output'
    if not os.path.exists(output_loc):
        os.mkdir(output_loc)
    out_file = f'{output_loc}/{x_file}'
    shutil.copy(f'original/{x_file}', out_file)
    return out_file


def convert(x_file, to_zip=False):
    ext = "xlsx"
    current = "zip"
    if to_zip:
        ext = "zip"
        current = "xlsx"
    zipped = f'{x_file.replace(f".{current}", "")}-excelled.{ext}'
    os.rename(x_file, zipped)
    return zipped


def remove_password(x_file):
    path = x_file.replace('-excelled', '')
    with zipper(x_file, 'r') as in_xfile, zipper(path, 'w') as out_xfile:
        files = [item for item in in_xfile.infolist()]
        for info in files:
            content = in_xfile.read(info)
            if b'sheetProtection' in content:
                regex = r'<sheetProtection(.*?)/>'
            elif b'workbookProtection' in content:
                regex = r'<workbookProtection(.*?)/>'
            else:
                out_xfile.writestr(info, content)
                continue
            content = content.decode("utf-8")
            text = re.sub(regex, '', content)
            if len(text) != len(content):
                out_xfile.writestr(info, text)
    os.remove(x_file)
    return path


if __name__ == '__main__':
    main()
