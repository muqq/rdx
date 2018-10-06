import os, sys
import json
import shutil
from pathlib import Path
from shutil import copyfile
from xlrd import open_workbook

class Rdx(object):
    def __init__(self, folderName, language, logo, pageName, textSide, header, text1, mediaSrc):
        self.folderName = folderName
        self.language = language
        self.logo = logo
        self.pageName = pageName
        self.textSide = textSide
        self.header = header
        self.text1 = text1
        self.mediaSrc = mediaSrc

    def __str__(self):
        return ("Rdx object:\n"
                "  Foler Name = {0}\n"
                "  Language = {1}\n"
                "  Logo = {2}\n"
                "  PageName = {3}\n"
                "  TextSide = {4} \n"
                "  Header = {5} \n"
                "  Text1 = {6} \n"
                "  MediaSrc = {7} \n"
                .format(self.folderName, self.language, self.logo, self.pageName, self.textSide, self.header,
                        self.text1,
                        self.mediaSrc))


dir = os.path.realpath(os.path.dirname(sys.argv[0]))
result_root = dir + '/data/RDX'
channel_path = result_root + '/Channel'
source_path = dir + "/data/source"
NEUTRAL = 'Neutral'
EN_US = 'en-US'


png_items = []
wmv_item = ""
for filename in os.listdir(source_path):
    if Path(filename).suffix == ".png":
        png_items.append(filename)
    elif Path(filename).suffix == ".wmv":
        wmv_item = filename


if not os.path.exists(channel_path):
    os.makedirs(channel_path)

results = {}

excel_path = '/data/X570UD RDX_Speed up_REV5.XLSX'
print("目標檔案夾" + dir)
wb = open_workbook(dir + excel_path)

for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    rows = []
    for row in range(1, number_of_rows):
        values = []
        for col in range(number_of_columns):
            value = (sheet.cell(row, col).value)
            try:
                value = str(int(value))
            except ValueError:
                pass
            finally:
                values.append(value)
        item = Rdx(*values)

        #check file exsit
        if item.mediaSrc in png_items:
            pass
        else:
            raise Exception('檔名和excel不符 {0}'.format(item.mediaSrc))

        language_path = os.path.join(channel_path, item.language)
        if not os.path.exists(language_path):
            os.makedirs(language_path)

        rdx_object = {
            "layout": "feature",
            "logo": "oem_assets/ASUS_logo_blue.png",
            "pageName": item.pageName,
            "textSide": item.textSide,
            "header": item.header,
            "text1": item.text1,
            "legalText": "",
            "media": {
                "blockType": "img",
                "src": "oem_assets/" + item.mediaSrc
            }
        }

        if item.language in results:
            results[item.language] = results[item.language] + [rdx_object]
        else:
            results[item.language] = [rdx_object]


for key in results.keys():
    output = {
        "identifier": "oem",
        "logo": "oem_assets/logo.png",
        "themeColor": "#222D5A",
        "pages": results[key]
    }
    path = os.path.join(channel_path, key)
    microsoft_prefix = "Microsoft.Getstarted_8wekyb3d8bbwe"
    microsoft_path = os.path.join(path, microsoft_prefix)
    asset_path = os.path.join(microsoft_path, 'oem_assets')
    if not os.path.exists(asset_path):
        os.makedirs(asset_path)

    for png in png_items:
        copyfile(os.path.join(source_path, png), os.path.join(asset_path, png))

    shutil.make_archive(os.path.join(path, 'oem'), 'zip', path, microsoft_prefix)
    shutil.rmtree(microsoft_path)
    with open(path + '/oem.json', 'w', encoding="utf8") as fp:
        json.dump(output, fp, ensure_ascii=False, sort_keys=True, indent=4)
        print("{0} 成功".format(key))

def createNeutral(path):
    if not os.path.exists(path):
        os.makedirs(path)
    copyfile(os.path.join(source_path, wmv_item), os.path.join(path, wmv_item))


def createEnglishCopy(path):
    dst_path = os.path.join(result_root, path)
    if not os.path.exists(dst_path):
        os.makedirs(dst_path)
    shutil.copytree(os.path.join(channel_path, EN_US), os.path.join(dst_path, EN_US))
    shutil.copytree(os.path.join(channel_path, NEUTRAL), os.path.join(dst_path, NEUTRAL))


createNeutral(os.path.join(channel_path, NEUTRAL))
createEnglishCopy('BBY')
createEnglishCopy('Signature')
