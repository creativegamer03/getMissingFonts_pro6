import os
import re
import xml.etree.ElementTree as ET
from base64 import b64decode as b64d
from win32 import win32gui

os.chdir(os.path.join(os.environ['USERPROFILE'], "Documents\ProPresenter6"))

files = [x for x in os.listdir() if os.path.splitext(x)[1] == ".pro6"]
missingFonts = []
x_files = []
log = []

def callback(font, tm, fonttype, names):
    names.append(font.lfFaceName)
    return True

fontnames = []
hdc = win32gui.GetDC(None)
win32gui.EnumFontFamilies(hdc, None, callback, fontnames)

def start():
    for x in files:
        try:
            xml_tree = ET.parse(x)
            xml_root = xml_tree.getroot()
            log.append(x)
            for tag in xml_root.findall(".//NSString[@rvXMLIvarName='RTFData']"):
                rtf_str = str(b64d(tag.text), encoding='utf-8')
                # do something here to get the font names from rtf data string
                try:
                    rtf_fonts = re.findall(r"\{\\fonttbl\{[^}]*\}\{[^}]*\}\{[^}]*\}\}", rtf_str)
                    fonts = re.findall(r"([a-zA-Z]+( [a-zA-Z]+)+);", rtf_fonts[0])
                    for font in fonts:
                        if font[0] in missingFonts:
                            pass
                        elif font[0] in fontnames:
                            pass
                        else:
                            missingFonts.append(font[0])
                except IndexError:
                    pass
        except Exception as e:
            x_files.append((x, e))
            pass
        
    missingFonts.sort()

    with open('./missingFonts.txt', 'w+') as fp:
        fp.writelines('\n'.join(missingFonts)+'\n\n=========[LOG]=========\n'+' - ok\n'.join(log)+'\n'+'\n')
        for x in x_files:
            fp.write("{} <{}>\n".format(x[0], x[1]))
    os.startfile('missingFonts.txt')


if __name__ == '__main__':
    start()
