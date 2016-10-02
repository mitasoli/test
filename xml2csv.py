import csv
import time
import os
from xml.etree.ElementTree import iterparse
import xml.etree.ElementTree as elementTree
from xml.dom.minidom import parseString
import logging
import logging.config
import zipfile


def zip_dir(zipname, dir_to_zip):
    dir_to_zip_len = len(dir_to_zip.rstrip(os.sep)) + 1
    with zipfile.ZipFile(zipname, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for dirname, subdirs, files in os.walk(dir_to_zip):
            for filename in files:
                path = os.path.join(dirname, filename)
                entry = path[dir_to_zip_len:]
                zf.write(path, entry)


def get_values_by_tagname_iterparse(file):
    print ('Retrieving values from tag ')
    values = []
    tag_stack = []
    for event, elem in elementTree.iterparse(file, events=('start', 'end')):
        if event == 'start':
            tag_stack.append(elem.tag)
            # elem_stack.append(elem)
            # print(elem_stack)
        elif event == 'end':
            try:
                value = elem.text
                values.append(value)
                tag_stack.pop()
                # elem_stack.pop()
                elem.clear()
            except IndexError:
                pass

    return values


def file_process(infilename, outfilename, file_contents_data):
    try:
        start_time = time.time()
        dict_log_config = {
            "version": 1,
            "handlers": {
                "fileHandler": {
                    "class": "logging.FileHandler",
                    "formatter": "myFormatter",
                    "filename": "XMLLOG/processedFilesList.log"
                }
            },
            "loggers": {
                "xmlConversionApp": {
                    "handlers": ["fileHandler"],
                    "level": "INFO",
                }
            },

            "formatters": {
                "myFormatter": {
                    "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
                }
            }
        }
        # print(dict_log_config)
        logging.config.dictConfig(dict_log_config)

        logger = logging.getLogger("xmlConversionApp")
        # logging.basicConfig(filename="LOG/processedFilesList.log", level=logging.INFO)
        logger.info("File: " + infilename)
        dom = parseString(file_contents_data)
        row_data = []
        mt_h = []
        main_node = dom.getElementsByTagName("md")
        # mi_node = dom.getElementsByTagName("mi")
        for index, value in enumerate(main_node):
            for mi_i, mi_v in enumerate(value.getElementsByTagName("mi")):
                row_data = []
                # mt_r = []
                moid_ls = []
                mt_h = []
                mt_r_dict = {}

                header_dict = ["DATE", "INTERVAL", "MO"]

                row_data.append(mi_v.getElementsByTagName("mts")[0].childNodes[0].data)
                row_data.append(mi_v.getElementsByTagName("gp")[0].childNodes[0].data)
                for mt_i, mt_v in enumerate(mi_v.getElementsByTagName("mt")):
                    mt_h.append(mt_v.childNodes[0].data)
                # this line creates an empty column for moid.
                # row_data.append("moid")
                for mv_i, mv_v in enumerate(mi_v.getElementsByTagName("mv")):
                    moid_ls.append(mv_v.getElementsByTagName("moid")[0].childNodes[0].data)
                    # mt_len = len(mi_v.getElementsByTagName("mt"))
                    for mt_i, mt_v in enumerate(mi_v.getElementsByTagName("mt")):
                        try:
                            if mv_v.getElementsByTagName("r")[mt_i].firstChild:
                                mt_r_dict[mv_i].append(mv_v.getElementsByTagName("r")[mt_i].childNodes[0].data)
                        except IndexError:
                            mt_r_dict[mv_i].append(" ")
                        except KeyError:
                            mt_r_dict[mv_i] = [mv_v.getElementsByTagName("r")[mt_i].childNodes[0].data]

            for i, mt in enumerate(mt_h):
                try:
                    header_dict.append(mt_h[i])
                except IndexError:
                    header_dict.append(" ")
            print ("oldname", old_name)
            # str_filename = outfilename + '/' + old_name + '_' + str(index) + '.csv'
            old_name_i = old_name.find("_")

            str_filename = outfilename + '/' + old_name[0:old_name_i] + '_' + str(index) + '.csv'

            filename = str_filename.replace("\\", "/")
            print (filename)
            FILE = open(filename, "w", newline='')
            output = csv.writer(FILE, dialect='excel')
            output.writerow(header_dict)
            fix = row_data
            # lin = []
            if not mt_r_dict:
                row = []
                row.extend(fix)
                output.writerow(row)
                row = []
            for key, value in mt_r_dict.items():
                try:

                    row = []
                    row.extend(fix)
                    row.append(moid_ls[key])
                    row.extend(value)
                    output.writerow(row)
                    row = []

                except IndexError:
                    print("Error")

            row_data = []
            mt_r_dict = {}
            FILE.close()
            logger.info("Done!")
        dom.unlink()
    except IOError:
        print("Error: can\'t find file or read data")
    else:
        print("Written content in the file successfully")


def readfile_contents(file):
    fp = open(file, "r")
    raw_data = fp.read()
    fp.close()
    return raw_data


if __name__ == "__main__":
    main_dir_name = "RNCXML"
    for dir_name, subdirs, files in os.walk(main_dir_name):
        for filename in files:
            if filename.endswith(".xml"):
                infile_name = os.path.join(main_dir_name, filename)
                old_name = os.path.splitext(filename)[0]
                old_name_index = old_name.find("_")
                str_dir_name = dir_name.replace("\\", "/")
                str_newpath = str_dir_name + '//' + old_name[0:old_name_index] + 'CSV'
                newpath = str_newpath.replace("//", "/")
                if not os.path.exists(newpath):
                    os.makedirs(newpath)

                str_content = os.path.join(dir_name, filename)
                file_contents = readfile_contents(str_content)
                file_process(filename, newpath, file_contents)
                zip_dir("%s.rar" % (os.path.join(dir_name)), str_dir_name)
                os.rmdir("dir_name")
                # path_name = os.path.abspath(os.path.join(dir_name, filename))
                # zout = zipfile.ZipFile("%s.zip" % (os.path.join(dir_name, old_name)), "w", zipfile.ZIP_DEFLATED)
                # arcname = path_name[len(os.path.abspath(dir_name)) + 1:]
                # zout.write(path_name, arcname)
                # zout.close()
                filename = ''
                file_contents = []
                newpath = ''

