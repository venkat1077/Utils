#!/usr/bin/env python

"""
This module is to write the details of the functions/class_methods of the given
python package in an excel sheet.
How to execute the package in command line:
    format: python pkgtoxls.py <python_package> <excel_file_path>
    example: python pkgtoxls.py pkg1.subpkg1 dir1/dir2/file1.xlsx
    (python_package should be able to be imported from the current
    execution path)
"""

import sys
# import os
import inspect
import pkgutil
import importlib

# Required third party modules
try:
    import xlsxwriter
except ImportError:
    print "Please install xlsxwriter python package"

__author__ = 'Venkatesh Sethuram'
__version__ = '1.0'


def get_module_functions(module):
    """ To get the functions defined in the given module """

    functions_list = []
    for val in module.__dict__.itervalues():
        if inspect.isfunction(val) and inspect.getmodule(val) == module and \
         not (val.__name__).startswith('_'):
            doc_str = inspect.getdoc(val)
            if isinstance(doc_str, str):
                doc_str = " ".join(doc_str.split())
            functions_list.append({'func_name': val.__name__.split('.')[-1],
                                   'class_name': "",
                                   "doc_str": doc_str,
                                   'args_list': inspect.getargspec(val)[0]})
        elif inspect.isclass(val) and inspect.getmodule(val) == module:
            functions_list.extend(get_class_methods(val))

    return functions_list


def get_class_methods(class_obj):
    """ To get the methods defined in the given class object """

    functions_list = []
    for val in class_obj.__dict__.itervalues():
        if inspect.isfunction(val) and not (val.__name__).startswith('_'):
            doc_str = inspect.getdoc(val)
            if isinstance(doc_str, str):
                doc_str = " ".join(doc_str.split())
            functions_list.append({'func_name': val.__name__.split('.')[-1],
                                   'class_name': class_obj.__name__,
                                   "doc_str": doc_str,
                                   'args_list': inspect.getargspec(val)[0]})
    return functions_list


def get_modules(package):
    """ To get the modules available in a package """

    modules_list = []
    try:
        # list the modules in the given package
        for _, modname, ispkg in pkgutil.walk_packages(
         path=package.__path__, prefix=package.__name__+'.',
         onerror=lambda x: None):
            # for nested directories(packages)
            if ispkg is False:
                real_module = importlib.import_module(modname)
                modules_list.append(real_module)
    except:
        print ("Something bad happened, please pass a valid python "
               "package as an input")

    return modules_list


def write_to_excel(functions_dict, file_path):
    """ Writes functions_dict values into an excel file """

    # create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(file_path)
    # formatting excel sheet cells
    head_format = workbook.add_format({'bold': True, 'bg_color': 'cyan',
                                       'align': 'center', 'border': True})
    content_format = workbook.add_format({'text_wrap': True, 'border': True})

    for name, value_list in functions_dict.iteritems():
        # skip the files with not valid class_methods/functions
        if not value_list:
            continue
        worksheet = workbook.add_worksheet(name=name)
        row, column = 2, 2
        worksheet.set_column(column, column, 5)
        worksheet.set_column(column+1, column+2, 25)
        worksheet.set_column(column+3, column+3, 60)
        worksheet.set_column(column+4, column+4, 25)
        worksheet.write(row, column, "S.No.", head_format)
        worksheet.write(row, column+1, "Function name", head_format)
        worksheet.write(row, column+2, "Class name", head_format)
        worksheet.write(row, column+3, "Document string", head_format)
        worksheet.write(row, column+4, "Arguments", head_format)
        serial_num = 1
        for value_dict in value_list:
            row = row + 1
            worksheet.write(row, column, serial_num, content_format)
            worksheet.write(row, column+1, value_dict['func_name'],
                            content_format)
            worksheet.write(row, column+2, value_dict['class_name'],
                            content_format)
            worksheet.write(row, column+3, value_dict['doc_str'],
                            content_format)
            worksheet.write(row, column+4, ", ".join(value_dict['args_list']),
                            content_format)
            serial_num += 1

    workbook.close()
    print "Function details successfully written to", file_path

if __name__ == '__main__':

    try:
        if len(sys.argv) > 1:
            if len(sys.argv) > 2:
                file_path = sys.argv[2]
                # if not os.path.isfile(file_path):
                #    print "Given filepath is not valid, using default file"
                #    file_path = "function_details.xlsx"
            else:
                # default output file(it will be created in the same dir)
                file_path = "function_details.xlsx"

            package = importlib.import_module(sys.argv[1])
            modules_list = get_modules(package)
            functions_dict = {}
            for module in modules_list:
                funcs_list = get_module_functions(module)
                module_name = module.__name__.split('.')[-1]
                # name of excel sheet can not be more than 32 chars
                functions_dict[module_name[:31]] = funcs_list

            write_to_excel(functions_dict, file_path)
        else:
            print ("Please provide the name of the package as "
                   "command line argument")
    except Exception as e:
        print "Execution failed: ", e
