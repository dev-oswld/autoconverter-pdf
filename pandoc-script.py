# To-do: Pass by arguments and batch files

import pypandoc

def show_info():
    '''
    General info related to Pandoc wrapper
    '''

    print(pypandoc.get_pandoc_version())
    print(pypandoc.get_pandoc_path())  
    print(pypandoc.get_pandoc_formats())

def convert_file():
    '''
    Function to convert an input file to the expected format
    '''

    target_name = "test"
    target_extension = ".md"
    target =  target_name + target_extension
    format = "docx"
    outcome_extension = ".docx"
    outcome = target_name + outcome_extension

    pypandoc.convert_file(source_file=target, to=format, outputfile=outcome)

convert_file()
