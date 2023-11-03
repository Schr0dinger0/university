import re
def read_file(file_name):
    with open (f'{file_name}', encoding="utf-8") as f:
        data = re.sub("[^0-9А-яA-z' ']", "", (f.read()).lower()).split()
        return(sorted(data))