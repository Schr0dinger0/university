def save_file(file_name, data):
    with open (f'{file_name}', 'w+', encoding="utf-8") as f:
        for i in range(0,len(data)):
            f.write(data[i]+'\n')
    return 'Data Saved'
