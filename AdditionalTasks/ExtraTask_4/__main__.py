import re
import urllib.request
import csv

adr_request = urllib.request.urlopen("https://msk.spravker.ru/avtoservisy-avtotehcentry/").read().decode()

pattern = r"(?:class=\"org-widget-header__title-link\">)(?P<name>.+)(?:<)(?:[^+]+)(?:location\">\s+)(?P<adr>.+)(?:[^+]+)(?P<phone>[^<]+)(?:<)(?:[^А-я]+)(?:.+[^А-я]+)(?P<time>.+)(?:<)"

match = re.findall(pattern, adr_request)

with open("data.csv", "w", encoding="utf-8", newline="") as file:
    writer = csv.writer(file)

    writer.writerow(["Название", "Адрес", "Телефон", "Часы работы"])

    for i in match:
        writer.writerow(i)
    print("Данные сохранены")