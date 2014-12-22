import csv
with open('data.csv', 'wb') as f:
    writer = csv.writer(f)
    writer.writerow(['C1','C2','C3'])
    lines = [range(3) for i in range(5)]
    for line in lines:
        writer.writerow(line)


with open('data.csv', 'rb') as f:
    reader = csv.reader(f, delimiter=',', quoting=csv.QUOTE_NONE)
    for row in reader:
        print row
