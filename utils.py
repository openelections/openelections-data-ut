import os
import glob
import csv

year = '2022'
election = '20221108'
path = election+'*precinct.csv'
output_file = election+'__ut__general__precinct.csv'

def generate_headers(year, path):
    os.chdir(year)
    vote_headers = []
    for fname in glob.glob(path):
        with open(fname, "r") as csvfile:
            reader = csv.reader(csvfile)
            headers = next(reader)
            print(list(fname + ': ' + h for h in headers if h not in ['county','precinct', 'office', 'district', 'candidate', 'party']))
            #vote_headers.append(h for h in headers if h not in ['county','precinct', 'office', 'district', 'candidate', 'party'])
#    with open('vote_headers.csv', "w") as csv_outfile:
#        outfile = csv.writer(csv_outfile)
#        outfile.writerows(vote_headers)

def generate_offices(year, path):
    os.chdir(year)
    offices = []
    for fname in glob.glob(path):
        with open(fname, "r") as csvfile:
            print(fname)
            reader = csv.DictReader(csvfile)
            for row in reader:
                if not row['office'] in offices:
                    offices.append(row['office'])
    with open('offices.csv', "w") as csv_outfile:
        outfile = csv.writer(csv_outfile)
        outfile.writerows(offices)

def generate_consolidated_file(year, path, output_file):
    results = []
    os.chdir(year)
    for fname in glob.glob(path):
        print(fname)
        with open(fname, "r") as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if 'election_day' in row:
                    election_day = row['election_day']
                else:
                    election_day = None
                if 'mail' in row:
                    mail = row['mail']
                else:
                    mail = None
                if 'absentee' in row:
                    absentee = row['absentee']
                else:
                    absentee = None
                if 'provisional' in row:
                    provisional = row['provisional']
                else:
                    provisional = None
                results.append([row['county'], row['precinct'], row['office'], row['district'], row['candidate'], row['party'], row['votes'], election_day, mail, absentee, provisional])

    with open(output_file, "w") as csv_outfile:
        outfile = csv.writer(csv_outfile)
        outfile.writerow(['county','precinct', 'office', 'district', 'candidate', 'party', 'votes', 'election_day', 'mail', 'absentee', 'provisional'])
        outfile.writerows(results)
