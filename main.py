import re
from fractions import Fraction
import requests
from bs4 import BeautifulSoup
import openpyxl

workbook = openpyxl.load_workbook('DRILL CABINET 2.xlsx')
sheet = workbook['JOBBER LENGTH DRILLS']

def fraction_to_decimal(fraction):
    # Check if the fraction is in the 'x/x' format
    if re.match(r'\d+/\d+', fraction):
        # If it is, split the fraction into the numerator and denominator
        numerator, denominator = fraction.split('/')
        # Convert the numerator and denominator to integers
        numerator = int(numerator)
        denominator = int(denominator)
        # Return the decimal representation of the fraction
        return round(numerator / denominator, 3)

    # If the fraction is not in the 'x/x' format, split it on the '-' character
    whole_number, fraction = fraction.split('-')
    # Convert the whole number to an integer
    whole_number = int(whole_number)
    # Split the fraction into the numerator and denominator
    numerator, denominator = fraction.split('/')
    # Convert the numerator and denominator to integers
    numerator = int(numerator)
    denominator = int(denominator)
    # Return the decimal representation of the fraction
    return round(whole_number + numerator / denominator, 3)

rownum = 2

while rownum < 120:

    rownum_str = str(rownum)

    ptdnum = sheet['C'+rownum_str].value

    URL = "https://www.marssupply.com/Product/"+ptdnum

    page = requests.get(URL)

    soup = BeautifulSoup(page.text, "html.parser")

    #print(page.text)

    results = str(soup.find_all("meta", attrs={"name": "description"}))

    matchoal = re.search(r'\d+-\d+/\d+ in Overall Length', results)

    matchfl = re.search(r'\d+-\d+/\d+ in Flute Length|\d+/\d+ in Flute Length', results)

    matchdia = re.search(r'\d\.\d* in Drill', results)

    if matchoal:
        overall_length = matchoal.group()
        overall_length = overall_length.replace(' in Overall Length', '')
        overall_length = fraction_to_decimal(overall_length)
        print(overall_length)

    if matchfl:
        flute_length = matchfl.group()
        flute_length = flute_length.replace(' in Flute Length', '')
        flute_length = fraction_to_decimal(flute_length)
        print(flute_length)

    if matchdia:
        drill_dia = matchdia.group()
        drill_dia = drill_dia.replace(' in Drill', '')
        drill_dia = float(drill_dia)
        print(drill_dia)

    sheet['F1'] = "OAL"

    sheet['G1'] = "FluteLength"

    sheet['H1'] = "DrillDia"

    sheet['F'+rownum_str] = overall_length

    sheet['G'+rownum_str] = flute_length

    sheet['H'+rownum_str] = drill_dia

    rownum += 1

workbook.save('DRILL CABINET 1.xlsx')