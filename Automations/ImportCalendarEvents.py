import logging
from ChurchToolsApi import ChurchToolsApi
from pprint import pprint
import pandas as pd
import numpy as np
import openpyxl
import re
from datetime import datetime, time, timedelta
import pytz
from dateutil.relativedelta import relativedelta

EXCEL = 'Gruppe Huck.xlsx'


def read_excel_data(excel_path):
    # Termine einlesen
    df = pd.read_excel(excel_path, sheet_name='Termine', skiprows=0, header=0)

    # Metadaten einlesen und als dict mit neuen Werten zurückgeben
    new_keys = ['Kalender', 'Gruppe', 'Automatischer Ort', 'Tage im Voraus']
    metadata = pd.read_excel(excel_path, sheet_name='Metadaten', header=None)
    if len(metadata) == len(new_keys):
        metadata = metadata.set_index(metadata.columns[0]).to_dict()[metadata.columns[1]]
        metadata = dict(zip(new_keys, metadata.values()))
    else:
        logging.warning('Bei den Metadaten wurde etwas hinzugefügt oder entfernt.')
        exit(1)

    try:
        # todo Aliasse richtig einlesen
        # aliases = pd.read_excel(excel_path, sheet_name='Aliasse', skiprows=0, header=0)
        aliases_df = pd.read_excel(excel_path, sheet_name='Aliasse')

        # Create an empty list for the nested dictionaries
        result = []

        # Iterate through each row of the table
        for index, row in aliases_df.iterrows():
            # Create an empty dictionary for this row
            row_dict = {}
            # Add the location and name from the first two columns
            row_dict['Ort'] = row['Ort, wenn nicht in CT vorhanden']
            row_dict['Name'] = row['Richtiger Name in CT']
            # Create an empty list for the aliases
            aliases_list = []
            # Iterate through the remaining columns and add all aliases
            for col in aliases_df.columns[2:]:
                if pd.notnull(row[col]):
                    aliases_list.append(row[col])
            # Add the list of aliases to the dictionary
            row_dict['Aliasse'] = aliases_list
            # Add the dictionary to the result list
            result.append(row_dict)

        aliases = result
    except:
        aliases = None

    return df, metadata, aliases


def check_plausibility(df):
    # Überprüfung des Datums
    date_format = "%Y-%m-%d"
    invalid_dates = df[
        df["Datum Start"].apply(lambda x: pd.to_datetime(x, format=date_format, errors="coerce")).isnull()]
    if not invalid_dates.empty:
        logging.warning("Folgende ungültige Datumsangaben gefunden:")
        for index, row in invalid_dates.iterrows():
            logging.warning(f"Zeile {index + 2}: Titel '{row['Titel']}' Datum Start: '{row['Datum Start']}'")
        has_errors = True
    else:
        has_errors = False

    # Überprüfung des Titels
    invalid_titles = df[df["Titel"].isnull() | df["Titel"].eq("")]
    if not invalid_titles.empty:
        logging.warning("Folgende ungültige Titel gefunden:")
        for index, row in invalid_titles.iterrows():
            logging.warning(f"Zeile {index + 2}: Titel fehlt oder ist leer")
        has_errors = True

    # Überprüfung, ob "Datum Start" länger als ein Jahr her ist
    df["Datum Start"] = pd.to_datetime(df["Datum Start"], format=date_format, errors="coerce")
    one_year_ago = datetime.now() - pd.DateOffset(years=1)
    invalid_old_dates = df[df["Datum Start"] < one_year_ago]
    if not invalid_old_dates.empty:
        logging.error("Folgende Datumsangaben sind länger als ein Jahr her:")
        for index, row in invalid_old_dates.iterrows():
            logging.error(f"Zeile {index + 2}: Titel '{row['Titel']}' Datum Start: '{row['Datum Start']}'")
        has_errors = True

    # Überprüfung, ob "Datum Ende" nach "Datum Start" liegt
    df["Datum Ende"] = pd.to_datetime(df["Datum Ende"], format=date_format, errors="coerce")
    invalid_end_dates = df[df["Datum Ende"] < df["Datum Start"]]
    if not invalid_end_dates.empty:
        logging.error("Folgende Datumsangaben liegen vor dem Datum Start:")
        for index, row in invalid_end_dates.iterrows():
            logging.error(f"Zeile {index + 2}: Titel '{row['Titel']}' Datum Ende: '{row['Datum Ende']}'")
        has_errors = True

    # Überprüfung, ob "Uhrzeit Ende" vorhanden ist, wenn "Uhrzeit Start" vorhanden ist
    invalid_time_end = df[df["Uhrzeit Start"].notnull() & df["Uhrzeit Ende"].isnull()]
    if not invalid_time_end.empty:
        logging.error("Folgende Datumsangaben haben 'Uhrzeit Start', aber fehlende 'Uhrzeit Ende':")
        for index, row in invalid_time_end.iterrows():
            logging.error(f"Zeile {index + 2}: Titel '{row['Titel']}'")
        has_errors = True

    return not has_errors


def get_CalendarId(api, cal_id_or_name):
    cell = cal_id_or_name

    calendars = api.get_AllCalendars()
    cal_id = None

    if isinstance(cell, int) or isinstance(cell, float):
        cell = int(cell)
        for calendar in calendars:
            if calendar['id'] == cell:
                cal_id = cell
        if cal_id is not None:
            return cal_id
        else:
            logging.error(f'Kalender-ID "{cell}" konnte nicht gefunden werden')
            exit(1)
    elif isinstance(cell, str):
        for calendar in calendars:
            if calendar['name'] == cell.strip():
                cal_id = calendar['id']
        if cal_id is not None:
            return cal_id
        else:
            logging.error(f'Kalendername "{cell}" konnte nicht gefunden werden')
            exit(1)
    else:
        logging.error(f"Feld mit Kalendername bzw. -ID ist ungültig: '{cell}'")
        exit(1)


def convert_to_german_time(date, time):
    # Zeitzone für Berlin definieren
    berlin_tz = pytz.timezone('Europe/Berlin')

    # Kombiniere das Datum mit der gegebenen Uhrzeit
    combined_datetime = datetime.combine(date, time)

    # Konvertiere in die Berliner Zeitzone
    berlin_datetime = berlin_tz.localize(combined_datetime)

    # Formatieren und UTC-Zeit anzeigen
    berlin_datetime_utc = berlin_datetime.astimezone(pytz.UTC)
    formatted_datetime = berlin_datetime_utc.strftime('%Y-%m-%dT%H:%M:%SZ')

    return formatted_datetime


def parse_address(address: str):
    name, street, plz, city = None, None, None, None
    address_parts = address.split(',')
    address_parts = [part.strip() for part in address_parts]

    if len(address_parts) == 3:
        name, street, plz_city = address_parts
        plz_city_parts = plz_city.split()
        if len(plz_city_parts) == 2:
            plz, city = plz_city_parts
    elif len(address_parts) == 2:
        if re.match(r'\d{5}', address_parts[1].split()[0]):
            street, plz_city = address_parts
            plz_city_parts = plz_city.split()
            if len(plz_city_parts) == 2:
                plz, city = plz_city_parts
        else:
            first_part, second_part = address_parts
            if re.match(r'.*\d+', first_part):
                street, city = address_parts
            else:
                name, city = address_parts
    elif len(address_parts) == 1:
        name_or_street = address_parts[0]
        if re.match(r'.*\d+', name_or_street):
            street = name_or_street
        else:
            name = name_or_street

    return {'Name': name, 'Straße': street, 'PLZ': plz, 'Stadt': city}


def address_test_function():
    # Test cases
    addresses = [
        "Moes Bar, Simpsonstr. 43, 32436 Springfield",
        "Musterstraße 34, 53476 Musterstadt",
        "Waldweg, 45754 Timbuktu",
        "Himmelstraße 43, Himmelstadt",
        "Freie Kirche Musterstadt",
        "Freie Kirche Musterstadt, Musterstadt",
        "Hauptstraße 12, 12345 Berlin",
        "Bäckerei Müller, Marktplatz 3, 54321 Hamburg",
        "Schlossallee 1, Schlossstadt",
        "Am Flussufer, 67890 Flussstadt",
        "Gartenweg 5, Gartenstadt",
        "Berggipfel, 13579 Bergdorf",
        "Seepromenade, Seestadt",
        "Waldhütte, Tief im Wald",
        "Burg Drachenfels, Drachenweg 1, 24680 Drachenstadt",
        "Am Strand, Strandstadt",
        "Landrat Belli Haus, Ostenwalder Str. 97, 48477 Hörstel",
        "Von-Arnim-Straße 15, 32791 Lage"
    ]

    for address in addresses:
        print('\n' + address)
        parsed = parse_address(address)
        print(parsed)


def get_adress_based_on_name():
    print('TODO')
    # TODO: erst Aliasse auslesen

import string
def find_column_letter(df: pd.DataFrame, column_name: str) -> str:
    column_position = df.columns.get_loc(column_name) + 1
    column_letter = string.ascii_uppercase[column_position - 1]
    return column_letter


def main():
    logging.getLogger().setLevel(logging.INFO)
    logging.info("Started calendar-import")

    # Read excel-file into dataframe and check for plausibility
    df, metadata, aliases = read_excel_data(EXCEL)
    # pprint(metadata)
    # print('\n\n')
    # pprint(aliases)

    excel_ok = check_plausibility(df)
    if not excel_ok:
        logging.error(f'In der Exceldatei passt etwas nicht. Bitte Logging-Datei prüfen')
        exit(1)

    # Create Session
    # TODO: Rechtekonzept überlegen, da sonst jeder überall Kalendereinträge erstellen kann
    from secure.defaults import domain as domain_temp
    from secure.secrets import ct_token
    api = ChurchToolsApi(domain_temp, ct_token)

    # Get calendar-ID from Excel
    cal_id = get_CalendarId(api, metadata['Kalender'])
    # print('CalID', cal_id)

    # Get all calendar-events between the earliest start- and latest enddate from excel
    earliest_date = df['Datum Start'].min().strftime('%Y-%m-%d')
    latest_date = ( datetime.now() + timedelta(days=metadata['Tage im Voraus']) ).strftime('%Y-%m-%d')
    print(earliest_date)
    print(latest_date)

    # Find column-letter of column 'EventID'
    column_letter = find_column_letter(df, 'EventID')

    pprint(api.get_appointments(calendarId=149, startDate=earliest_date, endDate=latest_date))

    exit()

    # Iterate over dataframe
    for index, row in df.iterrows():

        if row['Beschreibung'] != 'testzeile':
            continue
        row['Beschreibung'] = ''

        # Replace empty values with appropriate ones
        if isinstance(row['Uhrzeit Start'], float):
            row['Uhrzeit Start'] = time(0, 0)
        if isinstance(row['Uhrzeit Ende'], float):
            row['Uhrzeit Ende'] = time(0, 0)
        if row['Datum Ende'] is pd.NaT:
            row['Datum Ende'] = row['Datum Start']

        # Get timezone-corrected datetime-strings
        start_date = convert_to_german_time(row['Datum Start'], row['Uhrzeit Start'])
        end_date = convert_to_german_time(row['Datum Ende'], row['Uhrzeit Ende'])

        print(start_date,' bis ', end_date)
        pprint(row)

        continue

        pprint(api.set_appointment(
            calendarId=cal_id,
            isInternal=False,
            # repeatId=4,
            title=row['Titel'],
            description=row['Beschreibung'],
            startDate=start_date,
            endDate=end_date,
            campusId=0,
            address={
                    "city": "city",
                    # "country": {
                    #   "emoji": "emoji",
                    #   "id": 0,
                    #   "iso2": "iso2",
                    #   "name": "countryname",
                    #   "nameTranslated": "countryname translated"
                    # },
                    "district": "district",
                    "latitude": "lat",
                    "longitude": "longitude",
                    "meetingAt": "Neue Kirche",
                    "street": "street",
                    "zip": "12348"
            },
            # eventId=10165
        ))

        exit()


        # pprint(type( row['Uhrzeit Start']))




    # if not excel_ok or not cal_id_ok:
    exit()





    # pprint(api.who_am_i())
    # pprint(api.get_persons(isArchived=False, ids=[1,29]))

    # api.set_appointments(
    #     calendarId=149,
    #     allDay=True,
    #     comment='Kommentar',
    #     startDate='2023-06-14',
    #     endDate='2023-06-16',
    #     title='Titel'
    # )

    # pprint(api.get_events(eventId=9828))

    pprint(api.set_appointment(
        calendarId=149,
        allDay=True,
        comment='Kommentar jnfwef',
        isInternal=False,
        subtitle='Untertitel neu',
        repeatId=4,
        title='Titel wurde geändert',
        description='Neue Beschreibung',
        link='',
        startDate='2023-08-16T16:00:00Z',
        endDate='2023-08-16T19:00:00Z',
        campusId=0,
        address={
            'street': 'Musterstraße 45',
            'zip': '12345',
            'city': 'Musterstadt'
        },
        eventId=10121
    ))

    logging.info("finished calendar-import")


if __name__ == '__main__':
    main()
