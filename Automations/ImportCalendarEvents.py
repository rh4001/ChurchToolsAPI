import logging
from ChurchToolsApi import ChurchToolsApi
from pprint import pprint
import pandas as pd
import openpyxl
from datetime import datetime

EXCEL = 'Termine.xlsx'


def excel_to_dataframe(excel_path):
    df = pd.read_excel(excel_path, skiprows=1, header=0)
    return df


def check_plausibility(df):
    # Überprüfung des Datums
    date_format = "%Y-%m-%d"
    invalid_dates = df[
        df["Datum Start"].apply(lambda x: pd.to_datetime(x, format=date_format, errors="coerce")).isnull()]
    if not invalid_dates.empty:
        logging.warning("Folgende ungültige Datumsangaben gefunden:")
        for index, row in invalid_dates.iterrows():
            logging.warning(f"Zeile {index + 3}: Titel '{row['Titel']}' Datum Start: '{row['Datum Start']}'")
        has_errors = True
    else:
        has_errors = False

    # Überprüfung des Titels
    invalid_titles = df[df["Titel"].isnull() | df["Titel"].eq("")]
    if not invalid_titles.empty:
        logging.warning("Folgende ungültige Titel gefunden:")
        for index, row in invalid_titles.iterrows():
            logging.warning(f"Zeile {index + 3}: Titel fehlt oder ist leer")
        has_errors = True

    # Überprüfung, ob "Datum Start" länger als ein Jahr her ist
    df["Datum Start"] = pd.to_datetime(df["Datum Start"], format=date_format, errors="coerce")
    one_year_ago = datetime.now() - pd.DateOffset(years=1)
    invalid_old_dates = df[df["Datum Start"] < one_year_ago]
    if not invalid_old_dates.empty:
        logging.error("Folgende Datumsangaben sind länger als ein Jahr her:")
        for index, row in invalid_old_dates.iterrows():
            logging.error(f"Zeile {index + 3}: Titel '{row['Titel']}' Datum Start: '{row['Datum Start']}'")
        has_errors = True

    # Überprüfung, ob "Datum Ende" nach "Datum Start" liegt
    df["Datum Ende"] = pd.to_datetime(df["Datum Ende"], format=date_format, errors="coerce")
    invalid_end_dates = df[df["Datum Ende"] < df["Datum Start"]]
    if not invalid_end_dates.empty:
        logging.error("Folgende Datumsangaben liegen vor dem Datum Start:")
        for index, row in invalid_end_dates.iterrows():
            logging.error(f"Zeile {index + 3}: Titel '{row['Titel']}' Datum Ende: '{row['Datum Ende']}'")
        has_errors = True

    # Überprüfung, ob "Uhrzeit Ende" vorhanden ist, wenn "Uhrzeit Start" vorhanden ist
    invalid_time_end = df[df["Uhrzeit Start"].notnull() & df["Uhrzeit Ende"].isnull()]
    if not invalid_time_end.empty:
        logging.error("Folgende Datumsangaben haben 'Uhrzeit Start', aber fehlende 'Uhrzeit Ende':")
        for index, row in invalid_time_end.iterrows():
            logging.error(f"Zeile {index + 3}: Titel '{row['Titel']}'")
        has_errors = True

    return not has_errors


def get_CalendarId(api, excel_path):
    cell = pd.read_excel(excel_path, header=None).iloc[0, 2]
    print('Zelle', cell)
    if isinstance(cell, int):
        return cell
    elif isinstance(cell, str):
        calendars = api.get_AllCalendars()
        cal_id = None
        for calendar in calendars:
            if calendar['name'] == cell.strip():
                cal_id = calendar['id']
        if cal_id is not None:
            return cal_id
        else:
            logging.error(f'Kalendername "{cell}" konnte nicht gefunden werden')
    else:
        logging.error(f"Feld mit Kalendername bzw. -ID ist ungültig: '{cell}'")


def main():
    logging.getLogger().setLevel(logging.INFO)
    logging.info("Started calendar-import")

    # Read excel-file into dataframe
    df = excel_to_dataframe(EXCEL)
    # pprint(df)

    no_errors = check_plausibility(df)
    print(no_errors)

    # print('Typ', df['Datum Start'][1], type(df['Datum Start'][1]))


    # Create Session
    # TODO: Rechtekonzept überlegen, da sonst jeder überall Kalendereinträge erstellen kann
    from secure.defaults import domain as domain_temp
    from secure.secrets import ct_token
    api = ChurchToolsApi(domain_temp, ct_token)

    cal_id = get_CalendarId(api, EXCEL)
    print(cal_id)

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

    pprint(api.get_events(eventId=9828))

    pprint(api.set_appointments(
        calendarId=149,
        allDay=True,
        comment='With Sacrament',
        isInternal=False,
        subtitle='Untertitel',
        repeatId=4,
        title='Titel Test Python',
        description='Beschreibung',
        link='',
        startDate='2023-06-16T17:00:00Z',
        endDate='2023-06-16T19:00:00Z',
        campusId=0,
        address={
            'street': 'Musterstraße 1',
            'zip': '12345',
            'city': 'Musterstadt'
        }
    ))

    logging.info("finished calendar-import")


if __name__ == '__main__':
    main()
