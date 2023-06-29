import logging
import json
from ChurchToolsApi import ChurchToolsApi
from pprint import pprint
import itertools

# Get CT-parameters for cCampusID, sexID...
f = open('CT_Parameter.json', encoding='utf-8')
CT_PARAMETER = json.load(f)
f.close()


def group_persons_by_household(persons):
    # Sortiere die Personenliste nach Stadt, Adresse und Nachname
    sorted_persons = sorted(persons, key=lambda p: (p['city'], p['street'], p['lastName']))

    # pprint(sorted_persons[:3])

    households = []
    current_household = []

    current_address = None

    # Iteriere über die sortierte Liste und gruppieren Personen mit gleicher Stadt, Adresse und Nachname
    for key, group in itertools.groupby(sorted_persons, key=lambda p: (p['city'], p['street'], p['lastName'])):
        household = list(group)

        print(household)
        current_household.extend(household)

        # Überprüfe, ob die nächste Person eine andere Adresse oder einen anderen Nachnamen hat
        # Wenn ja, füge den aktuellen Haushalt zu den Haushalten hinzu und starte einen neuen Haushalt
        if len(current_household) > 1 and key != current_address:
            households.append(current_household[:-1])
            current_household = [current_household[-1]]
            current_address = key

    # Füge den letzten Haushalt hinzu
    if current_household:
        households.append(current_household)

    return households


if __name__ == '__main__':
    logging.getLogger().setLevel(logging.INFO)
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info("Started generating phonebook")

    # Create Session
    from secure.defaults import domain as domain_temp
    from secure.secrets import ct_token
    api = ChurchToolsApi(domain_temp, ct_token)

    all_persons = api.get_persons()

    lemgo = list(filter(
        lambda i: i['statusId'] == CT_PARAMETER['Status']['Mitglied'] and i['campusId'] == CT_PARAMETER['Standort'][
            'Lemgo'], all_persons))

    hameln = list(filter(
        lambda i: i['statusId'] == CT_PARAMETER['Status']['Mitglied'] and i['campusId'] == CT_PARAMETER['Standort'][
            'Hameln'], all_persons))

    # pprint(all_persons[0:5])

    lemgo_households = group_persons_by_household(lemgo)

    pprint(lemgo_households[0])

    logging.info(f"Es wurden {len(lemgo_households)} Haushalte gefunden.")