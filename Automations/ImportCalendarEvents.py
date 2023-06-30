import logging
from ChurchToolsApi import ChurchToolsApi
from pprint import pprint


if __name__ == '__main__':
    logging.getLogger().setLevel(logging.INFO)
    logging.info("Started calendar-import")

    # Create Session
    from secure.defaults import domain as domain_temp
    from secure.secrets import ct_token
    api = ChurchToolsApi(domain_temp, ct_token)

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