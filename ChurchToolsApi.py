import json
import logging
import os

import requests

from secure.secrets import users


class ChurchToolsApi:
    def __init__(self, domain):
        self.session = None
        self.domain = domain

        from secure.secrets import ct_token
        self.login_ct_rest_api(ct_token)

    def login_ct_ajax_api(self, user=list(users.keys())[0], pswd=""):
        """
        Login function using AJAX with Username and Password
        :param user: Username
        :param pswd: Password - default safed in secure.token.user dict for tests
        :return: if login successful
        """
        self.session = requests.Session()
        login_url = self.domain + '/?q=login/ajax&func=login'
        data = {'email': user, 'password': pswd}

        response = self.session.post(url=login_url, data=data)
        if response.status_code == 200:
            logging.info('Ajax User Login Successful')
            self.session.headers['CSRF-Token'] = self.get_ct_csrf_token()
            return json.loads(response.content)["status"] == 'success'
        else:
            logging.warning("Token Login failed with {}".format(response.content.decode()))
            return False

    def login_ct_rest_api(self, login_token):
        """
        Authorization: Login<token>
        If you want to authorize a request, you need to provide a Login Token as
        Authorization header in the format {Authorization: Login<token>}
        Login Tokens are generated in "Berechtigungen" of User Settings
        :param login_token: token to be used for login into CT
        :return: if login successful
        """

        self.session = requests.Session()
        login_url = self.domain + '/api/whoami'
        headers = {"Authorization": 'Login ' + login_token}
        response = self.session.get(url=login_url, headers=headers)

        if response.status_code == 200:
            response_content = json.loads(response.content)
            logging.info('Token Login Successful as {}'.format(response_content['data']['email']))
            self.session.headers['CSRF-Token'] = self.get_ct_csrf_token()
            return json.loads(response.content)['data']['id'] > 0
        else:
            logging.warning("Token Login failed with {}".format(response.content.decode()))
            return False

    def get_ct_csrf_token(self):
        """
        Requests CSRF Token https://hilfe.church.tools/wiki/0/API-CSRF
        This method was created when debugging file upload
        storing and transmitting CSRF token in headers is required for all legacy AJAX API calls unless disabled by admin
        e.g. self.session.headers['CSRF-Token'] = self.get_ct_csrf_token()
        :return: str token
        """
        url = self.domain + '/api/csrftoken'
        response = self.session.get(url=url)  # , headers=headers)
        if response.status_code == 200:
            csrf_token = json.loads(response.content)["data"]
            logging.info("CSRF Token erfolgreich abgerufen {}".format(csrf_token))
            return csrf_token
        else:
            logging.warning("CSRF Token not updated because of Response {}".format(response.content.decode()))

    def check_connection_ajax(self):
        """
        Checks whether a successful connection with the given token can be initiated using the legacy AJAX API
        :return:
        """
        url = self.domain + '/?q=churchservice/ajax&func=getAllFacts'
        headers = {
            'accept': 'application/json'
        }
        response = self.session.post(url=url, headers=headers)
        if response.status_code == 200:
            logging.debug("Response AJAX Connection successful")
            return True
        else:
            logging.debug("Response AJAX Connection failed with {}".format(json.load(response.content)))
            return False

    def get_songs(self, **kwargs):
        # song_id=None):
        """ Gets list of all songs from the server
        :key kwargs song_id: optional filter by song id
        :return: JSON Response of all songs from ChurchTools or single object
        """

        url = self.domain + '/api/songs'
        if "song_id" in kwargs.keys():
            url = url + '/{}'.format(kwargs["song_id"])
        headers = {
            'accept': 'application/json'
        }
        response = self.session.get(url=url, headers=headers)

        if response.status_code == 200:
            response_content = json.loads(response.content)
            response_data = response_content['data'].copy()
            logging.debug("First response of GET Songs successful {}".format(response_content))

            if 'meta' not in response_content.keys():  # Shortcut without Pagination
                return response_data

            # Long part extending results with pagination
            while response_content['meta']['pagination']['current'] \
                    < response_content['meta']['pagination']['lastPage']:
                logging.info("page {} of {}".format(response_content['meta']['pagination']['current'],
                                                    response_content['meta']['pagination']['lastPage']))
                params = {'page': response_content['meta']['pagination']['current'] + 1}
                response = self.session.get(url=url, headers=headers, params=params)
                response_content = json.loads(response.content)
                response_data.extend(response_content['data'])

            return response_data
        else:
            if "song_id" in kwargs.keys():
                logging.info("Did not find song ({}) with CODE {}".format(kwargs["song_id"], response.status_code))
            else:
                logging.warning("Something went wrong fetching songs: CODE {}".format(response.status_code))

    def get_song_ajax(self, song_id=None):
        """
        Legacy AJAX function to get a specific song
        used to e.g. check for tags
        :param song_id: the id of the song to be searched for
        :return: response content interpreted as json
        """

        url = self.domain + '/?q=churchservice/ajax&func=getAllSongs'

        response = self.session.post(url=url)
        song = json.loads(response.content)['data']['songs'][str(song_id)]

        return song

    def get_song_category_map(self):
        """
        Helpfer function creating requesting CT metadata for mapping of categories
        :return: a dictionary of CategoryName:CTCategoryID
        :rtype: dict
        """

        url = self.domain + '/api/event/masterdata'
        headers = {
            'accept': 'application/json'
        }
        response = self.session.get(url=url, headers=headers)
        response_content = json.loads(response.content)
        song_categories = response_content['data']['songCategories']
        song_category_dict = {}
        for item in song_categories:
            song_category_dict[item['name']] = item['id']

        return song_category_dict

    def get_groups(self, group_id=None):
        """ Gets list of all groups
        :param group_id: optional filter by group id
        :return dict mit allen Gruppen
        """
        url = self.domain + '/api/groups'
        if group_id is not None:
            url = url + '/{}'.format(group_id)
        headers = {
            'accept': 'application/json'
        }
        response = self.session.get(url=url, headers=headers)

        if response.status_code == 200:
            response_content = json.loads(response.content)
            response_data = response_content['data'].copy()
            logging.debug("First response of Groups successful {}".format(response_content))

            if 'meta' not in response_content.keys():  # Shortcut without Pagination
                return response_data

            # Long part extending results with pagination
            while response_content['meta']['pagination']['current'] \
                    < response_content['meta']['pagination']['lastPage']:
                logging.info("page {} of {}".format(response_content['meta']['pagination']['current'],
                                                    response_content['meta']['pagination']['lastPage']))
                params = {'page': response_content['meta']['pagination']['current'] + 1}
                response = self.session.get(url=url, headers=headers, params=params)
                response_content = json.loads(response.content)
                response_data.extend(response_content['data'])

            return response_data
        else:
            logging.warning("Something went wrong fetiching groups: {}".format(response.status_code))

    def file_upload(self, filename_path_for_upload, domain_type, domain_identifier,
                    custom_file_name=None,
                    overwrite=False):
        """
        Helper function to upload an attachment to any module of ChurchTools
        :param filename_path_for_upload: file to be opened e.g. with open('media/pinguin.png', 'rb')
        :param domain_type:  The domain type, currently supported are 'avatar', 'groupimage', 'logo', 'attatchments',
         'html_template', 'service', 'song_arrangement', 'importtable', 'person', 'familyavatar', 'wiki_.?'.
        :param domain_identifier: ID of the object in ChurchTools
        :param custom_file_name: optional file name - if not specified the one from the file is used
        :param overwrite: if true delete existing file before upload of new one to replace \
        it's content instead of creating a copy
        :return:
        """

        file_to_upload = open(filename_path_for_upload, 'rb')

        url = '{}/api/files/{}/{}'.format(self.domain, domain_type, domain_identifier)

        if overwrite:
            logging.debug("deleting old file before download")
            delete_file_name = file_to_upload.name.split('/')[-1] if custom_file_name is None else custom_file_name
            self.file_delete(domain_type, domain_identifier, delete_file_name)

        # add files as files form data with dict using 'files[]' as key and (tuple of filename and fileobject)
        if custom_file_name is None:
            files = {'files[]': (file_to_upload.name.split('/')[-1], file_to_upload)}
        else:
            if '/' in custom_file_name:
                logging.warning('/ in file name ({}) will fail upload!'.format(custom_file_name))
                files = {}
            else:
                files = {'files[]': (custom_file_name, file_to_upload)}

        response = self.session.post(url=url, files=files)
        file_to_upload.close()

        """
        # Issues with HEADERS in Request module when using non standard 'files[]' key in POST Request
        # Workaround for ChurchTools - generate session with /api/whoami GET request and reuse it
        # Requests module usually automatically completes required header Params e.g. Content-Type ...
        # in case manual header e.g. for AUTH is used, headers don't auto complete
        # and server rejects messsages or data is ommited 
        # Error Code 500 is also missing in API documentation
        
        headers = {'Authorization': 'Login GEHEIM'}
        response_test = requests.post(url=url, headers=headers, files=files)
        #> this fails !
        """

        if response.status_code == 200:
            try:
                response_content = json.loads(response.content)
                logging.debug("Upload successful {}".format(response_content))
            except:
                logging.warning(response.content.decode())
        else:
            logging.warning(response.content.decode())

    def file_delete(self, domain_type, domain_identifier, filename_for_selective_delete=None):
        """
        Helper function to delete ALL attachments of any specified module of ChurchTools#
        Downloads all existing files to tmp and reuploads them in case filename_for_delete is specified
        :param domain_type:  The domain type, currently supported are 'avatar', 'groupimage', 'logo', 'attatchments',
         'html_template', 'service', 'song_arrangement', 'importtable', 'person', 'familyavatar', 'wiki_.?'.
        :param domain_identifier: ID of the object in ChurchTools
        :param filename_for_selective_delete: name of the file to be deleted - all others will reupload
        :return: if successful
        """
        url = self.domain + '/api/files/{}/{}'.format(domain_type, domain_identifier)

        if filename_for_selective_delete is not None:

            response = self.session.get(url=url)
            online_files = [(item["id"], item['name'], item['fileUrl']) for item in
                            json.loads(response.content)['data']]

            for current_file in online_files:
                current_id = current_file[0]
                current_name = current_file[1]
                current_url = current_file[2]
                if current_name != filename_for_selective_delete:
                    response = self.session.get(current_url)
                    file_path_name = "tmp/{}_{}".format(current_id, current_name)
                    temp_file = open(file_path_name, "wb")
                    temp_file.write(response.content)
                    temp_file.close()

        # Delete all Files for the id online
        response = self.session.delete(url=url)

        # Recover Local Files with new upload and delete local copy from tmp
        if filename_for_selective_delete is not None:
            local_files = os.listdir('tmp')
            for current_file in local_files:
                file_name = ''.join((current_file.split("_")[1:]))
                self.file_upload('tmp/{}'.format(current_file), domain_type, domain_identifier,
                                 custom_file_name=file_name)
                os.remove('tmp/{}'.format(current_file))

        return response.status_code == 204  # success code for delete action upload

    def create_song(self, title: str, songcategory_id: int, author='', copyright='', ccli='', tonality='', bpm='',
                    beat=''):
        """
        Method to create a new song using legacy AJAX API
        Does not check for existing duplicates !
        function endpoint see https://api.church.tools/function-churchservice_addNewSong.html
        name for params reverse engineered based on web developer tools in Firefox and live churchTools instance

        :param title: Title of the Song
        :param songcategory_id: int id of site specific songcategories (created in CT Metadata) - required
        :param author: name of author or authors, ideally comma separated if multiple - optional
        :param copyright: name of organization responsible for rights distribution - optional
        :param ccli: CCLI ID see songselect.ccli.com/ - using "-" if empty on purpose - optional
        :param tonality: empty or specific string used for tonaly - see ChurchTools for details e.g. Ab,A,C,C# ... - optional
        :param bpm: Beats per Minute - optional
        :param beat: Beat - optional

        :return: int song_id: ChurchTools song_id of the Song created
        """
        url = self.domain + '/?q=churchservice/ajax&func=addNewSong'

        data = {
            'bezeichnung': title,
            'songcategory_id': songcategory_id,
            'author': author,
            'copyright': copyright,
            'ccli': ccli,
            'tonality': tonality,
            'bpm': bpm,
            'beat': beat
        }

        response = self.session.post(url=url, data=data)
        new_id = int(json.loads(response.content)['data'])
        return new_id

    def edit_song(self, song_id: int, songcategory_id=None, title=None, author=None, copyright=None, ccli=None,
                  practice_yn=None, ):
        """
        Method to EDIT an existing song using legacy AJAX API
        Changes are only applied to fields that have values in respective param
        None is considered empty while '' is an empty text which clears existing values

        function endpoint see https://api.church.tools/function-churchservice_editSong.html
        name for params reverse engineered based on web developer tools in Firefox and live churchTools instance
        NOTE - BPM and BEAT used in create are part of arrangement and not song therefore not editable in this method

        :param song_id: ChurchTools site specific song_id which should be modified - required

        :param title: Title of the Song
        :param songcategory_id: int id of site specific songcategories (created in CT Metadata)
        :param author: name of author or authors, ideally comma separated if multiple
        :param copyright: name of organization responsible for rights distribution
        :param ccli: CCLI ID see songselect.ccli.com/ - using "-" if empty on purpose
        :param practice_yn: bool as 0 and 1 - additional param which does not exist in create method!

        :return: response item
        """

        # system/churchservice/churchservice_db.php
        url = self.domain + '/?q=churchservice/ajax&func=editSong'

        existing_song = self.get_songs(song_id=song_id)

        data = {
            'id': song_id if song_id is not None else existing_song['name'],
            'bezeichnung': title if title is not None else existing_song['name'],
            'songcategory_id': songcategory_id if songcategory_id is not None else existing_song['category']['id'],
            'author': author if author is not None else existing_song['author'],
            'copyright': copyright if copyright is not None else existing_song['copyright'],
            'ccli': ccli if ccli is not None else existing_song['ccli'],
            'practice_yn': practice_yn if practice_yn is not None else existing_song['shouldPractice'],
        }

        response = self.session.post(url=url, data=data)
        return response

    def delete_song(self, song_id: int):
        """
        Method to DELETE a song using legacy AJAX API
        name for params reverse engineered based on web developer tools in Firefox and live churchTools instance

        :param song_id: ChurchTools site specific song_id which should be modified - required

        :return: response item
        """

        # system/churchservice/churchservice_db.php
        url = self.domain + '/?q=churchservice/ajax&func=deleteSong'

        data = {
            'id': song_id,
        }

        response = self.session.post(url=url, data=data)
        return response

    def add_song_tag(self, song_id: int, song_tag_id: int):
        """
        Method to add a song tag using legacy AJAX API on a specific song
        reverse engineered based on web developer tools in Firefox and live churchTools instance

        re-adding existing tag does not cause any issues
        :param song_id: ChurchTools site specific song_id which should be modified - required
        :param song_tag_id: ChurchTools site specific song_tag_id which should be added - required

        :return: response item
        """
        url = self.domain + '/?q=churchservice/ajax&func=addSongTag'

        data = {
            'id': song_id,
            'tag_id': song_tag_id
        }

        response = self.session.post(url=url, data=data)
        return response

    def remove_song_tag(self, song_id: int, song_tag_id: int):
        """
        Method to remove a song tag using legacy AJAX API on a specifc song
        reverse engineered based on web developer tools in Firefox and live churchTools instance

        re-removing existing tag does not cause any issues
        :param song_id: ChurchTools site specific song_id which should be modified - required
        :param song_tag_id: ChurchTools site specific song_tag_id which should be added - required

        :return: response item
        """
        url = self.domain + '/?q=churchservice/ajax&func=delSongTag'

        data = {
            'id': song_id,
            'tag_id': song_tag_id
        }

        response = self.session.post(url=url, data=data)
        return response

    def get_song_tags(self, song_id: int):
        """
        Method to get a song tag workaround using legacy AJAX API for getSong
        :param song_id: ChurchTools site specific song_id which should be modified - required
        :return: response item
        """
        song = self.get_song_ajax(song_id)
        return song['tags']

    def contains_song_tag(self, song_id: int, song_tag_id: int):
        """
        Helper which checks if a specifc song_tag_id is present on a song
        :param song_id: ChurchTools site specific song_id which should checked
        :param song_tag_id: ChurchTools site specific song_tag_id which should be searched for
        :return: bool if present
        """
        tags = self.get_song_tags(song_id)
        return str(song_tag_id) in tags

    def get_songs_with_tag(self, song_tag_id: int):
        """
        Helper which returns all songs that contain have a specific tag
        :param song_tag_id: ChurchTools site specific song_tag_id which should be searched for
        :return: list of songs
        """
        songs = self.get_songs()
        all_song_ids = [value['id'] for value in songs]
        filtered_song_ids = [id for id in all_song_ids if self.contains_song_tag(id, song_tag_id)]

        result = [self.get_songs(song_id=song_id) for song_id in filtered_song_ids]

        return result

    def get_events(self, from_: str = '', to: str = '', canceled: bool = False, direction: str = 'forward',
                   limit: int = 1, include: str = 'eventServices'):
        """
        Method to get all the events from given timespan or only the next event
        :param from_: starting date in format YYYY-MM-DD
        :param to: end date in format YYYY-MM-DD
        :param canceled: if 'True' canceled events will be also included
        :param direction: direction of output 'forward' or 'backward' from the date defined by parameter 'from_'
        :param limit: limits the number of events - Default = 1, if all events shall be retrieved insert 'None'
        :param include: if Parameter is set to 'eventServices', the services of the event will be included
        :return: list of events
        """
        url = self.domain + '/api/events'
        headers = {
            'accept': 'application/json'
        }

        self.session.params['canceled'] = canceled
        if from_ != '' and (from_ is not None):
            self.session.params['from'] = from_
        if to != '' and (to is not None):
            self.session.params['to'] = to
        if limit is not None:
            self.session.params['limit'] = limit
        if (include != '') and (include is not None):
            self.session.params['include'] = include
        if (direction != '') and (direction is not None):
            self.session.params['direction'] = direction

        response = self.session.get(url=url, headers=headers)

        if response.status_code == 200:
            response_content = json.loads(response.content)
            response_data = response_content['data'].copy()
            logging.debug("First response of Events successful {}".format(response_content))

            if 'meta' not in response_content.keys():  # Shortcut without Pagination
                return response_data

            if 'pagination' not in response_content['meta'].keys():
                return response_data

            # Long part extending results with pagination
            # TODO #1 copied from other method unsure if pagination works the same as with groups
            while response_content['meta']['pagination']['current'] \
                    < response_content['meta']['pagination']['lastPage']:
                logging.info("page {} of {}".format(response_content['meta']['pagination']['current'],
                                                    response_content['meta']['pagination']['lastPage']))
                params = {'page': response_content['meta']['pagination']['current'] + 1}
                response = self.session.get(url=url, headers=headers, params=params)
                response_content = json.loads(response.content)
                response_data.extend(response_content['data'])

            return response_data
        else:
            logging.warning("Something went wrong fetching events: {}".format(response.status_code))

    def get_event_agenda(self, event_id: int):
        """
        Retrieve agenda for event by ID from ChurchTools
        :return:
        """
        url = self.domain + '/api/events/{}/agenda'.format(event_id)
        headers = {
            'accept': 'application/json'
        }
        response = self.session.get(url=url, headers=headers)

        if response.status_code == 200:
            response_content = json.loads(response.content)
            response_data = response_content['data'].copy()
            logging.debug("Agenda load successful {}".format(response_content))

            return response_data
        else:
            logging.info("Event requested that does not have an agenda with status: {}".format(response.status_code))
            return None

    def file_download(self, filename: str, domain_type, domain_identifier, path_for_download='./Downloads'):
        """
        Retrieves file from ChurchTools for specific filename, domain_type and domain_identifier from churchtools
        :param filename:
        :param domain_type: Currently supported are 'avatar', 'groupimage', 'logo', 'attatchments', 'html_template', 'service', 'song_arrangement', 'importtable', 'person', 'familyavatar', 'wiki_.?'.
        :param domain_identifier = Id of Arrangement (in case of songs) This information can be checked for example by calling get_songs()
        :param path_for_download: local path as target for the download - will be created if not exists
        :return: bool if success
        """
        StateOK = False
        CHECK_FOLDER = os.path.isdir(path_for_download)

        # If folder doesn't exist, then create it.
        if not CHECK_FOLDER:
            os.makedirs(path_for_download)
            print("created folder : ", path_for_download)

        url = '{}/api/files/{}/{}'.format(self.domain, domain_type, domain_identifier)

        response = self.session.get(url=url)

        if response.status_code == 200:
            response_content = json.loads(response.content)
            arrangement_files = response_content['data'].copy()
            logging.debug("SongArrangement-Files load successful {}".format(response_content))
            fileUrl = None

            for file in arrangement_files:
                filenameoriginal = str(file['name'])
                fileUrl = str(file['fileUrl'])
                if filenameoriginal == filename:
                    break;
            if fileUrl == None:
                logging.warning("File {} does not exist".format(filename))
            else:
                logging.debug("Found File: {}".format(filename))
                # Build path OS independent
                path_file = os.sep.join([path_for_download, filename])
                StateOK = self.file_download_from_url(fileUrl, path_file)

            return StateOK
        else:
            logging.warning("Something went wrong fetching SongArrangement-Files: {}".format(response.status_code))

    def file_download_from_url(self, file_url: str, target_path: str):
        """
        Retrieves file from ChurchTools for specific file_url from churchtools
        This function is used by file_download(...)
        :param file_url: Example -> file_url=https://lgv-oe.church.tools/?q=public/filedownload&id=631&filename=738db42141baec592aa2f523169af772fd02c1d21f5acaaf0601678962d06a00
                Pay Attention: this file-url consists of a specific / random filename which was created by churchtools
        :param target_path: directory to drop the download into (must exist first)
        :return:
        """
        # NOTE the stream=True parameter below
        with self.session.get(url=file_url, stream=True) as r:
            if r.status_code == 200:
                with open(target_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        # If you have chunk encoded response uncomment if
                        # and set chunk_size parameter to None.
                        # if chunk:
                        f.write(chunk)
                logging.debug("Download of {} successful".format(file_url))
                return True
            else:
                logging.warning("Something went wrong during file_download: {}".format(r.status_code))
                return False
