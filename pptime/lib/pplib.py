import requests
from requests_oauthlib import OAuth1 as OAuth
from urllib.parse import parse_qs
import webbrowser
import json
import re
import os

def get_env_variable(var_name):
    """
    Get the environment variable or return exception
    :param var_name: Environment Variable to lookup
    """
    try:
        return os.environ[var_name]
    except KeyError: #Dev only !
        from env_var import env_var
        return env_var[var_name]



def connect_oauth():
    """
    This function connects to Project Place API with OAuth1
    It uses environement variables.
    For development only, it can use a dict called env-var in env-var.py at the project root.
    """

    PP_BASE_URL = os.environ.get('PP_BASE_URL', get_env_variable('PP_BASE_URL'))
    CLIENT_KEY = os.environ.get('PP_CLIENT_KEY', get_env_variable('CLIENT_KEY'))
    CLIENT_SECRET = os.environ.get('PP_CLIENT_SECRET', get_env_variable('CLIENT_SECRET'))
    access_token_key = os.environ.get('PP_access_token_key', get_env_variable('access_token_key'))
    access_token_secret = os.environ.get('PP_access_token_secret', get_env_variable('access_token_secret'))

    if access_token_key is None:
        print ('Getting request token...:',)
        oauth = OAuth(CLIENT_KEY, client_secret=CLIENT_SECRET)
        r = requests.post(PP_BASE_URL + 'initiate', auth=oauth)
        credentials = parse_qs(r.content)
        request_token_key = credentials.get('oauth_token')[0].decode('ascii')
        request_token_secret = credentials.get('oauth_token_secret')[0].decode('ascii')
        print (request_token_key, 'with secret', request_token_secret)

        print ("Opening webbrowser to authenticate request token")
        webbrowser.open(PP_BASE_URL + '/authorize?oauth_token=' + request_token_key)
        oauth_verifier = raw_input('Input oauth_verifier: ')

        print ("Exchanging request token for access token")
        oauth = OAuth(CLIENT_KEY, client_secret=CLIENT_SECRET, resource_owner_key=request_token_key,
                      resource_owner_secret=request_token_secret, verifier=oauth_verifier)
        r = requests.post(PP_BASE_URL + 'token', auth=oauth)
        credentials = parse_qs(r.content)
        access_token_key = credentials.get('oauth_token')[0].decode('ascii')
        access_token_secret = credentials.get('oauth_token_secret')[0].decode('ascii')
        print ("Successfully fetch access token", access_token_key, 'with secret', access_token_secret)

    print ('Getting user profile...',)
    oauth = OAuth(CLIENT_KEY, client_secret=CLIENT_SECRET, resource_owner_key=access_token_key,
                  resource_owner_secret=access_token_secret)
    r = requests.get(url=PP_BASE_URL + '1/user/me/profile', auth=oauth)
    print (json.dumps(r.json(), sort_keys=True, indent=4, separators=(',', ': ')))
    return(oauth)


def get_projects_dict(oauth):
    """
    To get all the projects that the user can access.
    If you connect with a robot account it will display all the projects of the account.

    :param oauth: OAuth object with credentials
    :return: a dict with all the projects
    """

    PP_BASE_URL = os.environ.get('PP_BASE_URL', get_env_variable('PP_BASE_URL'))

    projects_json = requests.get(url=PP_BASE_URL + '1/account/projects', auth=oauth).json()
    projects_dict = {}
    for project in projects_json['projects']:
        projects_dict[project['id']] = project['name']
    return projects_dict


def get_project_name(project_id,projects_dict):
    '''
    Extract a project name from a projects dict

    :param project_id: the string ID of the project
    :param projects_dict: the dict of all projects retrieved by get_projects_dict()
    :return: a project name
    '''
    if project_id in projects_dict :
      project_name = projects_dict[project_id]
    else:
      project_name = 'unknown project'
    return project_name


def get_users_dict(oauth):
    """
    To get all the users of the account

    :param oauth: Oauth object with credentials
    :return: a dict with all the users
    """

    PP_BASE_URL = os.environ.get('PP_BASE_URL', get_env_variable('PP_BASE_URL'))

    users_json = requests.get(url=PP_BASE_URL + '1/account/members', auth=oauth).json()
    users_dict = {}
    for user in users_json['members']:
        users_dict[user['id']] = user['name']
    return users_dict


def get_user_name(user_id,users_dict):
    """
    Extract an user name from an users dict

    :param user_id: ID of the user
    :param users_dict:  dictionary of users from the API
    :return: user name string
    """
    if user_id in users_dict :
      user_name = users_dict[user_id]
    else:
      user_name = 'unknown user'
    return user_name

def get_year(datestring):
    """
    Extract the year from a string with format: 2016xxxxxxx
    :param datestring:
    :return: int year
    """
    year = re.search('\d{4}', datestring).group(0) #PP provides format: 2016-07-11 15:48:35
    return int(year)


def make_time_clusters(oauth,projects_dict):
    """
    Sum time in reports from the same user, year and project.
    /!\ Running this function can take a while !

    :param oauth:
    :param projectid:
    :return: A dict of time clusters
    """

    PP_BASE_URL = os.environ.get('PP_BASE_URL', get_env_variable('PP_BASE_URL'))

    time_clusters = dict()

    for projectid in projects_dict:
        reports_extract = requests.get(url=PP_BASE_URL + '1/timereports/?project_ids=' + str(projectid), auth=oauth).json()

        for report in reports_extract:

            if ('projectId' in report):  # if the user don't have access to the report, the 'projectID' key will not be available.
                project_id = int(report['projectId'])
                user_id = int(report['userId'])
                year = get_year(report['reportedDate'])
                reported_hours = float(report['minutes']) / 60
                cluster_key = (project_id, user_id, year)

                if cluster_key in time_clusters:
                    time_clusters[cluster_key]['hours'] += reported_hours

                else:
                    time_clusters[cluster_key] = {'projectId': project_id,
                                                  'userId': user_id,
                                                  'year': year,
                                                  'hours': reported_hours}

    return(time_clusters)

def get_active_items(time_clusters):
    """
    Extract users, projects and years actually in time clusters
    :param time_clusters: the dict of time clusters
    :return: a dict of the 3 lists
    """
    # Initiate lists
    active_projects = []
    active_users = []
    active_years = []

    # Generate lists by looking inside time_clusters

    for cluster_key in time_clusters:

        project_id = time_clusters[cluster_key]['projectId']
        user_id = time_clusters[cluster_key]['userId']
        year = time_clusters[cluster_key]['year']

        if project_id not in active_projects:
            active_projects.append(project_id)

        if user_id not in active_users:
            active_users.append(user_id)

        if year not in active_years:
            active_years.append(year)

    # Order the list for further pretty printing
    active_projects.sort()
    active_users.sort()
    active_years.sort()

    # Concatenate everything in a dict
    active_items = {
        'active_projects' : active_projects,
        'active_users' : active_users,
        'active_years' : active_years,
    }

    return active_items

def get_all_years_time_clusters(time_clusters):

    all_years_time_clusters = {}

    for cluster_key in time_clusters:

        project_id = time_clusters[cluster_key]['projectId']
        user_id = time_clusters[cluster_key]['userId']
        hours = time_clusters[cluster_key]['hours']
        all_years_cluster_key = (project_id, user_id)

        if all_years_cluster_key in all_years_time_clusters:
            all_years_time_clusters[all_years_cluster_key]['hours'] += hours

        else:
            all_years_time_clusters[all_years_cluster_key] = {'projectId': project_id, 'userId': user_id, 'hours': hours}

    return all_years_time_clusters

