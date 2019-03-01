from django.conf import settings
from pptime.models import *
from pptime.lib.pplib import *
import datetime
import xlsxwriter

def make_pptimereport(comment = 'no comment provided'):
    """
    Write the report file and store it as an object
    :param comment: The user comment to store with the file
    :return: the report as a PPTimeReport object
    """

    oauth = connect_oauth()

    # Get usefull dicts and lists
    projects_dict = get_projects_dict(oauth)
    users_dict = get_users_dict(oauth)
    time_clusters = make_time_clusters(oauth, projects_dict)
    active_items = get_active_items(time_clusters)
    all_years_time_clusters = get_all_years_time_clusters(time_clusters)

    # Extract usefull lists from active items
    active_users = active_items['active_users']
    active_projects = active_items['active_projects']
    active_years = active_items['active_years']

    #### Define the headers & data for the 'All years' worksheet ####

    # Set the headers by iterating over the active users

    worksheet_1_headers = [{'header': 'Projects', 'total_string': 'Totals'}]

    for user_id in active_users:
        user_name = get_user_name(user_id,users_dict)
        to_append = {'header': user_name, 'total_function': 'sum'}
        worksheet_1_headers.append(to_append)

    worksheet_1_headers.append({'header': 'Total', 'total_function': 'sum'})

    print(worksheet_1_headers)

    # Set the data by iterating over the projects & users
    worksheet_1_data = []

    for project_id in active_projects:

        worksheet_1_data_line = ['']
        worksheet_1_data_line[0] = get_project_name(project_id,projects_dict)
        project_hours = 0.0  # total number of hours of the project

        for user_id in active_users:

            if (project_id, user_id) in all_years_time_clusters:
                hours = all_years_time_clusters[(project_id, user_id)]['hours']
                project_hours += hours

            else:
                hours = 0

            worksheet_1_data_line.append(hours)

        worksheet_1_data_line.append(project_hours)
        worksheet_1_data.append(worksheet_1_data_line)

    print(worksheet_1_data)

    #### Define the headers & data for the 'year' worksheet ####

    # Set the headers by iterating over the active users

    def yearly_header(year):

        worksheet_year_headers = [{'header': 'Projects', 'total_string': 'Totals'}]

        for user_id in active_users:
            user_name = get_user_name(user_id,users_dict)
            to_append = {'header': user_name, 'total_function': 'sum'}
            worksheet_year_headers.append(to_append)

        worksheet_year_headers.append({'header': 'Total', 'total_function': 'sum'})

        return worksheet_year_headers

    print(yearly_header(2018))

    # Set the data by iterating over the projects & years

    def yearly_data(year):

        worksheet_year_data = []

        for project_id in active_projects:

            worksheet_year_data_line = ['']
            worksheet_year_data_line[0] = get_project_name(project_id, projects_dict)
            project_hours = 0.0  # total number of hours of the project

            for user_id in active_users:

                if (project_id, user_id, year) in time_clusters:
                    hours = time_clusters[(project_id, user_id, year)]['hours']
                    project_hours += hours

                else:
                    hours = 0

                worksheet_year_data_line.append(hours)

            worksheet_year_data_line.append(project_hours)
            worksheet_year_data.append(worksheet_year_data_line)

        return worksheet_year_data

    print(yearly_data(2018))

    #### Define the headers & data for the 'All time reports' worksheet ####

    # Set the headers

    worksheet_n_headers = [{'header': 'Project'}, {'header': 'User'}, {'header': 'Year'}, {'header': 'Hours'}]
    print(worksheet_n_headers)

    # Set the data by iterating over time_clusters
    worksheet_n_data = []
    worksheet_n_data_line = []

    for ikey in time_clusters:
        project = get_project_name(time_clusters[ikey]['projectId'],projects_dict)  # get name instead of ID
        user = get_user_name(time_clusters[ikey]['userId'],users_dict)  # get name instead of ID
        year = time_clusters[ikey]['year']
        hours = time_clusters[ikey]['hours']
        worksheet_n_data_line = [project, user, year, hours]
        worksheet_n_data.append(worksheet_n_data_line)

    print(worksheet_n_data)

    # Create a path with a name containing today's date
    today_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    report_name = 'time_report_' + today_date + '.xlsx'
    upload_subdir = 'timereports'
    report_absolute_path = os.path.join(settings.MEDIA_ROOT, upload_subdir, report_name)

    # Launch the writer on file
    workbook = xlsxwriter.Workbook(report_absolute_path)

    # Create a worksheet of user's time per project for all years
    worksheet_1 = workbook.add_worksheet('All years')
    number_of_rows = len(worksheet_1_data) + 1
    number_of_lines = len(worksheet_1_data[1]) - 1
    worksheet_1.add_table(0, 0, number_of_rows, number_of_lines,
                          {'data': worksheet_1_data, 'columns': worksheet_1_headers, 'total_row': True})

    # Create a worksheet of user's time per project for all years
    for year in active_years:
        worksheet_year_headers = yearly_header(2018)
        worksheet_year_data = yearly_data(2018)

        worksheet_year = workbook.add_worksheet(str(year))
        number_of_rows = len(worksheet_year_data) + 1
        number_of_lines = len(worksheet_year_data[1]) - 1
        worksheet_year.add_table(0, 0, number_of_rows, number_of_lines,
                                 {'data': worksheet_year_data, 'columns': worksheet_year_headers, 'total_row': True})

    # Create a woksheet containing all time_clusters
    worksheet_n = workbook.add_worksheet('All time reports')
    number_of_rows = len(worksheet_n_data)
    worksheet_n.add_table(0, 0, number_of_rows, 3, {'data': worksheet_n_data, 'columns': worksheet_n_headers})

    workbook.close()
    #f.close()

    # Create a new Document object called my_doc
    pptimereport = PPTimeReport()

    # Settle my_doc object path so that it is the f file
    pptimereport.docfile.name = os.path.join(upload_subdir, report_name)
    pptimereport.comment = comment

    # Save my_doc object in the database
    pptimereport.save()

    return pptimereport




def make_testreport(comment = 'no comment provided'):
    """
    Use this function for test purposes. It is faster than actually connecting to Project Place.
    :param comment: User comment for the file
    :return: the report as a PPTimeReport object
    """
    # Create a path with a name containing today's date
    today_date = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    report_name = 'test_report_' + today_date + '.xlsx'
    upload_subdir = 'timereports'
    report_absolute_path = os.path.join(settings.MEDIA_ROOT, upload_subdir, report_name)

    # Launch the writer on file
    workbook = xlsxwriter.Workbook(report_absolute_path)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Hello, world!')
    workbook.close()

    # Create a new Document object called my_doc
    testreport = PPTimeReport()

    # Settle my_doc object path so that it is the file
    testreport.comment = comment
    testreport.docfile.name = os.path.join(upload_subdir, report_name)

    # Save my_doc object in the database
    testreport.save()

    return testreport