import win32com.client
import win32clipboard as clipboard
import datetime
import pandas as pd
import re
import os
import matplotlib.pyplot as plt
import matplotlib


class OutlookToPandas:
    def __init__(self):
        # This stores the path to the current folder.
        self.script_folder = (os.path.dirname(os.path.realpath(__file__)))

        # initialize list of projects from the exel file in the current folder.
        self.projects_df = pd.ExcelFile(f"{self.script_folder}/../resources/projects.xlsx").parse('projects')
        self.projects_df['task_lev'] = '1.00'
        # set date range for outlook analysis to the current fiscal year and get all appointments in this range
        start_date = datetime.date(2019, 4, 1)
        finish_date = datetime.date(2020, 3, 31)
        appointments = self.get_all_appointments_from_outlook_date_range(start_date,finish_date)


        # Filter appointment based on sap projects and create a dataframe with only data we need.
        appointments = self.filter_appointments_by_projects(appointments)
        # Merge filtered appointments with SAP projects.
        self.tasks = pd.merge(appointments,self.projects_df)

    def filter_appointments_by_projects(self, appointments):
        # create an events container to gather what we need from each appointment that
        events_df = pd.DataFrame(columns=['username', 'start', 'project_nickname', 'subject', 'duration', 'body'])
        # get list of nickname projects
        project_nicknames = self.projects_df.project_nickname.unique()
        for a in appointments:
            # get project nickname from subject using the regex on project_nickname:task_number:subject
            # ^(.*?):((.\d*?):)?(.*)$
            # 1-project_nickname
            # 3-task_number
            # 4-subject
            # if task number not present 2 and three will be blank
            searchObj = re.search(r'^(.*?):((.\d*?):)?(.*)$', a.subject, re.I)
            if searchObj:
                project_nickname = searchObj.group(1).upper()
                project_task_number = searchObj.group(3)
                subject = searchObj.group(4)
                if project_nickname in project_nicknames:
                    events_df = events_df.append(
                        {'username': os.getlogin(),
                         'start': datetime.datetime.strptime(str(a.Start)[:19], '%Y-%m-%d %H:%M:%S'),
                         'subject': subject,
                         'project_nickname': project_nickname,
                         'project_task_number': project_task_number,
                         'duration': a.Duration,  # in minutes
                         'body': a.body
                         },
                        ignore_index=True)
        return events_df

    def get_all_appointments_from_outlook_date_range(self,start_date,finish_date):
        # Load all appointment of year into a dataframe from Outlook using win32Com object within the fiscal year
        Outlook = win32com.client.Dispatch("Outlook.Application")
        ns = Outlook.GetNamespace("MAPI")
        # Get all meetings and sort by start date and include recurrences
        appointments = ns.GetDefaultFolder(9).Items
        appointments.Sort("[Start]")
        appointments.IncludeRecurrences = "True"
        # load the start of the fiscal year.
        begin = start_date.strftime("%m/%d/%Y")
        end = finish_date.strftime("%m/%d/%Y")
        # https://docs.microsoft.com/en-ca/office/vba/api/outlook.appointmentitem.body
        appointments = appointments.Restrict("[Start] >= '" + begin + "' AND [END] <= '" + end + "'")
        return appointments

    def array_to_clipboard(self, array):
        """
        Copies an array into a string format acceptable by Excel.
        Columns separated by \t, rows separated by \n
        """
        # Create string from array
        line_strings = []
        for line in array:
            line_strings.append("\t".join(line.astype(str)).replace("\n", ""))
        array_string = "\r\n".join(line_strings)

        # Put string into clipboard (open, clear, set, close)
        clipboard.OpenClipboard()
        clipboard.EmptyClipboard()
        clipboard.SetClipboardText(array_string)
        clipboard.CloseClipboard()

    def daterange(self, start_date, end_date):
        for n in range(int((end_date - start_date).days)):
            yield start_date + datetime.timedelta(n)

    def get_date_range_from_week(self, p_year, p_week):
        firstdayofweek = datetime.datetime.strptime(f'{p_year}-W{int(p_week) - 1}-1', "%Y-W%W-%w").date()
        lastdayofweek = firstdayofweek + datetime.timedelta(days=6.9)
        return firstdayofweek, lastdayofweek

    def get_days_in_week(self, first_day):
        days_array = []
        days = [0, 1, 2, 3, 4, 5, 6]
        for day in days:
            days_array.append(first_day + datetime.timedelta(days=day))
        return days_array

    def create_week_sap_report(self, year_number, week_number):
        sap_df = self.projects_df.copy()
        project_nicknames = sap_df.project_nickname.unique()
        first_day, last_day = self.get_date_range_from_week(year_number, week_number)
        datetime_array = self.get_days_in_week(first_day)
        for date in datetime_array:
            weekday = date.strftime('%A')
            sap_df[weekday] = 0.0
            for project_nickname in project_nicknames:
                hours = self.get_hours_spent_on_project_for_date_range(project_nickname, pd.Timestamp(date),
                                                                       pd.Timestamp(date + datetime.timedelta(days=1)))
                sap_df.loc[sap_df['project_nickname'] == project_nickname, weekday] = hours
        sap_df.fillna('', inplace=True)
        values = sap_df.values
        #get rid of nickname column
        values=values[:, 1:]
        # Copy to clipboard for SAP pasting later
        self.array_to_clipboard(values)
        return sap_df.drop(columns="project_nickname")

    def get_hours_spent_on_project_for_date_range(self, project_nickname, start_date, end_date):
        # filter by date
        date_appointments = self.tasks[self.tasks.start.between(start_date, end_date)]
        # filter by project
        hours = date_appointments.loc[date_appointments['project_nickname'] == project_nickname][
                    'duration'].sum() / 60.0
        return hours

    def get_hours_spent_on_projects_for_date_range(self, start_date, end_date):
        sap_df = self.projects_df.copy()
        sap_df['duration'] = 0.0
        project_nicknames = sap_df.project_nickname.unique()
        for project_nickname in project_nicknames:
            hours = self.get_hours_spent_on_project_for_date_range(project_nickname, pd.Timestamp(start_date),
                                                                   pd.Timestamp(end_date + datetime.timedelta(days=1)))
            sap_df.loc[sap_df['project_nickname'] == project_nickname, 'duration'] = hours
        return sap_df

    def get_sap_report_for_weeks(self, year, weeknumbers):
        writer = pd.ExcelWriter(f"{self.script_folder}/timesheets.xlsx")
        for weeknumber in weeknumbers:
            sap_df = self.create_week_sap_report(year, weeknumber)
            sap_df.to_excel(writer, sheet_name=f"{weeknumber}")
        writer.save()

    def plot_bar_for_hours_on_projects_in_range(self, start, end):
        range_hours = self.get_hours_spent_on_projects_for_date_range(start, end)
        bar = matplotlib.pyplot.bar(  # using data total)arrests
            range_hours['project_nickname'],
            # with the labels being officer names
            height=range_hours['duration'],
            width=0.8,
            bottom=None, align='center', data=None)
        # View the plot
        start_date = start.strftime("%Y-%m-%d")
        end_date = start.strftime("%Y-%m-%d")
        plt.suptitle(f"Hours Spent from {start_date} to {end_date}  by Project")
        plt.show()