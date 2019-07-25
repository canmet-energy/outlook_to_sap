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
        # get script folder
        self.script_folder = (os.path.dirname(os.path.realpath(__file__)))
        pd.set_option('display.max_columns', 500)
        # set date range for outlook analysis
        self.start_date = datetime.date(2019, 4, 1)
        self.finish_date = datetime.date(2020, 3, 31)
        # initialize list of projects
        self.projects_df = pd.ExcelFile(f"{self.script_folder}/projects.xlsx").parse('projects')

        # Load all appointment of year into a df
        Outlook = win32com.client.Dispatch("Outlook.Application")
        ns = Outlook.GetNamespace("MAPI")
        # Get all meetings and sort by start date and include recurrences
        self.appointments = ns.GetDefaultFolder(9).Items
        self.appointments.Sort("[Start]")
        self.appointments.IncludeRecurrences = "True"
        # load the start of the fiscal year.
        begin = self.start_date.strftime("%m/%d/%Y")
        end = self.finish_date.strftime("%m/%d/%Y")
        # https://docs.microsoft.com/en-ca/office/vba/api/outlook.appointmentitem.body
        self.appointments = self.appointments.Restrict("[Start] >= '" + begin + "' AND [END] <= '" + end + "'")
        self.events_df = pd.DataFrame(columns=['start', 'project_nickname', 'subject', 'duration', 'body'])
        # get list of nickname projects
        project_nicknames = self.projects_df.project_nickname.unique()
        for a in self.appointments:
            # get project nickname from subject using the regex on ADMIN:999:this stuff
            # ^(.*?):((.\d*?):)?(.*)$
            # 1-ADMIN
            # 2-999:
            # 3-999
            # 4-this stuff
            # if task number not present 2 and three will be blank
            searchObj = re.search(r'^(.*?):((.\d*?):)?(.*)$', a.subject, re.I)
            if searchObj:
                project_nickname = searchObj.group(1).upper()
                project_task_number = searchObj.group(3)
                subject = searchObj.group(4)
                if project_nickname in project_nicknames:
                    self.events_df = self.events_df.append(
                        {'start': datetime.datetime.strptime(str(a.Start)[:19], '%Y-%m-%d %H:%M:%S'),
                         'subject': subject,
                         'project_nickname': project_nickname,
                         'project_task_number': project_task_number,
                         'duration': a.Duration,  # in minutes
                         'body': a.body
                         },
                        ignore_index=True)
        self.joined_database = pd.merge(self.events_df, self.projects_df)

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

    def getDateRangeFromWeek(self, p_year, p_week):
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
        first_day, last_day = self.getDateRangeFromWeek(year_number, week_number)
        datetime_array = self.get_days_in_week(first_day)
        for date in datetime_array:
            weekday = date.strftime('%A')
            sap_df[weekday] = 0.0
            for project_nickname in project_nicknames:
                hours = self.get_hours_spent_on_project_for_date_range(project_nickname, pd.Timestamp(date),
                                                                        pd.Timestamp(date + datetime.timedelta(days=1)))
                sap_df.loc[sap_df['project_nickname'] == project_nickname, weekday] = hours
        sap_df.fillna('',inplace=True)
        values = sap_df.values
        # Copy to clipboard for SAP pasting later
        self.array_to_clipboard(values)
        return sap_df

    def get_hours_spent_on_project_for_date_range(self, project_nickname, start_date, end_date):
        # filter by date
        date_appointments = self.joined_database[ self.joined_database.start.between(start_date, end_date) ]
        # filter by project
        hours = date_appointments.loc[date_appointments['project_nickname'] == project_nickname]['duration'].sum() / 60.0
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
        
    def plot_bar_for_hours_on_projects_in_range(self,start,end):
        range_hours = self.get_hours_spent_on_projects_for_date_range(start, end)
        bar = matplotlib.pyplot.bar(    # using data total)arrests
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
        






