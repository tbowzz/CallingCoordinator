import datetime
import openpyxl as op
from difflib import SequenceMatcher


ORG_COL = 'A'
POSITION_COL = 'B'
NAME_COL = 'C'
SUSTAIN_COL = 'D'
SET_APART_COL = 'E'


class CallingCoordinator:
    def __init__(self, workbork_file):
        self.ward = Ward()
        self.load_ward(workbork_file)


    def print(self):
        self.ward.print_orgs()
        # self.ward.print_called_members()


    def load_ward(self, workbork_file):
        wb = op.load_workbook('org.xlsx')
        ws = wb.active
        
        begin = False
        current_organization = "Bishopric"
        
        for row in ws.rows:
            called_member = CalledMember()
            for cell in row:
                if not begin:
                    if cell.column == 'B' and str(cell.value) == "Position":
                        begin = True
                else:
                    if cell.value and self.__parse_cell__(cell.value):

                        if cell.column == ORG_COL:
                            # TODO: Fix this to get the cell display text instead of the hyperlink
                            current_organization = str(cell.internal_value)
                        elif cell.column == POSITION_COL:
                            called_member.title = str(cell.value)
                        elif cell.column == NAME_COL:
                            called_member.name = str(cell.value)
                        elif cell.column == SUSTAIN_COL:
                            called_member.set_sustain_date(str(cell.value))
                        elif cell.column == SET_APART_COL:
                            # TODO: Figure out how to parse that checkmark image
                            pass

            if called_member.name and called_member.title:
                called_member.organization = current_organization
                self.ward.add_member(called_member)

    
    def __parse_cell__(self, value):
        value = str(value)
        ignore_keywords = ["*", "None", " * custom calling", "Add Another Calling", "Count:", "Position", "Name", "Sustained", "Set Apart"]
        for each in ignore_keywords:
            if SequenceMatcher(None, value, each).ratio() >= 0.75:
                return False
        return True


class Ward:
    def __init__(self):
        self.organizations = Organizations()
        self.called_members = []


    def add_member(self, called_member):
        self.called_members.append(called_member)
        self.organizations.add_member(called_member)


    def print_called_members(self):
        print("CALLED MEMBERS:\n")
        for member in self.called_members:
            member.print()

    
    def print_orgs(self):
        self.organizations.print()


class Organizations:
    def __init__(self):
        self.org_chart = dict()


    def add_member(self, called_member):
        if called_member.organization not in self.org_chart:
            self.org_chart[called_member.organization] = set()
        self.org_chart[called_member.organization].add(called_member)


    def print(self):
        for name, members in self.org_chart.items():
            print("ORGANIZATION: " + name + "\n")
            for member in members:
                member.print()
            print("--------------------------------------------------------------------------------------\n")


class CalledMember:
    def __init__(self):
         self.name = None
         self.title = None
         self.organization = None
         self.sustain_date = None
         self.set_apart = False

    
    def set_sustain_date(self, sustain_date):
        self.sustain_date = datetime.datetime.strptime(sustain_date, "%Y-%m-%d %H:%M:%S").date()


    def print(self):
        print("Name: " + self.name)
        print("Title: " + self.title)
        print("Organization: " + self.organization)
        print("Sustained: " + str(self.sustain_date))
        print("Set Apart: " + str(self.set_apart))
        print("\n")
