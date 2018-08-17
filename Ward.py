from Organization import Organization
import openpyxl as op

class Ward:
    def __init__(self, workbork_file):
        self.organizations = {} # Dictionary of organizations { "Name":Organization }
        self.load_ward(workbork_file)


    def print(self):
        for name, org in self.organizations:
            print("ORGANIZATION: " + name)
            print(org.print())
            print("--------------------------------------------------")


    def add_organization(self, name, new_org):
        self.organizations.update( {name : new_org} )


    def load_ward(self, workbork_file):
        wb = op.load_workbook('org.xlsx')
        ws = wb.active
        for row in ws.rows:
            for cell in row:
                print(cell.value)

