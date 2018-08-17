import datetime 


class CalledMember:
    def __init__(self, name, position, organization, sustain_date, set_apart_date):
         self.name = name
         self.position = position
         self.organization = organization
         self.sustain_date = sustain_date
         self.set_apart_date = datetime.datetime.strptime(set_apart_date, "%d%m%Y").date()


    def print(self):
        print("Name: " + self.name)
        print("\tPosition: " + self.position)
        print("\tOrganization: " + self.organization)
        print("\tSustained: " + self.sustain_date)
        print("\tSet Apart: " + self.set_apart_date)
        print("\n")
