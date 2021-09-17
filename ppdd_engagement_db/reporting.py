import sqlite3
from openpyxl import Workbook
from ppdd_engagement_db.data import root_path

con = sqlite3.connect(root_path / "ppdd_engagement.db")
c = con.cursor()

command_dict = {
    # Project Contacts
    1: "SELECT ProName, StkhFirstName , StkhLastName from Projects "
       "inner join ProjectStakeholders on ProjectStakeholders.ProjectID = Projects.ProjectID "
       "inner join Stakeholders on Stakeholders.stkhid = ProjectStakeholders.stkhid;",
    5: "SELECT ProName, StkhFirstName , StkhLastName from Projects "
       "left join ProjectStakeholders on ProjectStakeholders.ProjectID = Projects.ProjectID "
       "left join Stakeholders on Stakeholders.stkhid = ProjectStakeholders.stkhid;",
    6: "SELECT ProName, StkhFirstName , StkhLastName from Projects "
       "left join ProjectStakeholders on ProjectStakeholders.ProjectID = Projects.ProjectID "
       "left join Stakeholders on Stakeholders.stkhid = ProjectStakeholders.stkhid "
       "WHERE ProName = 'A12 Chelmsford to A120 widening'",
    # Project Stakeholder Count
    2: "SELECT ProjectID , count(StkhID) from ProjectStakeholders GROUP by ProjectID order by "
       "count(StkhID ) DESC ;",
    # Project Stakeholder Count with zeros
    # 3: "select Projects.ProjectID, count(ProjectStakeholders.StkhID ) from Projects "
    #    "left join ProjectStakeholders on Projects.ProjectID = ProjectStakeholders.ProjectID "
    #    "GROUP BY Projects.ProjectID order by count(StkhID ) DESC;",
    4: "select ProName, count(ProjectStakeholders.StkhID ) from Projects "
       "left join ProjectStakeholders on Projects.ProjectID = ProjectStakeholders.ProjectID "
       "left join Stakeholders on Stakeholders.stkhid = ProjectStakeholders.stkhid "
       "GROUP BY ProName order by count(ProjectStakeholders.StkhID ) DESC;",
    # Stakeholder Project count
    7: "SELECT ProjectStakeholders.StkhID , StkhFirstName, StkhLastName, "
       "Count(ProjectStakeholders.StkhID) from ProjectStakeholders "
       "inner JOIN Stakeholders on Stakeholders.stkhid = ProjectStakeholders.stkhid "
       "GROUP BY ProjectStakeholders.StkhID ORDER BY count(ProjectStakeholders.StkhID) desc;",
    8: "select StkhFirstName , StkhLastname , count(ProjectStakeholders.StkhID ) from Stakeholders "
       "left join ProjectStakeholders on Stakeholders.StkhID = ProjectStakeholders.StkhID "
       "GROUP BY StkhFirstName order by count(ProjectStakeholders.StkhID ) DESC;",
    9: "select StkhFirstName , StkhLastname , ProName from Stakeholders "
       "left join ProjectStakeholders on Stakeholders.StkhID = ProjectStakeholders.StkhID "
       "left join Projects on ProjectStakeholders.ProjectID = Projects.ProjectID; ",
    # Project Engagement count
    10: "SELECT ProjectID , count(EngID) from ProjectEngagements GROUP by ProjectID order by "
       "count(EngID) DESC ;",
    11: "select ProName, count(ProjectEngagements.EngID) from Projects "
       "left join ProjectEngagements on Projects.ProjectID = ProjectEngagements.ProjectID "
       "left join Engagements on Engagements.EngID = ProjectEngagements.EngID "
       "GROUP BY ProName order by count(ProjectEngagements.EngID) DESC;",
}


def read_from_db(command):
    c.execute(command)
    data = c.fetchall()

    for d in data:
        print(d)

    return data


def put_in_excel(data):
    wb = Workbook()
    ws = wb.active
    for x, d in enumerate(data):
        print(d)
        for i, dd in enumerate(d):
            ws.cell(row=x+1, column=i+1).value = dd

    return wb


t = read_from_db(command_dict[11])
# output = put_in_excel(t)
put_in_excel(t).save(root_path / "project_engagement.xlsx")