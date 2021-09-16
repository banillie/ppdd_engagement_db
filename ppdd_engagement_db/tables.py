import sqlite3
from ppdd_engagement_db.data import get_data, root_path

con = sqlite3.connect(root_path / "ppdd_engagement.db")
cur = con.cursor()


def stakeholders_table():
    cur.execute("""
    DROP TABLE IF EXISTS Stakeholders;  
    """)
    cur.execute(
        '''CREATE TABLE Stakeholders(
        StkhID TEXT PRIMARY KEY, 
        StkhFirstName TEXT, 
        StkhLastName TEXT,  
        StkhOrg TEXT,
        StkhGroup TEXT,
        StkhTeam TEXT,
        StkhRole TEXT,
        StkhTeleNo INTEGER)'''
    )


def projects_table():
    cur.execute("""
        DROP TABLE IF EXISTS Projects;
        """)
    cur.execute(
        '''CREATE TABLE Projects(
        ProjectID TEXT PRIMARY KEY,
        ProType TEXT,
        ProName TEXT,
        ProAbbreviation TEXT,
        ProGovernance TEXT,
        BCStage TEXT
        )'''
    )


def ppdds_table():
    cur.execute("""
        DROP TABLE IF EXISTS PPDDs;
        """)
    cur.execute(
        '''CREATE TABLE PPDDs(
        PPDDID TEXT PRIMARY KEY,
        PPDDFirstName TEXT,
        PPDDLastName TEXT,
        PPDDRole TEXT,
        PPDDTeam TEXT,
        PPDDTeleNo TEXT
        )'''
    )


def engagements_table():
    cur.execute("""
        DROP TABLE IF EXISTS Engagements;
        """)
    cur.execute(
        '''CREATE TABLE Engagements(
        EngID TEXT PRIMARY KEY,
        EngDate DATE,
        EngShortSum TEXT,
        EngLongSum TEXT,
        EngFollowUpDate TEXT
        )'''
    )


def ps_table():
    cur.execute("""
        DROP TABLE IF EXISTS ProjectStakeholders;
        """)
    cur.execute(
        '''CREATE TABLE ProjectStakeholders(
        ID TEXT PRIMARY KEY,
        StkhID TEXT,
        ProjectID TEXT,
        FOREIGN KEY (StkhID) REFERENCES parent(Stakeholders),
        FOREIGN KEY (ProjectID) REFERENCES parent(Projects)
        )'''
    )


def stakeholder_entity_table():
    cur.execute("""
            DROP TABLE IF EXISTS StakeholderEntities;
            """)
    cur.execute(
        '''CREATE TABLE StakeholderEntities(
        StkhEntID INTEGER PRIMARY KEY,
        StkhID TEXT,
        EntID TEXT,
        FOREIGN KEY (StkhID) REFERENCES parent(Stakeholder),
        FOREIGN KEY (EntID) REFERENCES parent(Entity)
        )'''
    )


def enter_data(data: list, table_name: str):
    # to generalise need to specify table
    for d in data:
        print(d)
        cur.execute(f"INSERT INTO {table_name} VALUES (?,?,?)", d)


# stakeholder_data = ppdd_data(root_path / "ppdd_engagement_db_tables.xlsx", "Stakeholders", 9)
# project_data = ppdd_data(root_path / "ppdd_engagement_db_tables.xlsx", "Projects", 7)
# ppdd_data = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "PPDDs", 7)
# engagement_data = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "Engagements", 6)
ps_data = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "ProjectStakeholders", 4)

ps_table()

enter_data(ps_data, "ProjectStakeholders")
con.commit()
con.close()