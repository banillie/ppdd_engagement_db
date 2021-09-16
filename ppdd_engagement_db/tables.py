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


def p_stakeholder_table():
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


def p_engagement_table():
    cur.execute("""
        DROP TABLE IF EXISTS ProjectEngagements;
        """)
    cur.execute(
        '''CREATE TABLE ProjectEngagements(
        ID TEXT PRIMARY KEY,
        EngID TEXT,
        ProjectID TEXT,
        FOREIGN KEY (EngID) REFERENCES parent(Engagements),
        FOREIGN KEY (ProjectID) REFERENCES parent(Projects)
        )'''
    )


def e_stakeholders_table():
    cur.execute("""
        DROP TABLE IF EXISTS EngagementStakeholders;
        """)
    cur.execute(
        '''CREATE TABLE EngagementStakeholders(
        ID TEXT PRIMARY KEY,
        EngID TEXT,
        StkhID TEXT,
        FOREIGN KEY (EngID) REFERENCES parent(Engagements),
        FOREIGN KEY (StkhID) REFERENCES parent(Stakeholders)
        )'''
    )


def e_ppdd_table():
    cur.execute("""
        DROP TABLE IF EXISTS EngagementPPDDs;
        """)
    cur.execute(
        '''CREATE TABLE EngagementPPDDs(
        ID TEXT PRIMARY KEY,
        EngID TEXT,
        PPDDID TEXT,
        FOREIGN KEY (EngID) REFERENCES parent(Engagements),
        FOREIGN KEY (PPDDID) REFERENCES parent(PPDDs)
        )'''
    )


def engagement_types_table():
    cur.execute("""
        DROP TABLE IF EXISTS EngagementTypes;
        """)
    cur.execute(
        '''CREATE TABLE EngagementTypes(
        ID TEXT PRIMARY KEY,
        EngID TEXT,
        EngType TEXT,
        FOREIGN KEY (EngID) REFERENCES parent(Engagements)
        )'''
    )


def engagement_ws_table():
    cur.execute("""
        DROP TABLE IF EXISTS EngagementWorkStreams;
        """)
    cur.execute(
        '''CREATE TABLE EngagementWorkStreams(
        ID TEXT PRIMARY KEY,
        EngID TEXT,
        WSType TEXT,
        FOREIGN KEY (EngID) REFERENCES parent(Engagements)
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
# ps_data = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "ProjectStakeholders", 4)
# p_engagements_data = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "ProjectEngagements", 4)
# engagements_s_data = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "EngagementStakeholders", 4)
# engagements_ppdd_data = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "EngagementPPDDs", 4)
# engagement_types = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "EngagementTypes", 4)
engagement_ws = get_data(root_path / "ppdd_engagement_db_tables.xlsx", "EngagementWorkStreams", 4)

engagement_ws_table()

enter_data(engagement_ws, "EngagementWorkStreams")
con.commit()
con.close()