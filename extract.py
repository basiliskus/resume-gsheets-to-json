import ast
import json
import gspread
from distutils.util import strtobool
from google.oauth2.service_account import Credentials


keymap = {
    "Name": ["basics.name", "{0}"],
    "Title": ["basics.label", "{0}"],
    "Location": ["basics.location", "{ \"address\": \"\", \"postalCode\": \"\", \"city\": \"{0}\", \"countryCode\": \"\", \"region\": \"\" }"],
    "Phone": ["basics.phone", "{0}"],
    "Email": ["basics.email", "{0}"],
    "Summary": ["basics.summary", "{0}"],
    "Social Media": ["basics.profiles", "{ \"network\": \"{0}\", \"username\": \"{1}\", \"url\": \"\" }"],
    "Skills": ["skills", "{ \"name\": \"{0}\", \"level\": \"\", \"keywords\": [\"{1}\"] }"],
    "Languages": ["languages", "{ \"language\": \"{0}\", \"fluency\": \"{1}\" }"]
}

work_keymap = {
    "Start Date"
}

SCOPE1 = 'https://spreadsheets.google.com/feeds'
SCOPE2 = 'https://www.googleapis.com/auth/drive'

SERVICE_ACCOUNT_JSON = 'credentials.json'

JSON_TEMPLATE = '..\\resume\\resume.empty.json'
JSON_OUTPUT = '..\\resume\\resume.json'

GSHEET_NAME = 'Resume'
GSHEET_MAIN_TAB_NAME = 'Main'
GSHEET_WORK_TAB_NAME = 'Work Experience'
GSHEET_EDUCATION_TAB_NAME = 'Education'


def main():
    credentials = get_credentials()
    gc = gspread.authorize(credentials)
    sh = gc.open(GSHEET_NAME)

    with open(JSON_TEMPLATE) as json_file:
        template = json.load(json_file)

    process_main_sheet(sh.worksheet(
        GSHEET_MAIN_TAB_NAME).get_all_values(), template)
    process_work_sheet(sh.worksheet(
        GSHEET_WORK_TAB_NAME).get_all_values(), template)
    process_education_sheet(sh.worksheet(
        GSHEET_EDUCATION_TAB_NAME).get_all_values(), template)

    with open(JSON_OUTPUT, "w", encoding='utf8') as json_file:
        json.dump(template, json_file, indent=2, ensure_ascii=False)


def process_main_sheet(values, template):
    for row in values:
        path = keymap[row[0]][0].split(".")
        content = keymap[row[0]][1]

        if (len(row) > 2 and row[2]):
            content = content.replace("{0}", row[1])
            content = content.replace("{1}", row[2])
        else:
            content = content.replace("{0}", row[1])

        try:
            content = json.loads(content)
        except:
            pass

        if len(path) > 1:
            if isinstance(template[path[0]][path[1]], list):
                template[path[0]][path[1]].append(content)
            else:
                template[path[0]][path[1]] = content
        else:
            if isinstance(template[path[0]], list):
                template[path[0]].append(content)
            else:
                template[path[0]] = content


def process_work_sheet(values, template):
    template["work"] = []
    header = values.pop(0)

    for row in values:
        #work_json = { "company": "", "position": "", "website": "", "startDate": "", "endDate": "", "summary": "", "highlights": [], "starred": None }
        work_json = {}
        work_json["company"] = row[3]
        work_json["position"] = row[2]
        work_json["website"] = ""
        work_json["startDate"] = row[0]
        work_json["endDate"] = row[1]
        work_json["location"] = row[6]
        work_json["summary"] = ""
        work_json["highlights"] = row[7].splitlines()
        work_json["starred"] = bool(strtobool(row[8]))

        template["work"].append(work_json)


def process_education_sheet(values, template):
    template["education"] = []
    header = values.pop(0)

    for row in values:
        #education_json = { "institution": "", "area": "", "studyType": "", "startDate": "", "endDate": "", "gpa": "", "courses": [] }
        education_json = {}
        education_json["institution"] = row[2]
        education_json["area"] = ""
        education_json["studyType"] = row[1]
        education_json["startDate"] = ""
        education_json["endDate"] = row[0]
        education_json["location"] = row[3]
        education_json["gpa"] = ""
        education_json["courses"] = ""

        template["education"].append(education_json)


def get_credentials():
    scopes = [SCOPE1, SCOPE2]
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_JSON, scopes=scopes)
    return credentials


if __name__ == "__main__":
    main()
