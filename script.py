import openpyxl
import os
import json
import layout
from github import Github
from dotenv import load_dotenv

# Load the .env values
load_dotenv()

# Enter your GitHub API token here
API_TOKEN = os.environ.get("GITHUB_TOKEN")

# Enter the repository owner and name here
ORGANISATION = os.environ.get("ORGANISATION")
REPO_NAME = os.environ.get("REPO_NAME")

# Enter the milestone number here
MILESTONE_NUMBER = int(os.environ.get("MILESTONE_NUMBER"))

# Connect to the GitHub API
g = Github(API_TOKEN)

# Get the organisation
org = g.get_organization(ORGANISATION)

# Get the repository
repo = org.get_repo(REPO_NAME)

# Get the open issues in the repository
open_issues = repo.get_issues(state="open")

# Get the closed issues in the repository
closed_issues = repo.get_issues(state="closed")

# Get the milestone
milestone = repo.get_milestone(MILESTONE_NUMBER)

# Open the JSON file
with open('config.json', 'r') as f:
    # Load the JSON data into a Python object
    config = json.load(f)

# Create a new Excel workbook
wb = openpyxl.Workbook()

# Create a new sheet for the timeline
sheet = wb.active
sheet.title = "Timeline " + str(milestone.number)

# Create the header row for the timeline
sheet["A1"] = "Assignees"
sheet["B1"] = "Issue"
for i in range(12):
    sheet[chr(ord('@') + i + 3) + '1'] = "Week " + str(i + 1)

# Iterate over the issues and add them to the timeline
row = 2
for issue in open_issues and closed_issues:
    # Skip issues not in the milestone
    if issue.milestone == None or issue.milestone.number != MILESTONE_NUMBER:
        continue

    # Get the issue information
    title = issue.title
    body = issue.body
    assignee = issue.assignee.login

    # Split the body to get the week assigned
    body = body.split("\r", 1)[0]
    body = body.split(" ")
    body.pop(0)

    # Add the issue information to the timeline
    sheet[f"A{row}"] = assignee
    sheet[f"B{row}"] = title
    
    # Fill the cells corresponding to the week aasigned
    for week in body:
        sheet[chr(ord('@') + int(week) + 2) + str(row)].fill = openpyxl.styles.PatternFill(fill_type="solid", fgColor=config['authors'][assignee])

    # Move to the next row
    row += 1

# Save the Excel workbook
wb.save("timeline.xlsx")

layout.apply_layout()
